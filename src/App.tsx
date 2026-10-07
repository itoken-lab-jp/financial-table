/**
 * 表の描画。中身（文字・色）は viewModel が決めたものを並べるだけ。
 *
 * - 表の上：年度、見る人が主と比較を選ぶメニュー、単位
 * - 行の見出しの列と、列の見出しは、スクロールしても残す（sticky）
 * - 囲み：集計先とその下の行を左の帯で包み、集計先の行は帯を横に伸ばす（下囲みなら L 字）
 * - 小計・段階の行・見出しは太字、最後の段階の行の上は二重線、小計と途中の段階の上は実線、ほかの行の間は薄い実線
 * - 組織の欄を入れたら、左に組織の段ごとの列を足し、組織の名前をその組織の行にまとめて縦にかける（行の見出しと一緒に左に残す）。
 *   組織の名前と、区分・中分類の小計の行の前のしるしで開き閉じする（見る人の選択として保存。visual.ts）
 */
import * as React from "react";
import { LayoutPlan, LayoutHeader } from "./hierarchy";
import { ValueParts } from "./valueParts";
import powerbi from "powerbi-visuals-api";

import { CopyImageButton } from "./shared/CopyImageButton";
import { htmlImage } from "./shared/htmlImage";
import { SCROLL_STARTS, useScrollStart } from "./shared/scrollStart";
import { OrgCell } from "./orgs";
import { GridPosition, SpannedCell, VisibleRows, buildGrid, foldOf, isClearKey, isMenuKey, isSelectKey, moveInGrid, moveOf, positionOf } from "./keyboard";
import { SelectTarget, cellActive, orgSelected, periodActive, plainCellActive, plainRowActive, rowActive, targetKey } from "./selection";
import { Band, Cell, EventMenu, SlotView, TableRow, TooltipItem, ViewModel, periodNameOf, textOn } from "./viewModel";
import { ComparePatch, PeriodPatch, PickPatch, VisualState } from "./visualState";
import { CompareDialog } from "./CompareDialog";
import { PeriodDialog } from "./PeriodDialog";
import { PickDialog } from "./PickDialog";
import { UNIT_PLACES, COLUMN_PADDINGS, ROW_HEIGHTS, DEFAULT_FONT_SIZE } from "./settings";

export interface AppProps {
    viewModel: ViewModel;
    viewport: powerbi.IViewport;
    /** 見る人が主・比較を選んだとき。無ければメニューを出さず名前だけ出す（操作できない場所） */
    onChooseMain?: (event: string) => void;
    /** 見る人が比較の列を選んだとき（上のバーの「比較」と期間の見出しの ▾）。無ければどちらも出さない（操作できない場所） */
    onChooseCompare?: (patch: ComparePatch) => void;
    /** 見る人が組織（道筋）・行（コード）を開き閉じしたとき。無ければしるしを出さない（操作できない場所） */
    onToggleOrg?: (path: string) => void;
    onToggleRow?: (code: string) => void;
    /** すべて開く・すべて閉じる（開いた組織と閉じた行をまとめて決める） */
    onSetFolds?: (patch: Partial<VisualState>) => void;
    /** 見る人が見せる科目を選んだとき。無ければ「科目を選ぶ」を出さない（操作できない場所） */
    onPick?: (parent: string, accounts: string[], patch: PickPatch) => void;
    /** 見る人が出す期間を選んだとき。無ければ「期間を選ぶ」を出さない（操作できない場所） */
    onPickPeriods?: (patch: PeriodPatch) => void;
    /** セルと列の見出しのツールチップ（Power BI の tooltipService）。無ければブラウザーのツールチップ（title）で出す */
    tooltip?: CellTooltip;
    /** 選択。今の選択と、押したとき・右クリック・空いた所を押したとき。無ければ選べない（操作できない場所） */
    selection?: SelectTarget[];
    onSelect?: (target: SelectTarget, multi: boolean) => void;
    onContextMenu?: (target: SelectTarget | null, x: number, y: number) => void;
    onClearSelection?: () => void;
    /** 「画像としてコピー」のボタン（書式ペインの「表示」で切れる）。無ければ出さない（操作できない場所）。browserMenu はブラウザーのメニューが出る所か */
    copyImage?: { browserMenu: boolean };
}

/**
 * 画像に写さない所：表の上のバーのメニューとボタン、ダイアログ、期間の見出しの ▾、お知らせ。題名・単位・表は写す
 *
 */
const COPY_SKIP = ".ft-events, .ft-folds, .ft-period-open, .ft-pick-open, .ft-pick-pop, .ft-colmenu-open, .ft-colmenu-spacer, .ft-notice";
// 画像は年度から下（上端のボタンの行は描かない）
const captureTable = (root: HTMLElement, background: string) => htmlImage(root.querySelector<HTMLElement>(".ft-sheet") ?? root, background, COPY_SKIP);

/** 押したときの扱い（行・セル・見出し・組織の名前に配る） */
interface Picking {
    selection: SelectTarget[];
    /** 期間のキー → 期間（セルごとに探さない） */
    periods: Map<string, ViewModel["periods"][number]>;
    onSelect?: (target: SelectTarget, multi: boolean) => void;
    onContextMenu?: (target: SelectTarget | null, x: number, y: number) => void;
    /** キーボードで今いるセル（Tab で止まる 1 か所）と、セルに移ったとき */
    focused: string;
    onFocusCell: (key: string) => void;
    /** キーボードでメニューを出したとき（すぐあとにブラウザーが起こす右クリックで、もう一度出さない） */
    keyMenu: KeyMenu;
}

/** キーボードでメニューを出した時刻。Windows のブラウザーは Shift+F10・メニュー キーのあとに右クリック（contextmenu）も起こすことがある */
interface KeyMenu {
    mark: () => void;
    recent: () => boolean;
}

/** 比較の列の ▾ のメニューを開いてから、ビジュアルからフォーカスが離れても閉じない時間（ms） */
const MENU_BLUR_GUARD_MS = 500;

/** キーボードでメニューを出したあと、ブラウザーの右クリックを無視する時間（ms） */
const KEY_MENU_GUARD_MS = 800;

/** 開き閉じのしるしのあるセル（Shift+→・← で開き閉じする） */
interface Fold {
    open: boolean;
    toggle: () => void;
}

/** キーボードで止まるセルの印。Tab で止まるのは今いるセルだけで、ほかは矢印で移る */
function focusProps(picking: Picking, key: string): React.HTMLAttributes<HTMLElement> & { "data-cell": string } {
    return {
        "data-cell": key,
        tabIndex: picking.focused === key ? 0 : -1,
        onFocus: () => picking.onFocusCell(key),
    };
}

/**
 * キーボードで出すメニューの位置：セルのうち、スクロールの中で見えている部分の真ん中（行をまたぐ組織の名前など、はみ出した大きなセルで、
 * ビジュアルの外に出さない）
 */
function menuPoint(element: HTMLElement): { x: number; y: number } {
    const rect = element.getBoundingClientRect();
    const box = element.closest(".ft-scroll")?.getBoundingClientRect() ?? rect;
    const left = Math.max(rect.left, box.left);
    const right = Math.min(rect.right, box.right);
    const top = Math.max(rect.top, box.top);
    const bottom = Math.min(rect.bottom, box.bottom);
    if (right <= left || bottom <= top) return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
    return { x: (left + right) / 2, y: (top + bottom) / 2 };
}

/** キーボードでメニューを出す（すぐあとのブラウザーの右クリックを無視する印を付けて） */
function showKeyMenu(picking: Picking, target: SelectTarget | null, element: HTMLElement): void {
    if (!picking.onContextMenu) return;
    picking.keyMenu.mark();
    const at = menuPoint(element);
    picking.onContextMenu(target, at.x, at.y);
}

/**
 * 押した・右クリックした・キーを押したときの手当て。target が無い所（選べない行）の右クリックは、Power BI の既定のメニュー。
 * Enter・Space で選び（Ctrl で足し引き）、Shift+F10・メニュー キーでメニュー、Shift+→・← で開き閉じ（fold があるセル）。
 * 押し続けたときの繰り返し（repeat）では選ばない・メニューを出さない（選ぶ・解くが交互に Power BI に送られる）
 */
function pickProps(picking: Picking, target: SelectTarget | null, fold?: Fold): React.HTMLAttributes<HTMLElement> {
    const key = target ? targetKey(target) : null;
    return {
        // 支援技術に、選べる所と選んだもの、開き閉じできる所を知らせる（選べない所には付けない）
        "aria-selected": key !== null && picking.onSelect ? picking.selection.some((t) => targetKey(t) === key) : undefined,
        "aria-expanded": fold ? fold.open : undefined,
        onClick: (e) => {
            if (!target || !picking.onSelect) return;
            e.stopPropagation();
            picking.onSelect(target, e.ctrlKey || e.metaKey);
        },
        onContextMenu: (e) => {
            if (!picking.onContextMenu) return;
            e.preventDefault();
            e.stopPropagation();
            if (picking.keyMenu.recent()) return;
            picking.onContextMenu(target, e.clientX, e.clientY);
        },
        onKeyDown: (e) => {
            const ctrl = e.ctrlKey || e.metaKey;
            if (isSelectKey(e.key) && !e.shiftKey && !e.altKey) {
                // 選べない所でも Space で画面が動かないように止める
                e.preventDefault();
                e.stopPropagation();
                if (target && picking.onSelect && !e.repeat) picking.onSelect(target, ctrl);
                return;
            }
            if (isMenuKey(e.key, e.shiftKey, ctrl, e.altKey) && picking.onContextMenu) {
                e.preventDefault();
                e.stopPropagation();
                if (!e.repeat) showKeyMenu(picking, target, e.currentTarget);
                return;
            }
            const want = foldOf(e.key, e.shiftKey, ctrl, e.altKey);
            if (want && fold) {
                e.preventDefault();
                e.stopPropagation();
                if ((want === "open") !== fold.open) fold.toggle();
            }
        },
    };
}

/**
 * キーボードで止まるセルの名前。描き直しても（選んだ、開き閉じした）同じセルに同じ名前が付くよう、行のキー・期間のキー・組織の道筋から作る
 */
const cellKey = {
    orgHead: (i: number) => `oh\u0001${i}`,
    corner: "corner",
    period: (periodKey: string) => `p\u0001${periodKey}`,
    column: (i: number) => `h\u0001${i}`,
    // 束ねる根（道筋が無い）の名前は、始まる行で見分ける
    // セグメントは 1 列で、道筋ごとにセルは 1 つ。開き閉じで閉じた表と計の表が入れ替わっても同じ名前にし、フォーカスを保つ
    org: (cell: OrgCell, rowKey: string) => `o\u0001${cell.identity ?? cell.path ?? `\u0002${rowKey}`}`,
    name: (rowKey: string) => `n\u0001${rowKey}`,
    value: (rowKey: string, i: number) => `v\u0001${rowKey}\u0001${i}`,
};

/** 表にあるセルの名前と、Tab で初めに止まるセル（最初の行の見出し。行が無ければ左上の角） */
export function cellKeysOf(viewModel: ViewModel): { keys: Set<string>; first: string } {
    const keys = new Set<string>([cellKey.corner]);
    viewModel.orgColumns.forEach((_, i) => keys.add(cellKey.orgHead(i)));
    for (const p of viewModel.periods) keys.add(cellKey.period(p.key));
    viewModel.columns.forEach((_, i) => keys.add(cellKey.column(i)));
    for (const row of viewModel.rows) {
        for (const cell of row.hierarchyCells ?? []) keys.add(cellKey.name(`hierarchy:${cell.row.key}`));
        for (const cell of row.orgCells ?? []) keys.add(cellKey.org(cell, row.key));
        keys.add(cellKey.name(row.key));
        row.cells.forEach((_, i) => keys.add(cellKey.value(row.key, i)));
    }
    const first = viewModel.rows.length > 0 ? cellKey.name(viewModel.rows[0].key) : cellKey.corner;
    return { keys, first };
}

/**
 * 表の中のキーボード：今いるセル（Tab で止まる 1 か所）と、矢印などで動く・Ctrl+Shift+C で解く・選べない所の Shift+F10。
 * 選ぶ・メニュー・開き閉じは、選べるセルの pickProps が先に受ける
 */
function useKeyboard(
    viewModel: ViewModel,
    onClearSelection: (() => void) | undefined,
    onContextMenu: ((target: SelectTarget | null, x: number, y: number) => void) | undefined,
    onOpenColumnMenu: ((cell: HTMLTableCellElement) => void) | undefined
): {
    focused: string;
    onFocusCell: (key: string) => void;
    keyMenu: KeyMenu;
    table: React.RefObject<HTMLTableElement | null>;
    tableProps: React.HTMLAttributes<HTMLTableElement>;
    /** 表の外（題名・メニュー）をマウスで押したとき。表からフォーカスが離れたことにする */
    onPointerOutside: () => void;
} {
    const [focusKey, setFocusKey] = React.useState<string | null>(null);
    const table = React.useRef<HTMLTableElement>(null);
    // 行・列をまたぐセルを通り抜けても、動き出した列（行）を保つ。マウスで押したら捨てる
    const anchor = React.useRef<GridPosition | null>(null);
    // フォーカスが表の中にあるか（最後に表の中へ移り、まだ文書の中のほかの所へ移っていない）
    const inside = React.useRef(false);
    // マウスで押してフォーカスが移るとき（見える所までのスクロールをしない）
    const pointer = React.useRef(false);
    const menuAt = React.useRef(0);
    const keyMenu = React.useMemo<KeyMenu>(
        () => ({
            mark: () => {
                menuAt.current = Date.now();
            },
            recent: () => Date.now() - menuAt.current < KEY_MENU_GUARD_MS,
        }),
        []
    );
    const { keys, first } = React.useMemo(() => cellKeysOf(viewModel), [viewModel]);
    const alternative = focusKey?.startsWith("n\u0001hierarchy:") ? focusKey.replace("n\u0001hierarchy:", "n\u0001") : focusKey?.replace("n\u0001", "n\u0001hierarchy:");
    const focused = focusKey !== null && keys.has(focusKey) ? focusKey : alternative && keys.has(alternative) ? alternative : first;

    // 描き直しでいたセルが消え（組織を開き閉じして行が替わった、フィルターで行が消えた）、フォーカスが行き場を失ったら、表の中に戻す。
    // 表の外へ移った・この文書（ビジュアル）がフォーカスを持っていないときは戻さない（ほかのビジュアルからフォーカスを奪わない）。
    // メニューから「含めない」を選んだときは、行が消える描き直しがメニュー（ビジュアルの外）にフォーカスのあるうちに来るので、
    // ビジュアルにフォーカスが戻ったとき（window の focus）にも戻す
    const focusedRef = React.useRef(focused);
    focusedRef.current = focused;
    const regain = React.useCallback(() => {
        const root = table.current;
        if (!root || !inside.current) return;
        const lost = document.activeElement === null || document.activeElement === document.body;
        if (!lost || !document.hasFocus()) return;
        root.querySelector<HTMLElement>(`[data-cell="${CSS.escape(focusedRef.current)}"]`)?.focus({ preventScroll: true });
    }, []);
    React.useLayoutEffect(regain);
    React.useEffect(() => {
        window.addEventListener("focus", regain);
        return () => window.removeEventListener("focus", regain);
    }, [regain]);

    const onKeyDown = (e: React.KeyboardEvent<HTMLTableElement>) => {
        const cell = (e.target as HTMLElement).closest<HTMLTableCellElement>("[data-cell]");
        if (!cell || !e.currentTarget.contains(cell)) return;
        const ctrl = e.ctrlKey || e.metaKey;
        if (isClearKey(e.key, ctrl, e.shiftKey)) {
            e.preventDefault();
            if (!e.repeat) onClearSelection?.();
            return;
        }
        // 選べない所（左上の角・組織の列の見出し）の Shift+F10 は、Power BI の既定のメニュー
        if (isMenuKey(e.key, e.shiftKey, ctrl, e.altKey) && onContextMenu) {
            e.preventDefault();
            if (e.repeat) return;
            keyMenu.mark();
            const at = menuPoint(cell);
            onContextMenu(null, at.x, at.y);
            return;
        }
        // 期間の見出し・比較の列の見出しで Alt+↓：▾ のメニューを開く（ドロップダウンを開く標準のキー）
        if (e.key === "ArrowDown" && e.altKey && !ctrl && cell.dataset.period && onOpenColumnMenu) {
            e.preventDefault();
            onOpenColumnMenu(cell);
            return;
        }
        // 選べない所の Enter・Space は何もしない（Space で表がスクロールしないように止める）
        if (isSelectKey(e.key)) {
            e.preventDefault();
            return;
        }
        const move = moveOf(e.key, ctrl, e.shiftKey, e.altKey);
        if (!move) return;
        e.preventDefault();
        const rows: Array<Array<SpannedCell<HTMLTableCellElement>>> = Array.from(e.currentTarget.rows, (tr) =>
            Array.from(tr.cells, (c) => ({ item: c, rowSpan: c.rowSpan, colSpan: c.colSpan }))
        );
        const grid = buildGrid(rows);
        const from = positionOf(grid, cell, anchor.current);
        const scroller = e.currentTarget.closest<HTMLElement>(".ft-scroll");
        const next = from && moveInGrid(grid, from, move, visibleRows(scroller, e.currentTarget, from.row));
        if (!next) return;
        anchor.current = next;
        const target = grid[next.row][next.col]!;
        target.focus({ preventScroll: true });
        if (!scroller) return;
        // 行の頭・表の頭へ動いたら、左と上に残る見出しへ移っても、表の頭までスクロールを戻す（固定した見出しのある表の Home と同じ）
        if (move === "rowStart" || move === "first") scroller.scrollLeft = 0;
        if (move === "colStart" || move === "first") scroller.scrollTop = 0;
        reveal(scroller, target);
    };

    const tableProps: React.HTMLAttributes<HTMLTableElement> = {
        onKeyDown,
        onMouseDown: () => {
            anchor.current = null;
            // 押して移るフォーカスは mousedown の直後に起きる。フォーカスの移らない押し方（いるセル・しるし）で印が残らないよう、すぐ倒す
            pointer.current = true;
            window.setTimeout(() => {
                pointer.current = false;
            }, 0);
        },
        onFocus: (e) => {
            inside.current = true;
            // Tab で入ったとき（マウスで押したときのほか）も、見出しの下に隠れていれば見える所まで動かす
            const cell = (e.target as HTMLElement).closest<HTMLTableCellElement>("[data-cell]");
            const scroller = e.currentTarget.closest<HTMLElement>(".ft-scroll");
            if (!pointer.current && cell && scroller) reveal(scroller, cell);
            pointer.current = false;
        },
        onBlur: (e) => {
            // 文書の中のほかの所（メニュー・ボタン）へ移った。移り先の無い blur（セルが消えた、ビジュアルの外へ出た）は中にいたままとみなす
            const next = e.relatedTarget as Node | null;
            if (next && !e.currentTarget.contains(next)) inside.current = false;
        },
    };

    return {
        focused,
        onFocusCell: setFocusKey,
        keyMenu,
        table,
        tableProps,
        onPointerOutside: () => {
            inside.current = false;
        },
    };
}

/**
 * 見えている体の行（升目の行の番号）。上と下の端がどちらも、列の見出しの下からスクロールの下の端までに入る行。
 * 1 行も入らなければ（行がスクロールより背が高い）今の行だけ
 */
function visibleRows(scroller: HTMLElement | null, table: HTMLTableElement, current: number): VisibleRows {
    if (!scroller) return { first: current, last: current };
    const box = scroller.getBoundingClientRect();
    const scale = scroller.offsetWidth > 0 ? box.width / scroller.offsetWidth : 1;
    const top = table.tHead?.getBoundingClientRect().bottom ?? box.top;
    const bottom = box.top + scroller.clientHeight * scale;
    let first = -1;
    let last = -1;
    Array.from(table.rows).forEach((tr, i) => {
        if (tr.parentElement === table.tHead) return;
        const rect = tr.getBoundingClientRect();
        if (rect.top >= top - 1 && rect.bottom <= bottom + 1) {
            if (first < 0) first = i;
            last = i;
        }
    });
    return first < 0 ? { first: current, last: current } : { first, last };
}

/**
 * キーボードで移ったセルが、スクロールの外や、上と左に残す見出しの下に隠れていたら、見える所までスクロールする。
 * ブラウザーに任せる（focus のスクロール）と、残す見出しの下に隠れたままになる。見出し（上・左に残るセル）はその向きには動かさない。
 * Power BI がビジュアルを拡大して描くと getBoundingClientRect は拡大した値を返すので、スクロールの量は倍率で割る
 */
function reveal(scroller: HTMLElement, cell: HTMLTableCellElement): void {
    const box = scroller.getBoundingClientRect();
    if (box.width <= 0 || scroller.offsetWidth <= 0) return;
    const scale = box.width / scroller.offsetWidth;
    const rect = cell.getBoundingClientRect();
    const tr = cell.parentElement as HTMLTableRowElement;
    const inHead = cell.closest("thead") !== null;
    // 左に残す列（組織の列と行の見出し）は style の left を持つ
    const stuckLeft = cell.style.left !== "";

    if (!inHead) {
        const head = cell.closest("table")?.querySelector("thead")?.getBoundingClientRect();
        const top = head ? head.bottom : box.top;
        const bottom = box.top + scroller.clientHeight * scale;
        const delta = overflowOf(rect.top, rect.bottom, top, bottom);
        if (delta !== 0) scroller.scrollTop += delta / scale;
    }
    if (!stuckLeft) {
        // 左に残る見出しの右端（体の行はその行の行の見出し、列の見出しは左上の角）
        const stuck = inHead
            ? cell.closest("table")?.querySelector<HTMLElement>("thead th.ft-corner:not(.ft-org-head)")
            : tr.querySelector<HTMLElement>(":scope > th.ft-name");
        const left = stuck ? stuck.getBoundingClientRect().right : box.left;
        const right = box.left + scroller.clientWidth * scale;
        const delta = overflowOf(rect.left, rect.right, left, right);
        if (delta !== 0) scroller.scrollLeft += delta / scale;
    }
}

/**
 * [start, end] を見える範囲 [min, max] に入れるための動き（画面の px、+ は先へ）。入らないほど大きいセルは、少しでも見えていれば動かさない
 * （行をまたぐ組織の名前に移るたびに、そのブロックの頭まで戻らないように）
 */
export function overflowOf(start: number, end: number, min: number, max: number): number {
    if (end - start > max - min) return end <= min || start >= max ? start - min : 0;
    if (start < min) return start - min;
    if (end > max) return end - max;
    return 0;
}

export interface CellTooltip {
    show: (items: TooltipItem[], clientX: number, clientY: number, move: boolean) => void;
    hide: () => void;
}

/** 1 段の字下げ（em） */
const INDENT_EM = 1;

export function App({
    viewModel,
    viewport,
    onChooseMain,
    onChooseCompare,
    onToggleOrg,
    onToggleRow,
    onSetFolds,
    onPick,
    onPickPeriods,
    tooltip,
    selection = [],
    onSelect,
    onContextMenu,
    onClearSelection,
    copyImage,
}: AppProps): React.JSX.Element {
    const { style } = viewModel;
    // 比較の列の ▾ のメニュー（開いていなければ null）
    const rootRef = React.useRef<HTMLDivElement>(null);
    const layoutFocus = React.useRef<{ key: string; table?: string } | null>(null);
    React.useLayoutEffect(() => {
        if (!viewModel.layout || !layoutFocus.current || !document.hasFocus() || document.activeElement !== document.body) return;
        const { key, table } = layoutFocus.current;
        const scope = table ? rootRef.current?.querySelector<HTMLElement>(`table[data-layout="${CSS.escape(table)}"]`) ?? rootRef.current : rootRef.current;
        const rowKey = key.startsWith("frame:") ? key.slice(6) : key.startsWith("n\u0001hierarchy:") ? key.slice(12) : key.startsWith("n\u0001") ? key.slice(2) : null;
        const candidates = [key, ...(rowKey ? [`frame:${rowKey}`, `n\u0001${rowKey}`, `n\u0001hierarchy:${rowKey}`] : [])];
        for (const candidate of candidates) {
            const cell = scope?.querySelector<HTMLElement>(`[data-cell="${CSS.escape(candidate)}"]`);
            if (cell) { cell.focus({ preventScroll: true }); break; }
        }
    }, [viewModel]);
    const [columnMenu, setColumnMenu] = React.useState<OpenColumnMenu | null>(null);
    // 見せる科目のダイアログ（開いていれば、開いたボタンのすぐ下の位置）
    const [pickAt, setPickAt] = React.useState<{ left: number; top: number } | null>(null);
    const pickButton = React.useRef<HTMLButtonElement>(null);
    const closePick = React.useCallback((returnFocus: boolean) => {
        setPickAt(null);
        if (returnFocus) pickButton.current?.focus();
    }, []);
    // 比較の列を選ぶダイアログ（全期間。同じ形）
    const [compareAt, setCompareAt] = React.useState<{ left: number; top: number } | null>(null);
    const compareButton = React.useRef<HTMLButtonElement>(null);
    const closeCompare = React.useCallback((returnFocus: boolean) => {
        setCompareAt(null);
        if (returnFocus) compareButton.current?.focus();
    }, []);
    // 期間を選ぶダイアログ（同じ形。開いていれば、開いたボタンのすぐ下の位置）。2 つのダイアログは同時に開かない
    const [periodAt, setPeriodAt] = React.useState<{ left: number; top: number } | null>(null);
    const periodButton = React.useRef<HTMLButtonElement>(null);
    const closePeriods = React.useCallback((returnFocus: boolean) => {
        setPeriodAt(null);
        if (returnFocus) periodButton.current?.focus();
    }, []);
    const below = (button: HTMLElement) => ({ left: button.offsetLeft, top: button.offsetTop + button.offsetHeight + 2 });
    // 閉じたあとにフォーカスを戻す見出し（描き直しのあとで戻す。状態の更新の中で focus を呼ばない）
    const refocus = React.useRef<HTMLElement | null>(null);
    const openColumnMenu = (cell: HTMLTableCellElement) => {
        const periodKey = cell.dataset.period ?? "";
        const slot = cell.dataset.slot !== undefined ? Number(cell.dataset.slot) : null;
        if (!rootRef.current || !periodKey) return;
        tooltip?.hide();
        setColumnMenu(columnMenuAt(rootRef.current, cell, slot, periodKey));
    };
    const toggleColumnMenu = (cell: HTMLTableCellElement, periodKey: string) => {
        if (columnMenu?.periodKey === periodKey) {
            setColumnMenu(null);
            return;
        }
        openColumnMenu(cell);
    };
    const menuRef = React.useRef(columnMenu);
    menuRef.current = columnMenu;
    const closeColumnMenu = React.useCallback((returnFocus: boolean) => {
        const open = menuRef.current;
        if (!open) return;
        if (returnFocus) refocus.current = open.returnTo;
        setColumnMenu(null);
    }, []);
    React.useLayoutEffect(() => {
        const target = refocus.current;
        refocus.current = null;
        target?.focus({ preventScroll: true });
    });
    // ▾ のダイアログを開いたまま表を組み直した（列を足し引きした）ら、開いた見出しを探し直して、位置と閉じたときのフォーカスの戻り先を
    // 当て直す（見出しが動いてダイアログが離れたり、消えた見出しにフォーカスを戻そうとしたりしない）。比較の列の見出しが消えたら期間の見出し
    React.useLayoutEffect(() => {
        const open = menuRef.current;
        const root = rootRef.current;
        if (!open || !root) return;
        const blockKey = open.returnTo.closest<HTMLTableElement>("table")?.dataset.layout;
        const scope = blockKey ? root.querySelector<HTMLElement>(`table[data-layout="${CSS.escape(blockKey)}"]`) ?? root : root;
        const period = scope.querySelector<HTMLElement>(`th.ft-period[data-period="${CSS.escape(open.periodKey)}"]`);
        const own = open.slot !== null ? scope.querySelector<HTMLElement>(`th[data-period="${CSS.escape(open.periodKey)}"][data-slot="${open.slot}"]`) : null;
        const cell = own ?? period;
        if (!cell) return;
        const next = columnMenuAt(root, cell, own ? open.slot : null, open.periodKey);
        // 開いたときにフォーカスを置いた列は変えない（見出しが消えても、ダイアログの中のフォーカスを動かさない）
        if (next.left !== open.left || next.top !== open.top || next.returnTo !== open.returnTo) setColumnMenu({ ...next, slot: open.slot });
    }, [viewModel]);
    // 単位をバーに置くときは、バーを表の幅までで折り返し、単位を表の右の端にそろえる（表がビジュアルより狭いと、単位が表から離れて
    // 右の端に浮いた）。表の幅はレイアウトの px（offsetWidth）。狭すぎる表でも題名とメニューが入る幅は残す。
    // 単位を左上の角に置くときは、バーはビジュアルの幅いっぱいに使う（表の幅で折り返すと、狭い表でボタンが 1 つずつ行に残った）
    const [tableWidth, setTableWidth] = React.useState(0);
    // スクロールすると、表の左と上の枠線は流れて消えるので、スクロールの箱の縁に線を引く（visual.less の .ft-scrolled-x・-y）
    const [scrolled, setScrolled] = React.useState({ x: false, y: false });
    React.useLayoutEffect(() => {
        const content = scrollStart.ref.current?.firstElementChild as HTMLElement | null | undefined;
        const width = keyboard.table.current?.offsetWidth ?? content?.offsetWidth ?? 0;
        if (width !== tableWidth) setTableWidth(width);
    });
    // バーの「比較」のダイアログも、選ぶたびにボタンの名前が伸び縮みしてバーが折り返すと、ボタンの位置が変わる。表を組み直したら
    // ボタンの下に当て直す
    React.useLayoutEffect(() => {
        const button = compareButton.current;
        if (!button) return;
        setCompareAt((open) => {
            if (!open) return open;
            const next = below(button);
            return next.left === open.left && next.top === open.top ? open : next;
        });
    }, [viewModel]);
    // 年度・期間の並びや大きさが変わったら閉じる（見出しから離れた位置に残さない。開いた期間が消えたあと戻って勝手に開かない）。
    // ▾ のダイアログで選ぶたびに表は組み直すので、表の形（列の数）では閉じない
    const periodShape = JSON.stringify([viewModel.yearKey, viewModel.periods.map((p) => p.key)]);
    React.useEffect(() => {
        if (menuRef.current) setColumnMenu(null);
    }, [periodShape, viewport.width, viewport.height]);
    // 見せる科目・期間・比較のダイアログは、大きさが変わったら閉じる（位置は開いたときのボタンの位置で、バーの折り返しが変わるとずれる）。
    // 選ぶたびに表の形は変わるので、表の形では閉じない
    React.useEffect(() => {
        setPickAt(null);
        setPeriodAt(null);
        setCompareAt(null);
    }, [viewport.width, viewport.height]);
    // ほかのビジュアル・キャンバスを押すと、この文書（ビジュアル）からフォーカスが離れる。そのとき閉じる。
    // 開いてすぐのものは無視する（ビジュアルを押したあとに Power BI がフォーカスを取り返すと、開いた直後に閉じるおそれがある）
    React.useEffect(() => {
        if (!columnMenu) return;
        const openedAt = Date.now();
        const blur = () => {
            if (Date.now() - openedAt > MENU_BLUR_GUARD_MS) setColumnMenu(null);
        };
        window.addEventListener("blur", blur);
        return () => window.removeEventListener("blur", blur);
    }, [columnMenu]);
    // 比較の列の ▾ は期間の見出しに 1 つずつ（比較対象の候補があって、操作できるとき。比較の列が無い期間にも出す：足すため）
    const hasColumnMenu = onChooseCompare !== undefined && (viewModel.menu?.compareItems.length ?? 0) > 0;
    const keyboard = useKeyboard(viewModel, onClearSelection, onContextMenu, hasColumnMenu ? openColumnMenu : undefined);
    React.useLayoutEffect(() => {
        const cells = Array.from(rootRef.current?.querySelectorAll<HTMLElement>(".ft-value[data-value-column]") ?? []);
        const widths = new Map<string, number>();
        for (const cell of cells) {
            const width = Math.max(0, ...Array.from(cell.querySelectorAll<HTMLElement>(".ft-number-tail-content"), tail => tail.offsetWidth));
            const key = cell.dataset.valueColumn!;
            widths.set(key, Math.max(widths.get(key) ?? 0, width));
        }
        // 「数値」カードの「単位の字の幅をそろえる」を切ると、そろえない（単位の字の無い行の右に空きを出さない）
        for (const cell of cells) {
            if (viewModel.style.alignTails) cell.style.setProperty("--ft-tail-width", `${Math.ceil(widths.get(cell.dataset.valueColumn!) ?? 0)}px`);
            else cell.style.removeProperty("--ft-tail-width");
        }
        // 同じ構成の分割表は、名前・数値の列幅を共有する。テキストの実測値を使い、縮小しない。
        const tables = Array.from(rootRef.current?.querySelectorAll<HTMLTableElement>(".ft-layout-table") ?? []);
        // 前回の固定位置を外してから自然なセル位置を測る。ブラウザは結合セルや罫線の分だけ列幅を広げることがある。
        for (const table of tables) for (const cell of Array.from(table.querySelectorAll<HTMLElement>("th[data-width-key]"))) {
            cell.style.position = "relative";
            cell.style.left = "";
        }
        const columnWidths = new Map<string, number>();
        const scale = rootRef.current && rootRef.current.offsetWidth ? rootRef.current.getBoundingClientRect().width / rootRef.current.offsetWidth : 1;
        for (const table of tables) for (const cell of Array.from(table.querySelectorAll<HTMLTableCellElement>("[data-width-key]"))) {
            if (cell.colSpan !== 1) continue;
            const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT);
            const rects: DOMRect[] = [];
            while (walker.nextNode()) {
                if (!walker.currentNode.textContent?.trim()) continue;
                const range = document.createRange();
                range.selectNodeContents(walker.currentNode);
                rects.push(range.getBoundingClientRect());
            }
            for (const slot of Array.from(cell.querySelectorAll(".ft-toggle-slot, .ft-number-tail"))) rects.push(slot.getBoundingClientRect());
            const style = window.getComputedStyle(cell);
            const contentWidth = rects.length ? Math.max(...rects.map(r => r.right)) - Math.min(...rects.map(r => r.left)) : 0;
            const width = contentWidth / scale + parseFloat(style.paddingLeft) + parseFloat(style.paddingRight) + 2;
            const key = `${table.dataset.columnShape}:${cell.dataset.widthKey}`;
            columnWidths.set(key, Math.max(columnWidths.get(key) ?? 0, width));
        }
        // 帯の名称と単位も表幅に含める。長い単位で帯だけが表より広がらないようにする。
        for (const table of tables) {
            const band = table.closest(".ft-layout-split")?.querySelector<HTMLElement>(":scope > .ft-layout-label");
            if (!band?.querySelector(".ft-layout-unit")) continue;
            const style = getComputedStyle(band);
            const required = Array.from(band.children).reduce((sum, e) => sum + e.getBoundingClientRect().width / scale, 0)
                + parseFloat(style.columnGap) + parseFloat(style.paddingLeft) + parseFloat(style.paddingRight);
            const current = Array.from(table.querySelectorAll<HTMLTableColElement>("col")).reduce((sum, col) =>
                sum + Math.ceil(columnWidths.get(`${table.dataset.columnShape}:${col.dataset.widthKey}`) ?? 80), 0);
            if (required > current) {
                const key = `${table.dataset.columnShape}:n0`;
                columnWidths.set(key, (columnWidths.get(key) ?? 80) + required - current);
            }
        }
        for (const table of tables) {
            let width = 0;
            for (const col of Array.from(table.querySelectorAll<HTMLTableColElement>("col"))) {
                const size = Math.ceil(columnWidths.get(`${table.dataset.columnShape}:${col.dataset.widthKey}`) ?? 80);
                col.style.width = `${size}px`;
                width += size;
            }
            table.style.width = `${width}px`;
            table.style.minWidth = `${width}px`;
            const offsets = new Map<string, number>();
            const columns = Array.from(table.querySelectorAll<HTMLTableColElement>("col"));
            const origin = columns[0]?.getBoundingClientRect().left ?? 0;
            let namesWidth = 0;
            for (const col of columns) {
                const rect = col.getBoundingClientRect();
                offsets.set(col.dataset.widthKey!, (rect.left - origin) / scale);
                if (!col.dataset.widthKey?.startsWith("v")) namesWidth += rect.width / scale;
            }
            const sticky = namesWidth < viewport.width * 0.65;
            for (const cell of Array.from(table.querySelectorAll<HTMLElement>("th[data-width-key]"))) {
                if (cell.dataset.widthKey?.startsWith("v")) continue;
                // 1px 左（スクロールの箱の縁の内側）で止める。Desktop の端数の倍率で、縁の線とのあいだに隙間ができて下の字が透けた（visual.less の .ft-scrolled-x）
                cell.style.left = sticky ? `${(offsets.get(cell.dataset.widthKey!) ?? 0) - 1}px` : "";
                cell.style.position = sticky ? "sticky" : "relative";
            }
        }
    }, [viewModel, viewport.width, viewport.height]);
    const picking: Picking = {
        selection,
        onSelect,
        onContextMenu,
        periods: new Map(viewModel.periods.map((p) => [p.key, p])),
        focused: keyboard.focused,
        onFocusCell: keyboard.onFocusCell,
        keyMenu: keyboard.keyMenu,
    };
    const lefts = useStickyLefts(viewModel.orgColumns.length, viewport.width);
    // 列がはみ出すときの最初の位置（左端か右端）。行は開いたときの上から。右端なら、年度・列の並びが変わったら当て直す
    const scrollStart = useScrollStart<HTMLDivElement>({
        axis: "x",
        target: viewModel.scrollStart === SCROLL_STARTS.end ? "end" : "start",
        // 期間の並びで決める。見る人が比較の列を出し入れしても当て直さない（右端から始める表で ▾ のチェックのたびにスクロールが動き、
        // メニューが閉じた）
        shape: periodShape,
    });
    const rootStyle = {
        width: viewport.width,
        height: viewport.height,
        fontFamily: style.fontFamily,
        // ボタン・メニュー・ダイアログの文字は既定の大きさのまま。書式の「年度・単位の文字サイズ」は年度と単位だけに効く
        fontSize: `${DEFAULT_FONT_SIZE}pt`,
        "--ft-size-heading": `${style.fontSize}pt`,
        ...(style.headingColor ? { "--ft-heading-color": style.headingColor } : {}),
        // 表の中の文字サイズ・見出しの背景・罫線の色は visual.less が変数で受ける
        ...Object.fromEntries(Object.entries(style.sizes).map(([key, size]) => [`--ft-size-${key}`, `${size}pt`])),
        "--ft-size-unit": `${style.fontSize}pt`,
        "--ft-row-pad": ROW_HEIGHTS[style.rowHeight].pad,
        "--ft-col-pad": COLUMN_PADDINGS[style.columnPadding],
        "--ft-row-line": ROW_HEIGHTS[style.rowHeight].line,
        "--ft-row-two": String(ROW_HEIGHTS[style.rowHeight].two),
        ...Object.fromEntries(Object.entries(style.lines).map(([key, color]) => [`--ft-line-${key}`, color])),
        "--ft-head-bg": style.headBackground,
        "--ft-head-text": style.headText,
        "--ft-head-muted": style.headMuted,
        // 字・副文字・地はテーマの色
        "--ft-text": style.text,
        "--ft-muted": style.muted,
        "--ft-bg": style.background,
    } as React.CSSProperties;
    // 期間の区切り：期間の最初の列（表の最初の期間は除く）の左に縦の線を引く
    const periodStarts = new Set(
        style.periodLines ? viewModel.columns.flatMap((c, i) => (i > 0 && viewModel.columns[i - 1].periodKey !== c.periodKey ? [i] : [])) : []
    );
    const unitInCorner = style.unitPlace === UNIT_PLACES.corner;
    // 表が 1 つ（分けていない）なら、単位は年度と同じ行の右に置く。分けた表は、表ごと（見出しの帯）に置く
    // 年度を左上の角に置くときは、表の上に出さない（行の名前の列が無い表は、角が無いので表の上）
    const headingTitle = style.titleInCorner && !viewModel.hideNames ? "" : viewModel.title;
    const headingUnit = !unitInCorner && viewModel.layout?.kind === "table" && style.unitPlace === UNIT_PLACES.right ? viewModel.unitCaption : "";

    if (viewModel.landing.length > 0) {
        return (
            <div className="ft-root ft-landing" style={rootStyle}>
                <p>財務諸表の表：次のフィールドを入れてください</p>
                <ul>
                    {viewModel.landing.map((m) => (
                        <li key={m}>{m}</li>
                    ))}
                </ul>
                <p className="ft-landing-note">値のメジャーを値の欄に入れると、合計の 1 行が出ます。行の欄に区分 → 中分類 → 科目名のような段を上から入れると行に分かれ、月の欄で期間の列、貸方フラグで収益と費用の向きがそろいます。区分の小計はビジュアルが組み、段階利益・合計・率は書式ペインの「計算行」で足します</p>
            </div>
        );
    }

    return (
        <div
            ref={rootRef}
            className={`ft-root${viewModel.style.periodLines ? " ft-period-lines" : ""}${viewModel.style.aggregateDouble ? " ft-aggregate-double" : ""}${viewModel.style.boldAggregates ? " ft-bold-aggregates" : ""}${viewModel.style.breakdown.muted ? " ft-breakdown-muted" : ""}${Object.entries(viewModel.style.bold).filter(([, m]) => m !== "auto").map(([k, m]) => ` ft-bold-${k}-${m}`).join("")}`}
            style={rootStyle}
            // フォーカスの枠はキーボードで動かしたときだけ出す（Desktop ではマウスで押しても :focus-visible が効き、押したセルに黒い枠が出た）
            onKeyDownCapture={() => rootRef.current?.classList.add("ft-keyboard")}
            onPointerDownCapture={() => rootRef.current?.classList.remove("ft-keyboard")}
            onFocusCapture={e => {
                const target = e.target as HTMLElement;
                const key = target.closest<HTMLElement>("[data-cell]")?.dataset.cell;
                layoutFocus.current = target.closest(".ft-scroll") && key ? { key, table: target.closest<HTMLElement>("table[data-layout]")?.dataset.layout } : null;
            }}
            // 表の中の押せる所のほか（題名・左上の角・表の外の空いた所）を押すと選択を解く。メニューとボタンは除く
            onClick={(e) => {
                if ((e.target as HTMLElement).closest("select, button")) return;
                onClearSelection?.();
            }}
            // 押せる所のほかの右クリックも、Power BI のメニュー（認定の要件）。キーボードで出した直後のものは無視する
            onContextMenu={(e) => {
                if (!onContextMenu) return;
                e.preventDefault();
                if (keyboard.keyMenu.recent()) return;
                onContextMenu(null, e.clientX, e.clientY);
            }}
            // 表の外（題名・メニュー・ボタン）を押したら、表からフォーカスが離れたことにする（描き直しでフォーカスを表に戻さない）
            onMouseDown={(e) => {
                if (!keyboard.table.current?.contains(e.target as Node)) keyboard.onPointerOutside();
            }}
        >
            {/*
              ボタン（開き閉じ・期間を選ぶ・行を選ぶ）は上端の行に分け、年度・主・比較・単位は表のすぐ上に置く。表を資料に貼る人が、2 行目から下を
              切り取れば表だけになる（「表で使う文字（年度とか）は表と近づけて、ボタンは上端に行を分けて、画像を切り取りやすいような
              レイアウトに」）。画像としてコピーも ft-sheet（年度から下）だけを描く
            */}
            {/* 上のボタンと、基準・比較のメニューを 1 行に並べる（画像としてコピーする範囲（ft-sheet）の外） */}
            {(viewModel.folds || viewModel.periodPicker || viewModel.picks || viewModel.menu) && (
                <div className="ft-toolbar">
                    {viewModel.folds && onSetFolds && <Folds folds={viewModel.folds} onSetFolds={onSetFolds} />}
                    {viewModel.periodPicker && onPickPeriods && (
                        <button
                            type="button"
                            ref={periodButton}
                            className="ft-period-open"
                            aria-haspopup="dialog"
                            aria-expanded={periodAt !== null}
                            onClick={(e) => {
                                const button = e.currentTarget;
                                setPickAt(null);
                                setCompareAt(null);
                                setPeriodAt((open) => (open !== null ? null : below(button)));
                            }}
                        >
                            期間を選ぶ
                        </button>
                    )}
                    {viewModel.picks && onPick && (
                        <button
                            type="button"
                            ref={pickButton}
                            className="ft-pick-open"
                            aria-haspopup="dialog"
                            aria-expanded={pickAt !== null}
                            onClick={(e) => {
                                // ボタンのすぐ下（ビジュアルの枠の中に収めるのは PickDialog）
                                const button = e.currentTarget;
                                setPeriodAt(null);
                                setCompareAt(null);
                                setPickAt((open) => (open !== null ? null : below(button)));
                            }}
                        >
                            行を選ぶ
                        </button>
                    )}
                    {viewModel.menu && <Menu menu={viewModel.menu} onChooseMain={onChooseMain} />}
                    {viewModel.menu && (
                        <CompareButton
                            menu={viewModel.menu}
                            buttonRef={compareButton}
                            open={compareAt !== null}
                            onToggle={
                                onChooseCompare && viewModel.menu.compareItems.length > 0
                                    ? (button) => {
                                          setPickAt(null);
                                          setPeriodAt(null);
                                          setCompareAt((open) => (open !== null ? null : below(button)));
                                      }
                                    : undefined
                            }
                        />
                    )}
                </div>
            )}
            <div className="ft-sheet">
                {!unitInCorner && !viewModel.layout && viewModel.unitCaption && <div className="ft-caption" style={!unitInCorner && tableWidth > 0 ? { maxWidth: Math.max(tableWidth, CAPTION_MIN_WIDTH) } : undefined}>
                    <span className="ft-unit">{viewModel.unitCaption}</span>
                </div>}
                {(headingTitle || headingUnit) && (
                    // 年度と単位は表のすぐ上の 1 行に、年度を左・単位を右に置く（表を切り取ると、表の題として一緒に入る）
                    <div className="ft-heading" style={tableWidth > 0 ? { maxWidth: Math.max(tableWidth, CAPTION_MIN_WIDTH) } : undefined}>
                        <span className="ft-title">{headingTitle}</span>
                        {headingUnit && <span className="ft-unit">{headingUnit}</span>}
                    </div>
                )}
                {viewModel.notice && <div className="ft-notice">{viewModel.notice}</div>}
                <div
                    className={`ft-scroll${viewModel.layout ? " ft-layout-scroll" : ""}${scrolled.x ? " ft-scrolled-x" : ""}${scrolled.y ? " ft-scrolled-y" : ""}`}
                    ref={scrollStart.ref}
                    onScroll={(e) => {
                        scrollStart.onScroll();
                        const x = e.currentTarget.scrollLeft > 0;
                        const y = e.currentTarget.scrollTop > 0;
                        if (x !== scrolled.x || y !== scrolled.y) setScrolled({ x, y });
                        // 表を動かしたら、見出しから離れた ▾ のメニューを閉じる
                        if (columnMenu) closeColumnMenu(false);
                    }}
                >
                    {viewModel.layout ? <LayoutView plan={viewModel.layout} unitInBand={!!headingUnit} viewModel={viewModel} viewport={viewport} tooltip={tooltip}
                        selection={selection} onSelect={onSelect} onContextMenu={onContextMenu} onClearSelection={onClearSelection}
                        onToggleOrg={onToggleOrg} onToggleRow={onToggleRow} onChooseCompare={onChooseCompare}
                        openMenu={hasColumnMenu ? openColumnMenu : undefined} toggleMenu={toggleColumnMenu} activeMenu={columnMenu ?? undefined} /> : <table
                        // 2 段の比較のセルがある表では、どの行も 2 段の高さにそろえる（visual.less の .ft-two-lines）
                        className={`ft-table${viewModel.twoLines ? " ft-two-lines" : ""}`}
                        role="grid"
                        aria-label={viewModel.title || "財務諸表"}
                        aria-multiselectable={onSelect ? true : undefined}
                        ref={keyboard.table}
                        {...keyboard.tableProps}
                    >
                        <thead ref={lefts.head}>
                            <tr>
                                {viewModel.orgColumns.map((name, i) => (
                                    <th
                                        key={`org${i}`}
                                        ref={lefts.refs[i]}
                                        className={`ft-corner ft-org-head${lefts.stick ? "" : " ft-org-free"}`}
                                        rowSpan={2}
                                        style={{ left: lefts.at(i) }}
                                        {...focusProps(picking, cellKey.orgHead(i))}
                                    >
                                        {name}
                                    </th>
                                ))}
                                {!viewModel.hideNames && (
                                    <th className="ft-corner" rowSpan={2} style={{ left: lefts.at(viewModel.orgColumns.length) }} {...focusProps(picking, cellKey.corner)}>
                                        {unitInCorner && <span className="ft-corner-unit">{viewModel.unitCaption}</span>}
                                    </th>
                                )}
                                {viewModel.periods.map((p, i) => (
                                    <th
                                        key={p.key}
                                        colSpan={p.span}
                                        className={`ft-period ft-period-${p.kind}${i > 0 && style.periodLines ? " ft-period-start" : ""}${onSelect ? " ft-pickable" : ""}${periodActive(selection, p.key) ? "" : " ft-dim"}`}
                                        {...focusProps(picking, cellKey.period(p.key))}
                                        {...pickProps(picking, { kind: "period", periodKey: p.key, months: p.months })}
                                        {...(hasColumnMenu ? { "data-period": p.key, "aria-haspopup": "dialog" as const } : {})}
                                    >
                                        {hasColumnMenu ? (
                                            <span className="ft-period-label">
                                                {/* ▾ と同じ幅の空きを左にも置き、期間の名前を真ん中に保つ */}
                                                <span className="ft-colmenu-spacer" aria-hidden="true" />
                                                {p.label}
                                                <PeriodMenuButton
                                                    period={p}
                                                    viewModel={viewModel}
                                                    open={columnMenu?.periodKey === p.key}
                                                    onToggle={(cell) => toggleColumnMenu(cell, p.key)}
                                                />
                                            </span>
                                        ) : (
                                            p.label
                                        )}
                                        {i > 0 && style.periodLines && <span className="ft-period-line" />}
                                    </th>
                                ))}
                            </tr>
                            <tr>
                                {viewModel.columns.map((c, i) => (
                                    <th
                                        key={`${c.periodKey}:${c.slot ?? "main"}`}
                                        className={`ft-colhead ft-col-${c.kind}${periodStarts.has(i) ? " ft-period-start" : ""}${c.tooltip ? " ft-colhead-mixed" : ""}${onSelect ? " ft-pickable" : ""}${periodActive(selection, c.periodKey) ? "" : " ft-dim"}`}
                                        {...focusProps(picking, cellKey.column(i))}
                                        {...tooltipProps(c.tooltip, tooltip)}
                                        {...pickProps(picking, periodTarget(viewModel, c.periodKey))}
                                        {...(c.slot !== undefined && hasColumnMenu ? { "data-slot": c.slot, "data-period": c.periodKey, "aria-haspopup": "dialog" as const } : {})}
                                    >
                                        <div>{c.header}</div>
                                        {/* 1 行の見出し（主、比較の差・比・率・値だけの列）は 2 行目を置かずに縦横の真ん中に置く */}
                                        {c.sub && <div className="ft-colhead-sub">{c.sub}</div>}
                                        {periodStarts.has(i) && <span className="ft-period-line" />}
                                    </th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {viewModel.rows.map((row, index) => (
                                <Row
                                    key={row.key}
                                    row={row}
                                    periodStarts={periodStarts}
                                    boxEdge={index > 0 && isBoxEdge(viewModel.rows[index - 1], row)}
                                    viewModel={viewModel}
                                    tooltip={tooltip}
                                    left={lefts.at}
                                    top={lefts.top}
                                    onToggleOrg={onToggleOrg}
                                    onToggleRow={onToggleRow}
                                    picking={picking}
                                />
                            ))}
                        </tbody>
                    </table>}
                </div>
            </div>
            {pickAt !== null && viewModel.picks && onPick && (
                <PickDialog
                    parents={viewModel.picks.parents}
                    current={viewModel.picks.current}
                    left={pickAt.left}
                    top={pickAt.top}
                    onPick={onPick}
                    onClose={closePick}
                />
            )}
            {compareAt !== null && viewModel.menu && onChooseCompare && (
                <CompareDialog
                    title="比較の列（全期間）"
                    items={viewModel.menu.compareItems}
                    slots={viewModel.menu.compares.all}
                    period={null}
                    left={compareAt.left}
                    top={compareAt.top}
                    opener=".ft-compare-open"
                    onChoose={onChooseCompare}
                    onClose={closeCompare}
                />
            )}
            {periodAt !== null && viewModel.periodPicker && onPickPeriods && (
                <PeriodDialog
                    years={viewModel.periodPicker.years}
                    current={viewModel.periodPicker.current}
                    total={viewModel.periodPicker.total}
                    picked={viewModel.periodPicker.picked}
                    left={periodAt.left}
                    top={periodAt.top}
                    onPick={onPickPeriods}
                    onClose={closePeriods}
                />
            )}
            {columnMenu && onChooseCompare && viewModel.menu && (
                <PeriodCompareDialog open={columnMenu} viewModel={viewModel} menu={viewModel.menu} onChoose={onChooseCompare} onClose={closeColumnMenu} />
            )}
            {copyImage && viewModel.copyButton && (
                <CopyImageButton
                    alt={viewModel.title || "財務諸表の表"}
                    // 描いたものが変わる値（表・大きさ・選択）。スクロールはボタンが自分で拾う
                    stamp={[viewModel, viewport.width, viewport.height, selection]}
                    browserMenu={copyImage.browserMenu}
                    capture={captureTable}
                />
            )}
        </div>
    );
}

/** 主のメニュー（比較の列は「比較」のボタンと期間の見出しの ▾）。選べないとき（ダッシュボードのタイルなど）は名前だけ */
function Menu({
    menu,
    onChooseMain,
}: {
    menu: EventMenu;
    onChooseMain?: (e: string) => void;
}): React.JSX.Element {
    const label = (value: string) => menu.mainItems.find((i) => i.value === value)?.label ?? value;
    return (
        <span className="ft-events">
            <span>基準</span>
            {onChooseMain && menu.mainItems.length > 1 ? (
                <select className="ft-select" aria-label="基準" value={menu.main} onChange={(e) => onChooseMain(e.target.value)}>
                    {menu.mainItems.map((item) => (
                        <option key={item.value} value={item.value}>
                            {item.label}
                        </option>
                    ))}
                </select>
            ) : (
                <span className="ft-event">{label(menu.main)}</span>
            )}
        </span>
    );
}

/** 列 1〜3 のうち表に出す列のシナリオの名前（同じシナリオは 1 度）。無ければ「なし」 */
function slotNames(slots: SlotView[]): string {
    const names = Array.from(new Set(slots.filter((slot) => slot.available).map((slot) => slot.label)));
    return names.length > 0 ? names.join("・") : "なし";
}

/**
 * 表の上のバーの「比較」：全期間の比較の列のシナリオの名前。押すと比較の列を選ぶダイアログ（CompareDialog）。操作できない場所では名前だけ。
 * 比較の列が無ければ「なし」
 */
function CompareButton({
    menu,
    buttonRef,
    open,
    onToggle,
}: {
    menu: EventMenu;
    buttonRef: React.RefObject<HTMLButtonElement | null>;
    open: boolean;
    onToggle?: (button: HTMLElement) => void;
}): React.JSX.Element {
    const text = slotNames(menu.compares.all);
    return (
        <span className="ft-events">
            <span>比較</span>
            {onToggle ? (
                <button
                    type="button"
                    ref={buttonRef}
                    className="ft-compare-open"
                    aria-haspopup="dialog"
                    aria-expanded={open}
                    aria-label={`比較の列（全期間）：${text}`}
                    onClick={(e) => onToggle(e.currentTarget)}
                >
                    {text} ▾
                </button>
            ) : (
                <span className="ft-event">{text}</span>
            )}
        </span>
    );
}

/** 開いている期間の ▾ のダイアログ：どの比較の列・どの期間の見出しから開いたか、置く位置（ルートの中の px）、閉じたときに戻すフォーカス */
interface OpenColumnMenu {
    /** 比較の列の見出しから開いたときの列の番号（その列の箱へフォーカスを移す）。期間の見出しからなら null */
    slot: number | null;
    periodKey: string;
    left: number;
    top: number;
    returnTo: HTMLElement | null;
}

/**
 * 期間の見出しの ▾。期間の名前の右に置き、押すとその期間の比較の列のダイアログを開く。この期間だけ比較の列を変えていても塗らない
 * （見出しの名前で分かる。塗ると鬱陶しかった）。説明にだけ添える。
 * Tab で止めない（キーボードでは見出しで Alt+↓）。押しても見出しの選択（月で絞る）にしない
 */
function PeriodMenuButton({
    period,
    viewModel,
    open,
    onToggle,
}: {
    period: ViewModel["periods"][number];
    viewModel: ViewModel;
    open: boolean;
    onToggle: (cell: HTMLTableCellElement) => void;
}): React.JSX.Element {
    const changed = period.own;
    const label = `${period.label}の比較の列を選ぶ${changed ? "（この期間だけ変えている）" : ""}`;
    return (
        <button
            type="button"
            className="ft-colmenu-open"
            aria-label={label}
            aria-expanded={open}
            title={changed ? `${period.label}の比較の列（この期間だけ変えている）` : `${period.label}の比較の列`}
            tabIndex={-1}
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
                e.stopPropagation();
                const cell = (e.currentTarget as HTMLElement).closest<HTMLTableCellElement>("th");
                if (cell) onToggle(cell);
            }}
        >
            ▾
        </button>
    );
}

/** 見出しのセルの下に、ルートの中の位置でメニューを置く（Power BI が拡大して描くと getBoundingClientRect は拡大した値なので倍率で割る） */
function columnMenuAt(root: HTMLElement, cell: HTMLElement, slot: number | null, periodKey: string): OpenColumnMenu {
    const box = root.getBoundingClientRect();
    const rect = cell.getBoundingClientRect();
    const scale = root.offsetWidth > 0 ? box.width / root.offsetWidth : 1;
    return { slot, periodKey, left: (rect.left - box.left) / scale, top: (rect.bottom - box.top) / scale, returnTo: cell };
}

/**
 * 期間の見出しの ▾ のダイアログ：その期間の比較の列 1〜3（表の上のバーの「比較」と同じ形で、上に「全期間と同じ」のチェック）。
 * 選ぶたびに表を組み直し、ダイアログは開いたまま（何列か続けて選べる）。Esc・外を押す・表をスクロールすると閉じる（キーボードで閉じたら
 * 見出しにフォーカスを戻す）
 */
function PeriodCompareDialog({
    open,
    viewModel,
    menu,
    onChoose,
    onClose,
}: {
    open: OpenColumnMenu;
    viewModel: ViewModel;
    menu: EventMenu;
    onChoose: (patch: ComparePatch) => void;
    onClose: (returnFocus: boolean) => void;
}): React.JSX.Element | null {
    const period = viewModel.periods.find((p) => p.key === open.periodKey);
    if (!period) return null;
    return (
        <CompareDialog
            title={`比較の列（${period.label}）`}
            items={menu.compareItems}
            slots={period.slots}
            period={periodNameOf(period)}
            same={!period.own}
            focusSlot={open.slot}
            left={open.left}
            top={open.top}
            opener=".ft-colmenu-open"
            onChoose={onChoose}
            onClose={onClose}
        />
    );
}

/** すべて開く（すべての組織を開き、行を閉じない）・すべて閉じる（根のほかの組織を閉じ、行もすべて閉じる：区分だけの一覧） */
function Folds({ folds, onSetFolds }: { folds: NonNullable<ViewModel["folds"]>; onSetFolds: (patch: Partial<VisualState>) => void }): React.JSX.Element {
    return (
        <span className="ft-folds">
            <button type="button" className="ft-fold-all" onClick={() => onSetFolds(folds.openAll)}>
                すべて開く
            </button>
            <button type="button" className="ft-fold-all" onClick={() => onSetFolds(folds.closeAll)}>
                すべて閉じる
            </button>
        </span>
    );
}

/** 薄くする覆いの白の濃さ（visual.less の .ft-dim::after と合わせる） */
const DIM_WHITE = 0.65;

/** 色を、薄くする覆いを重ねたときと同じだけ白に寄せる（#RRGGBB・#RGB。読めない色はそのまま） */
export function dimmed(color: string): string {
    const hex = color.trim().replace(/^#/, "");
    const full = hex.length === 3 ? hex.replace(/./g, (c) => c + c) : hex;
    if (!/^[0-9a-f]{6}$/i.test(full)) return color;
    const channel = (i: number) => Math.round(parseInt(full.slice(i, i + 2), 16) * (1 - DIM_WHITE) + 255 * DIM_WHITE);
    return `rgb(${channel(0)}, ${channel(2)}, ${channel(4)})`;
}

/** 行の上の線の太さ（px、visual.less と合わせる）。上の行から続く帯の塗りは、この太さの所を覆う */
/** 表の上のバーを表の幅で折り返すときの、いちばん狭い幅（px。題名と主・比較のメニューが 1 行に入る幅） */
const CAPTION_MIN_WIDTH = 280;
const ROW_BORDER_PX = 1;
const TOTAL_BORDER_PX = 3;

/**
 * 左に残す列（組織の列と行の見出し）の左端と、列の見出しの高さ。組織の列の幅は中身で決まるので、描いたあとに見出しの幅を測って
 * 足していく（sticky の left は列ごとに要る）。組織の名前は見出しの下に残す（top）。変わったときだけ描き直す。
 * 幅は offsetWidth（レイアウトの px）で測る（Power BI がビジュアルを拡大して描くと、getBoundingClientRect は拡大した幅を返し、left とずれる）。
 * 組織の列の幅がビジュアルの幅の 3 分の 1 を超えたら、組織の列は残さず、行の見出しだけを左に残す（狭いビジュアルで値が隠れないように。
 * ）
 */
function useStickyLefts(
    count: number,
    width: number
): {
    refs: Array<React.RefObject<HTMLTableCellElement | null>>;
    head: React.RefObject<HTMLTableSectionElement | null>;
    at: (column: number) => number | undefined;
    top: number;
    stick: boolean;
} {
    const refs = React.useMemo(() => Array.from({ length: count }, () => React.createRef<HTMLTableCellElement>()), [count]);
    const head = React.useRef<HTMLTableSectionElement>(null);
    const [measured, setMeasured] = React.useState<number[]>([]);
    React.useLayoutEffect(() => {
        const next = [head.current?.offsetHeight ?? 0, 0];
        for (const ref of refs) next.push(next[next.length - 1] + (ref.current?.offsetWidth ?? 0));
        if (next.length !== measured.length || next.some((v, i) => Math.abs(v - measured[i]) > 0.5)) setMeasured(next);
    });
    const stick = (measured[count + 1] ?? 0) <= width / 3;
    const at = (column: number) => (stick ? (measured[column + 1] ?? 0) : column === count ? 0 : undefined);
    return { refs, head, at, top: measured[0] ?? 0, stick };
}

interface RowProps {
    row: TableRow;
    viewModel: ViewModel;
    tooltip?: CellTooltip;
    /** 左に残す列の左端（組織の列の番号。組織の列の数なら行の見出し） */
    left: (column: number) => number | undefined;
    /** 列の見出しの高さ（組織の名前をその下に残す） */
    top: number;
    onToggleOrg?: (path: string) => void;
    onToggleRow?: (code: string) => void;
    picking: Picking;
    /** 期間の区切りの線を左に引く列の番号 */
    periodStarts: Set<number>;
    /** 上の行とのあいだが囲みの箱の始まりか終わりか（isBoxEdge） */
    boxEdge: boolean;
}

/** 行の見出しを明るく残すか。選べない行（計算行・指標）は薄くしない（組織を選んだときのほかの組織は薄くする） */
function nameActive(row: TableRow, selection: SelectTarget[]): boolean {
    const org = row.orgPath ?? null;
    return row.select ? rowActive(selection, row.key, org) : plainRowActive(selection, org);
}

/** 列の見出しを押したときの選択（期間の月） */
function periodTarget(viewModel: ViewModel, periodKey: string): SelectTarget | null {
    const period = viewModel.periods.find((p) => p.key === periodKey);
    return period ? { kind: "period", periodKey, months: period.months } : null;
}

/**
 * 行の上の線の種類。行の境目ごとに種類を 1 つ決め、名前の列から数字の列の
 * 右端まで同じ線種・色で引く（名前の列で変わるのは、箱の字下げに合わせた引き始めの位置だけ）。強い順に、セグメントの切れ目（ft-row-block）・
 * 最後の段階（ft-row-total、二重線）・明細の下で締める小計と途中の段階（ft-row-sumline・ft-row-midstep）・囲みの箱の始まりか終わり（ft-row-boxedge、囲みの線）・
 * それ以外。箱の中の科目どうしは行の間の線で、箱は左の縦線と始まり・終わりの実線で見せる（セグメントの箱と同じ作り）。名前と値の列で同じ境界を使う。
 */
function isBoxEdge(above: TableRow, row: TableRow): boolean {
    return row.bands.some((band) => band.first) || above.bands.some((band) => band.last);
}

interface LayoutViewProps extends AppProps {
    plan: LayoutPlan;
    unitInBand?: boolean;
    openMenu?: (cell: HTMLTableCellElement) => void;
    toggleMenu?: (cell: HTMLTableCellElement, period: string) => void;
    activeMenu?: { periodKey: string; returnTo: HTMLElement };
}

function LayoutView(props: LayoutViewProps): React.JSX.Element {
    const { plan } = props;
    if (plan.kind === "table") return <PresentedTable {...props} plan={plan} />;
    if (plan.kind === "stack") return <div className={`ft-layout-stack ft-layout-${plan.direction}`}>
        {plan.children.map(child => <LayoutView {...props} key={child.key} plan={child} />)}
    </div>;
    const style = props.viewModel.style;
    const ownTable = (p: LayoutPlan): boolean => p.kind === "table" || (p.kind === "stack" && p.children.some(ownTable));
    const unit = plan.style === "split" && ownTable(plan.content) && style.unitPlace === UNIT_PLACES.right ? props.viewModel.unitCaption : undefined;
    const fill = plan.header.kind === "org" ? (style.segmentFill ? style.segmentColors[0] : style.background) : (style.bandFill ? style.bandColors[0] : style.background);
    return <section className={`ft-layout-frame ft-layout-frame-${plan.header.kind} ft-layout-${plan.style}`} style={{ ["--ft-frame-fill" as string]: fill, ["--ft-frame-text" as string]: textOn(fill, style) }}>
        <FrameLabel {...props} header={plan.header} unit={unit} />
        <LayoutView {...props} plan={plan.content} unitInBand={props.unitInBand || !!unit} />
    </section>;
}

function FrameLabel({ header, unit, viewModel, selection = [], onSelect, onContextMenu, onToggleOrg, onToggleRow }: LayoutViewProps & { header: LayoutHeader; unit?: string }): React.JSX.Element {
    const stamp = React.useRef(0);
    const toggle = header.fold === undefined ? undefined : () => (header.kind === "org" ? onToggleOrg : onToggleRow)?.(header.fold!);
    const target: SelectTarget | null = header.kind === "org" && header.org != null ? { kind: "org", path: header.org }
        : header.row?.select ? { kind: "row", rowKey: header.row.key, codes: header.row.select.codes, org: header.row.select.org } : null;
    const picking: Picking = { selection, onSelect, onContextMenu, focused: header.key, onFocusCell: () => undefined,
        periods: new Map(viewModel.periods.map(p => [p.key, p])), keyMenu: { mark: () => { stamp.current = Date.now(); }, recent: () => Date.now() - stamp.current < KEY_MENU_GUARD_MS } };
    const active = header.row ? nameActive(header.row, selection) : plainRowActive(selection, header.org ?? null);
    return <div className={`ft-layout-label${active ? "" : " ft-dim"}`} tabIndex={0} data-cell={header.key}
        {...pickProps(picking, target, toggle ? { open: header.open!, toggle } : undefined)}>
        <span className="ft-layout-title"><span className="ft-toggle-slot">{toggle && <Toggle open={header.open!} above={header.above} label={header.label} onClick={toggle} />}</span>{header.label}</span>
        {unit && <span className="ft-layout-unit">{unit}</span>}
    </div>;
}

/**
 * 期間の種類が変わる列か（月 → 四半期 → 半期 → 通期の境目）。「集計の列の区切りを二重線」は、種類が変わる所だけ二重線にし、
 * 同じ種類の列どうし（1Q と 2Q など）は 1 本の線のままにする（4月 | 5月 | 6月 ‖ 1Q | 2Q ‖ 1H | 2H ‖ 通期）
 */
function kindChangeAt(viewModel: ViewModel, columnIndex: number): boolean {
    const kind = viewModel.columns[columnIndex]?.periodKind;
    const before = viewModel.columns[columnIndex - 1]?.periodKind;
    return before !== undefined && kind !== before;
}

function PresentedTable(props: LayoutViewProps & { plan: Extract<LayoutPlan, { kind: "table" }> }): React.JSX.Element {
    const { plan, viewModel, selection = [], onSelect, onContextMenu, onClearSelection, openMenu, tooltip, onToggleRow, onToggleOrg } = props;
    const model = React.useMemo<ViewModel>(() => ({ ...viewModel, rows: plan.rows, layout: undefined, orgColumns: Array.from({ length: plan.orgColumns ?? 0 }, () => viewModel.style.headers.segment) }), [viewModel, plan]);
    const keyboard = useKeyboard(model, onClearSelection, onContextMenu, openMenu);
    const picking: Picking = { selection, onSelect, onContextMenu, periods: new Map(viewModel.periods.map(p => [p.key, p])),
        focused: keyboard.focused, onFocusCell: keyboard.onFocusCell, keyMenu: keyboard.keyMenu };
    // 期間の区切り：期間の最初の列（表の最初の期間は除く）の左に縦の線（書式ペインで消せる）
    const starts = new Set(viewModel.style.periodLines ? viewModel.columns.flatMap((c, i) => i && c.periodKey !== viewModel.columns[i - 1].periodKey ? [i] : []) : []);
    const cornerUnit = viewModel.style.unitPlace === UNIT_PLACES.corner ? viewModel.unitCaption : undefined;
    const cornerTitle = viewModel.style.titleInCorner ? viewModel.title : "";
    /** セグメントの列（横積みで段ごとに並ぶ）の見出し。「セグメント」を段の数だけ並べず 1 つにまとめる */
    const orgHeadCell = (rowSpan: number) => {
        const header = plan.orgHeader;
        const toggle = header?.fold !== undefined && onToggleOrg ? () => onToggleOrg(header.fold!) : undefined;
        return <th data-width-key="o0" className="ft-corner ft-org-head" rowSpan={rowSpan} colSpan={model.orgColumns.length}
            {...focusProps(picking, cellKey.orgHead(0))}
            {...pickProps(picking, null, toggle ? { open: header!.open!, toggle } : undefined)}>
            {toggle && <span className="ft-toggle-slot"><Toggle open={header!.open!} label={header!.label} onClick={toggle} /></span>}{model.orgColumns[0]}
        </th>;
    };
    return <table className={`ft-table ft-layout-table${viewModel.twoLines ? " ft-two-lines" : ""}`} data-layout={plan.key} data-column-shape={`${plan.nameColumns}:${plan.orgColumns ?? 0}`}
        role="grid" aria-label={viewModel.title || "財務諸表"} aria-multiselectable={onSelect ? true : undefined}
        ref={keyboard.table} {...keyboard.tableProps}>
        {!props.unitInBand && viewModel.unitCaption && viewModel.style.unitPlace === UNIT_PLACES.right && <caption className="ft-block-unit">{viewModel.unitCaption}</caption>}
        <colgroup>{model.orgColumns.map((_, i) => <col data-width-key={`o${i}`} key={`o${i}`} />)}
            {Array.from({ length: viewModel.hideNames ? 0 : plan.nameColumns }, (_, i) => <col data-width-key={`n${i}`} key={i} />)}
            {viewModel.columns.map((_, i) => <col data-width-key={`v${i}`} key={`v${i}`} />)}</colgroup>
        <thead><tr>
            {/* セグメントの列（横積みで段ごとに並ぶ）は、見出しを 1 つにまとめる（「セグメント」を段の数だけ並べない） */}
            {model.orgColumns.length > 0 && cornerTitle && !viewModel.hideNames && <th className="ft-corner ft-corner-title" colSpan={model.orgColumns.length + plan.nameColumns}>{cornerTitle}</th>}
            {model.orgColumns.length > 0 && !(cornerTitle && !viewModel.hideNames) && orgHeadCell(2)}
            {/* 年度を左上の角に置くときは、角を 2 段に分け、上の段（期間の見出しと同じ段）に年度を置く */}
            {!viewModel.hideNames && (cornerTitle
                ? (model.orgColumns.length > 0 ? null : <th className="ft-corner ft-corner-title" colSpan={plan.nameColumns}>{cornerTitle}</th>)
                : <th className="ft-corner" data-width-key="n0" rowSpan={2} colSpan={plan.nameColumns} {...focusProps(picking, cellKey.corner)}>
                    <span className={cornerUnit ? "ft-corner-unit" : "ft-corner-label"}>{cornerUnit || viewModel.style.headers.row}</span>
                </th>)}
            {viewModel.periods.map((p, i) => <th key={p.key} className={`ft-period ft-period-${p.kind}${i > 0 && viewModel.style.periodLines ? " ft-period-start" : ""}${i > 0 && p.kind !== viewModel.periods[i - 1].kind ? " ft-period-agg" : ""}`} colSpan={p.span}
                {...focusProps(picking, cellKey.period(p.key))} {...pickProps(picking, periodTarget(viewModel, p.key))}
                {...(openMenu ? { "data-period": p.key, "aria-haspopup": "dialog" as const } : {})}>
                {openMenu ? <span className="ft-period-label"><span className="ft-colmenu-spacer" aria-hidden="true" />{p.label}
                    <PeriodMenuButton period={p} viewModel={viewModel}
                        open={props.activeMenu?.periodKey === p.key && props.activeMenu.returnTo.closest<HTMLTableElement>("table")?.dataset.layout === plan.key}
                        onToggle={cell => props.toggleMenu ? props.toggleMenu(cell, p.key) : openMenu(cell)} /></span> : p.label}
                {i > 0 && viewModel.style.periodLines && <span className="ft-period-line" />}
            </th>)}
        </tr><tr>{model.orgColumns.length > 0 && cornerTitle && !viewModel.hideNames && orgHeadCell(1)}{!viewModel.hideNames && cornerTitle && <th className="ft-corner" data-width-key="n0" colSpan={plan.nameColumns} {...focusProps(picking, cellKey.corner)}>
                <span className={cornerUnit ? "ft-corner-unit" : "ft-corner-label"}>{cornerUnit || viewModel.style.headers.row}</span>
            </th>}{viewModel.columns.map((c, i) => <th key={`${c.periodKey}:${c.slot ?? "main"}`} data-width-key={`v${i}`}
            className={`ft-colhead ft-col-${c.kind}${starts.has(i) ? " ft-period-start" : ""}${starts.has(i) && kindChangeAt(viewModel, i) ? " ft-period-agg" : ""}`}
            {...focusProps(picking, cellKey.column(i))} {...pickProps(picking, periodTarget(viewModel, c.periodKey))} {...tooltipProps(c.tooltip, tooltip)}
            {...(openMenu && c.slot !== undefined ? { "data-period": c.periodKey, "data-slot": c.slot, "aria-haspopup": "dialog" as const } : {})}>
            <div>{c.header}</div>{c.sub && <div className="ft-colhead-sub">{c.sub}</div>}
            {starts.has(i) && <span className="ft-period-line" />}
        </th>)}</tr></thead>
        <tbody>{plan.rows.map((row, i) => <Row key={row.key} row={row} viewModel={model} picking={picking} periodStarts={starts}
            boxEdge={i > 0 && isBoxEdge(plan.rows[i - 1], row)} left={() => undefined} top={0} tooltip={tooltip}
            onToggleRow={onToggleRow} onToggleOrg={onToggleOrg} />)}</tbody>
    </table>;
}

function Row({ row, viewModel, tooltip, left, top, onToggleOrg, onToggleRow, picking, periodStarts, boxEdge }: RowProps): React.JSX.Element {
    const classes = ["ft-row", `ft-row-${row.type}`];
    if (row.aggregatePosition === "bottom") classes.push("ft-row-end");
    if (row.aggregatePosition === "top") classes.push("ft-row-start");
    if (row.open === false) classes.push("ft-row-collapsed");
    if (row.open === true) classes.push("ft-row-expanded");
    if (row.hierarchyCells?.length) classes.push("ft-row-hierarchy-start");
    if (boxEdge) classes.push("ft-row-boxedge");
    const { selection } = picking;
    const org = row.orgPath ?? null;
    const rowTarget: SelectTarget | null = row.select
        ? {
              kind: "row",
              rowKey: row.key,
              codes: row.select.codes,
              org: row.select.org,
          }
        : null;
    const headActive = nameActive(row, selection);
    if (rowTarget && picking.onSelect) classes.push("ft-pickable-row");
    // 組織のブロックの最初の行は、上に組織の境の線を引く
    if (row.orgCells && row.orgCells.length > 0) classes.push("ft-row-block");
    // 段階の行は、続く段階の無い最後の合計（当期純利益・資産合計など）だけ二重線。途中の段階（売上総利益・営業利益など）は小計の線
    const total = row.type === "step" && !row.continued && row.open !== false && row.aggregatePosition !== "top";
    if (total) classes.push("ft-row-total");
    else if (row.type === "step") classes.push("ft-row-midstep");
    // 自分がまとまりを締める集計先か
    const head = row.bands.find((b) => b.head && b.level === row.depth);
    // 小計の線（濃い実線）は、見えている明細の下で締める小計だけ（明細を足した所の区切り）。明細の上に置く小計・閉じた区分は箱の始まりなので
    // 囲みの線（「PL はほぼ全行が太字＋濃い横線で、営業利益が埋もれる」。閉じた区分が並ぶ損益計算書で、どの行も濃い線だった）
    // 囲みを切った表（帯が無い）は今までどおりどの小計も小計の線
    if (row.type === "subtotal" && (!head || !head.first)) classes.push("ft-row-sumline");
    const { bandColors, bandFill } = viewModel.style;
    const colorOf = (level: number) => bandColors[Math.min(level, bandColors.length - 1)];
    // 入れ子の箱（区分・中分類）の中の行か。箱の中の行は、囲む親の段ごとの帯と、自分の段の名前の所を箱として描く。
    // 箱の外の行（計算行・区分に入らない科目）は今までどおり行の線だけ
    const inBox = row.bands.length > 0;
    // 塗るのはまとまりを締める行（区分・中分類の小計・見出し）の名前の所だけ。明細の行は地の色で、左の帯だけ親の段の色
    //
    const fillColor = inBox && bandFill && head ? colorOf(row.depth) : null;
    const border = total ? TOTAL_BORDER_PX : ROW_BORDER_PX;
    // 開き閉じはブロックごとに覚える（fold はブロックの道筋とコード）
    const toggleRow = row.open !== undefined && onToggleRow ? () => onToggleRow(row.fold ?? row.code) : undefined;
    return (
        <tr className={classes.join(" ")}>
            {/* key はセルの名前にする。列の番号にすると、組織を開き閉じして行が使い回されたとき、フォーカスが別の組織の名前に移った */}
            {row.orgCells?.map((cell) => (
                <OrgHead
                    key={cellKey.org(cell, row.key)}
                    cell={cell}
                    rowKey={row.key}
                    left={left(cell.column)}
                    top={top}
                    onToggle={onToggleOrg}
                    picking={picking}
                    fill={viewModel.style.segmentFill ? viewModel.style.segmentColors : null}
                />
            ))}
            {row.hierarchyCells?.map(cell => {
                const parent = cell.row;
                const toggle = parent.open !== undefined && onToggleRow ? () => onToggleRow(parent.fold ?? parent.code) : undefined;
                const target: SelectTarget | null = parent.select ? { kind: "row", rowKey: parent.key, codes: parent.select.codes, org: parent.select.org } : null;
                return <th key={parent.key} className={`ft-hierarchy-name${nameActive(parent, selection) ? "" : " ft-dim"}`} rowSpan={cell.rowSpan} data-width-key={`n${cell.column}`}
                    {...focusProps(picking, cellKey.name(`hierarchy:${parent.key}`))}
                    {...pickProps(picking, target, toggle ? { open: parent.open!, toggle } : undefined)}>
                    <span className="ft-toggle-slot">{toggle && <Toggle open={parent.open!} label={parent.name} onClick={toggle} />}</span>
                    {parent.name}
                </th>;
            })}
            {!viewModel.hideNames && <th
                scope="row"
                data-width-key={`n${row.nameColumn ?? 0}`}
                colSpan={row.nameSpan}
                className={`ft-name${head && (row.type === "subtotal" || row.type === "heading") ? " ft-name-head" : ""}${inBox ? " ft-name-inbox" : ""}${headActive ? "" : " ft-dim"}`}
                {...focusProps(picking, cellKey.name(row.key))}
                {...pickProps(picking, rowTarget, toggleRow ? { open: row.open!, toggle: toggleRow } : undefined)}
                style={{
                    left: left(viewModel.orgColumns.length),
                    paddingLeft: `calc(var(--ft-label-inset) + ${row.depth * INDENT_EM}em)`,
                    ...(fillColor ? { color: textOn(fillColor, viewModel.style) } : {}),
                }}
            >
                {inBox ? (
                    <>
                        <BoxMarks row={row} colorOf={bandFill ? colorOf : null} background={viewModel.style.background} border={border} />
                        {/* 下に置いた合計の締めの線は、自分の段から右だけ（外の箱の帯を横切らない） */}
                        {row.aggregatePosition === "bottom" && <span className="ft-box-bottom" style={{ left: `${row.depth * INDENT_EM}em`, right: 0 }} />}
                    </>
                ) : (
                    // 線の種類は行で決まる（isBoxEdge）。上の行が箱の中なら箱の底の辺なので囲みの線になる
                    <span className="ft-name-line" style={{ left: 0, top: 0 }} />
                )}
                <span className="ft-name-text">
                    <span className="ft-toggle-slot">{toggleRow && <Toggle open={row.open!} above={row.aggregatePosition === "bottom" || (row.aggregatePosition === undefined && row.toggleAbove)} label={row.name} onClick={toggleRow} />}</span>
                    {row.type === "breakdown" && viewModel.style.breakdown.tag && <span className="ft-tag">うち</span>}
                    {row.type === "blank" ? " " : row.name}
                    {row.unit && <span className="ft-name-unit">{row.unit}</span>}
                </span>
            </th>}
            {row.cells.map((cell, i) => {
                const column = viewModel.columns[i];
                const period = column ? picking.periods.get(column.periodKey) : undefined;
                const target: SelectTarget | null =
                    row.select && period
                        ? {
                              kind: "cell",
                              rowKey: row.key,
                              periodKey: period.key,
                              codes: row.select.codes,
                              org: row.select.org,
                              months: period.months,
                          }
                        : null;
                const active = row.select ? cellActive(selection, row.key, column?.periodKey ?? "", org) : plainCellActive(selection, column?.periodKey ?? "", org);
                return (
                    <ValueCell
                        key={i}
                        columnIndex={i}
                        cell={cell}
                        kind={column?.kind}
                        viewModel={viewModel}
                        breakdown={row.type === "breakdown"}
                        periodStart={periodStarts.has(i)}
                        tooltip={tooltip}
                        dim={!active}
                        pick={{
                            ...focusProps(picking, cellKey.value(row.key, i)),
                            ...pickProps(picking, target),
                        }}
                    />
                );
            })}
        </tr>
    );
}

/**
 * 左の組織の列のセル。名前をその組織の行にまとめて縦にかける。計の名前は残りの段の列をまたぐ。
 * 名前は、下へスクロールしてもその組織の行が見えているあいだ、列の見出しの下に残す
 */
function OrgHead({
    cell,
    rowKey,
    left,
    top,
    onToggle,
    picking,
    fill,
}: {
    cell: OrgCell;
    rowKey: string;
    left: number | undefined;
    top: number;
    onToggle?: (path: string) => void;
    picking: Picking;
    /** セグメントの箱を塗るときの段ごとの色（塗らなければ null） */
    fill: string[] | null;
}): React.JSX.Element {
    const toggle = cell.toggle && onToggle ? cell.toggle : null;
    const target: SelectTarget | null = cell.path !== null ? { kind: "org", path: cell.path } : null;
    const selected = orgSelected(picking.selection, cell.path);
    const depth = cell.depth ?? 0;
    // ブロックの切れ目の横線は、上のブロックと共通する一番近い親の縦線から引く
    return (
        <th
            rowSpan={cell.rowSpan}
            data-width-key={`o${cell.column}`}
            colSpan={cell.colSpan}
            className={`ft-org${cell.total ? " ft-org-total" : ""}${left === undefined ? " ft-org-free" : ""}${selected ? " ft-org-selected" : ""}${target && picking.onSelect ? " ft-pickable" : ""}`}
            style={{ left, ["--ft-seg-line" as string]: railAt(cell.lineFrom ?? 0) }}
            {...focusProps(picking, cellKey.org(cell, rowKey))}
            {...pickProps(picking, target, toggle ? { open: toggle.open, toggle: () => onToggle!(toggle.path) } : undefined)}
        >
            {/*
              入れ子の箱：囲む親の段ごとに 1 段の幅の帯（親の箱の左の部分）、自分の段から右は名前の所。
              帯は親の箱がこのブロックで始まれば上の辺、終われば下の辺を引く。名前の所は左と上の辺（次のブロックが深ければ下の辺も）
            */}
            {/* 囲みなしは塗りと帯を引かず、ブロックの切れ目の横線だけ引く */}
            <NestedBox
                // ブロックの境目の横線は下のブロックだけが引く。帯の上の辺は、親の箱がこのブロックで始まる所だけ
                strips={cell.plain ? [] : (cell.rails ?? []).map((d) => ({ level: d, topLine: cell.boxStart?.includes(d) ?? false }))}
                depth={depth}
                // 帯が上から続いてくる所（小計が下の親の小計）は、帯の幅に横線を引かない
                ownStripTop={cell.plain || !cell.ownAbove}
                colorOf={fill && !cell.plain ? (d) => fill[Math.min(d, fill.length - 1)] : null}
                indent={SEGMENT_INDENT_EM}
                border={1}
                // 囲みなしは段が 0 なので、左の縦線も引かない（NestedBox は 0 段の左の辺を引かない）
                className="ft-seg-box"
            />
            <div className="ft-org-label" style={{ top }}>
                {cell.heads?.map((head) => (
                    <div key={head.path ?? ""} className="ft-org-line" style={{ paddingLeft: `${head.depth * SEGMENT_INDENT_EM}em` }}>
                        <span className="ft-toggle-slot">{head.toggle && onToggle && <Toggle open={head.toggle.open} label={head.label} onClick={() => onToggle(head.toggle!.path)} />}</span>
                        {head.label}
                    </div>
                ))}
                <div className="ft-org-line" style={{ paddingLeft: `${depth * SEGMENT_INDENT_EM}em` }}>
                    {/* 計を中身の下に置いたとき（子のブロックが上）は、開いたしるしを上向きにする（行の小計を下に置いたときと同じ） */}
                    <span className="ft-toggle-slot">{toggle && <Toggle open={toggle.open} above={!!cell.ownAbove} label={cell.label} onClick={() => onToggle!(toggle.path)} />}</span>
                    {cell.label}
                </div>
            </div>
        </th>
    );
}

/** セグメントの字下げの幅（em）。段ごとの縦線は、その段の名前の開き閉じのしるしの下に引く */
const SEGMENT_INDENT_EM = 1;
const railAt = (depth: number) => `calc(0.5em + ${depth * SEGMENT_INDENT_EM}em + 0.3em)`;

/** 開き閉じのしるし（▾ 開いている、▸ 閉じている） */
function Toggle({ open, above, label, onClick }: { open: boolean; above?: boolean; label: string; onClick: () => void }): React.JSX.Element {
    return (
        <button
            type="button"
            className="ft-toggle"
            aria-expanded={open}
            aria-label={`${label}を${open ? "閉じる" : "開く"}`}
            // Tab で止めない（キーボードでは、しるしのあるセルで Shift+→・← を押す）。押してもフォーカスを取らない
            tabIndex={-1}
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
                // 開き閉じは選択にしない
                e.stopPropagation();
                onClick();
            }}
        >
            {open ? above ? "▴" : "▾" : "▸"}
        </button>
    );
}

/**
 * 入れ子の箱。囲む親の段ごとに 1 段の幅の帯を並べ（親の箱の左の部分。親の色）、自分の段から右を
 * 名前の所（自分の段の色）にする。帯は親の箱がこの行で始まれば上の辺、終われば下の辺を引く。名前の所は左と上の辺（次の行が深い段なら下の辺も）。
 * 名前のセルは上の枠線を持たず（visual.less）、線と塗りはセルの上の端（行の上の端）から描く。数字の列の行の線は枠線で、枠線の太さは
 * 画面の画素に丸められるので、「枠線の太さだけ上へ」と px で決め打ちで重ねると、Desktop の倍率で 1 画素ずれた。上の端どうしなら、同じ所に同じ太さで描かれる。
 * 一番外の段（0 段）の箱の左の辺は引かない。左隣にセグメントの列の右の線か表の外枠があり、並べると線が 2 本に見えた
 */
function BoxMarks({
    row,
    colorOf,
    background,
    border,
}: {
    row: TableRow;
    colorOf: ((level: number) => string) | null;
    background: string;
    border: number;
}): React.JSX.Element {
    // 自分が締めるまとまり（小計・見出しの行）
    const own = row.bands.find((band) => band.head && band.level === row.depth);
    return (
        <NestedBox
            // 行の境目の横線は下の行だけが引く。帯の上の辺は、親の箱がこの行で始まる所だけ（小計が上の親の小計のすぐ下は、帯が小計の
            // 名前の所から続くので引かない）
            strips={row.bands.filter((band) => band.level < row.depth).map((band) => ({ level: band.level, topLine: band.first }))}
            depth={row.depth}
            // 名前の所の、自分の段の帯の幅の上の辺：自分の箱が上の行から続いている（小計が下）なら引かない
            ownStripTop={!(own && !own.first)}
            colorOf={colorOf}
            // 明細の行（締めるまとまりが無い）の名前の所は地の色で、外の段の塗りを隠す
            ownColor={own ? undefined : background}
            indent={INDENT_EM}
            border={border}
            className="ft-box"
        />
    );
}

/**
 * 入れ子の箱の塗りと線（区分の箱とセグメントの箱で共通）。塗りはどれも右端まで伸ばし、外の段から順に重ねる。線は塗りと分け、左の辺は塗りの左の縁、上の辺は高さ 0 の線を
 * 要る幅だけに引く。塗りはセルの中（上の枠線の所まで）だけで、上の行へ重ねない（重ねた 1px が上の行の名前の所へはみ出し、Desktop の
 * 倍率で線の上に 1 段外の色がにじんで見えた。2026-09-29、125% に拡大して色を 1px ずつ読んで確かめた）
 */
function NestedBox({
    strips,
    depth,
    ownStripTop,
    colorOf,
    ownColor,
    indent,
    border,
    className,
}: {
    strips: Array<{ level: number; topLine: boolean }>;
    depth: number;
    ownStripTop: boolean;
    colorOf: ((level: number) => string) | null;
    /** 自分の段の塗りを段の色でなくこの色にする（塗りがあるときだけ） */
    ownColor?: string;
    indent: number;
    border: number;
    className: string;
}): React.JSX.Element {
    const at = (level: number) => `${level * indent}em`;
    const paint = (level: number) => (colorOf ? colorOf(level) : undefined);
    const edge = (level: number): React.CSSProperties => (level === 0 ? { borderLeftStyle: "none" } : {});
    return (
        <>
            {strips.map((strip) => (
                <span key={strip.level} className={className} style={{ left: at(strip.level), right: 0, top: 0, background: paint(strip.level), ...edge(strip.level) }} />
            ))}
            <span className={className} style={{ left: at(depth), right: 0, top: 0, background: colorOf && ownColor ? ownColor : paint(depth), ...edge(depth) }} />
            {/*
              上の行から続く帯は、帯の幅だけ 1px 上へ重ねて塗る（行の境目が小数の px だと、白い横線が透けた）。上の行のその幅は同じ箱の同じ色なので
              にじまない
            */}
            {strips
                .filter((strip) => !strip.topLine)
                .map((strip) => (
                    <span
                        key={`p${strip.level}`}
                        className={className}
                        style={{ left: at(strip.level), width: `${indent}em`, top: -1, height: border + 2, bottom: "auto", background: paint(strip.level), ...edge(strip.level) }}
                    />
                ))}
            {!ownStripTop && (
                <span className={className} style={{ left: at(depth), width: `${indent}em`, top: -1, height: border + 2, bottom: "auto", background: paint(depth), ...edge(depth) }} />
            )}
            {strips
                .filter((strip) => strip.topLine)
                .map((strip) => (
                    <span key={`l${strip.level}`} className={`${className}-line`} style={{ left: at(strip.level), width: `${indent}em`, top: 0 }} />
                ))}
            <span className={`${className}-line`} style={{ left: at(ownStripTop ? depth : depth + 1), right: 0, top: 0 }} />
        </>
    );
}

function ValueCell({
    cell,
    kind,
    viewModel,
    breakdown,
    periodStart,
    tooltip,
    dim,
    pick,
    columnIndex,
}: {
    cell: Cell;
    kind: string | undefined;
    viewModel: ViewModel;
    breakdown: boolean;
    periodStart: boolean;
    tooltip?: CellTooltip;
    dim: boolean;
    pick: React.HTMLAttributes<HTMLElement>;
    columnIndex: number;
}): React.JSX.Element {
    const color = cell.tone === "good" ? viewModel.style.good : cell.tone === "bad" ? viewModel.style.bad : undefined;
    // 内訳（うち）の値は角括弧で囲み、合計に足さない行だと分かるようにする（「行」カードの「うちの行」で切れる）
    const bracket = breakdown && kind === "main" && viewModel.style.breakdown.brackets;
    const wrap = (text: string) => (bracket && text ? `[${text}]` : text);
    return (
        <td
            data-value-column={columnIndex}
            data-width-key={`v${columnIndex}`}
            className={`ft-value ft-col-${kind ?? "main"}${periodStart ? " ft-period-start" : ""}${periodStart && kindChangeAt(viewModel, columnIndex) ? " ft-period-agg" : ""}${dim ? " ft-dim" : ""}`}
            style={color ? { color } : undefined}
            {...tooltipProps(cell.tooltip, tooltip)}
            {...(pick as React.HTMLAttributes<HTMLTableCellElement>)}
        >
            <div>{cell.parts ? <NumberParts parts={cell.parts} bracket={bracket && cell.text !== ""} /> : withSuffix(wrap(cell.text))}</div>
            {periodStart && <span className="ft-period-line" />}
            {cell.sub !== undefined && <div className="ft-sub">{cell.subParts ? <NumberParts parts={cell.subParts} bracket={false} /> : withSuffix(cell.sub)}</div>}
        </td>
    );
}

function NumberParts({ parts, bracket }: { parts: ValueParts; bracket: boolean }): React.JSX.Element {
    return <span className="ft-number-parts"><span className="ft-number-core">{bracket ? "[" : ""}{parts.number}</span>
        <span className="ft-number-tail"><span className="ft-number-tail-content"><span className={`ft-number-unit${parts.suffix === "%" ? " ft-number-percent" : ""}`}>{parts.suffix}</span>{parts.closing}{bracket ? "]" : ""}</span></span>
    </span>;
}

/**
 * 数字の後ろの単位（h・円/h・人・pt）を少し小さくして、数字と分ける（10,504円/h が数字と単位の同じ強さで続き、窮屈だった）。% は数字の一部として読むのでそのまま。マイナスの括弧・うちの角括弧は単位にしない
 */
const SUFFIX = /^(.*\d)([^\d\])%]+)(\]?)$/;
export function withSuffix(text: string): React.ReactNode {
    const match = SUFFIX.exec(text);
    if (!match) return text;
    return (
        <>
            {match[1]}
            <span className="ft-suffix">{match[2]}</span>
            {match[3]}
        </>
    );
}

/** セルと列の見出しのツールチップ。tooltipService があれば Power BI のツールチップ、無ければ title */
function tooltipProps(items: TooltipItem[] | undefined, tooltip: CellTooltip | undefined): React.HTMLAttributes<HTMLTableCellElement> {
    if (!items) return {};
    if (!tooltip)
        return {
            title: items.map((i) => `${i.displayName}：${i.value}`).join("\n"),
        };
    return {
        onMouseEnter: (e) => tooltip.show(items, e.clientX, e.clientY, false),
        onMouseMove: (e) => tooltip.show(items, e.clientX, e.clientY, true),
        onMouseLeave: () => tooltip.hide(),
    };
}
