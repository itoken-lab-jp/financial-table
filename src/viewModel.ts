/**
 * DataView → 表の中身。描画（App.tsx）はここで決めた文字と色を並べるだけにする。
 *
 * 流れ：data.ts で読む → events.ts でイベントの比較順を決める → rows.ts で行の木を組む（指標の行も置く）→ signs.ts で金額を
 * 借方プラスにそろえる → compute.ts で期間と主・比較の値を出す → ここで文字にする。
 * 組織の欄を入れたら、orgs.ts で組織のブロックを並べ、ブロックごとにその組織より下のファクトで計算する。
 */
import powerbi from "powerbi-visuals-api";
import { AccountNode, HierarchyCell, LayoutPlan, accountTree, accountPlan, compactOrgBoxes, planRows, stack } from "./hierarchy";
import { HierarchyDefaults, hierarchyOptions, readHierarchyOverrides } from "./hierarchySettings";
import { ROOT_FOLD } from "./orgs";
import { ValueParts, valueParts } from "./valueParts";

import { resolveUnit } from "./shared/units";
import { SCROLL_STARTS } from "./shared/scrollStart";
import { contrastingText, readableText } from "./shared/color";
import { NEGATIVE_STYLES, TONE_MODES, ZERO_STYLES, toneOf, isPercentFormat } from "./shared/numberFormat";
import { Calculator, EventRef } from "./compute";
import { ALL_PERIODS_MONTH, COLUMN_LIMIT, TOTAL_ACCOUNT, Fact, INDICATOR_KEY, InputData, ORDER_OF_SECTIONS, ORG_SEPARATOR, SelectionNodes, readInput, savedString } from "./data";
import { EventModel, SIGN_VALUES, fallbackOf, resolveEvents } from "./events";
import {
    AMOUNT_FORMAT,
    AmountUnit,
    EXACT_EXTRA_DECIMALS,
    NOT_AVAILABLE,
    RowFormat,
    SignOptions,
    diffSign,
    formatDiff,
    formatExact,
    formatRate,
    formatRatio,
    formatRowValue,
    parseRowFormat,
} from "./numberFormat";
import {
    FiscalCalendar,
    MonthIndex,
    PeriodChoiceYear,
    PeriodKind,
    PeriodOptions,
    YearLabel,
    fiscalYearName,
    fiscalYearStart,
    monthOf,
    periodChoices,
    periodIdOf,
    tablePeriods,
    yearOf,
} from "./periods";
import {
    DEFAULT_ROOT_NAME,
    OrgCell,
    OrgNode,
    buildOrgTree,
    codesByOrg,
    containsOrg,
    emptyAccountRows,
    foldable,
    hiddenRows,
    toggleablePaths,
} from "./orgs";
import { OTHERS_KEY, PickParent, RowPick, UNDER_KEY, applyPicks, isAccountRow, pickParents } from "./picks";
import { RowNumberOverride, overrideKey, parentsOf, resolveRowNumbers, rowNumberOverrides, rowNumberTargets } from "./rowNumbers";
import { CalcRowSpec, DisplayRow, IndicatorSpec, RowDef, RowType, TOTAL_KEY, buildRows } from "./rows";
import { AmountSign, resolveSigns, toDebitPlus } from "./signs";
import { EMPTY_VISUAL_STATE, VisualState, shownSlots, slotsOf } from "./visualState";
import { DEFAULT_THEME, Theme } from "./theme";
import {
    CALC_PROPS,
    CALC_ROW_SLOTS,
    ROW_NUMBER_PROPS,
    ROW_NUMBER_SLOTS,
    RowNumberSaved,
    CALC_TABLE_END,
    CALC_TABLE_START,
    CalcSaved,
    DataDrivenFormat,
    DEFAULT_BAD_COLOR,
    DEFAULT_CURRENCY,
    DEFAULT_FONT_SIZE,
    DEFAULT_HEADING_SIZE,
    DEFAULT_GOOD_COLOR,
    DEFAULT_HEAD_BACKGROUND,
    DEFAULT_LINE_COLORS,
    DEFAULT_RATIO_CAP,
    DEFAULT_TEXT_SIZES,
    COMPARE_VIEWS,
    CompareView,
    DEFAULT_COMPARE_VIEW,
    INDICATOR_ALL_EVENTS,
    INDICATOR_MODES,
    IndicatorSetting,
    LATEST,
    LATEST_LABEL,
    LATEST_HEADERS,
    LATEST_YEAR,
    PRIOR_YEAR,
    PRIOR_YEAR_LABEL,
    TOTAL_ROW,
    UNIT_PLACES,
    TITLE_PLACES,
    ROW_HEIGHTS,
    RowHeight,
    BOLD_KEYS,
    BoldKey,
    BoldMode,
    ColumnPadding,
    COLUMN_PADDINGS,
    VisualFormattingSettingsModel,
    calcProp,
    rowNumberProp,
    dropdownValue,
} from "./settings";

import DataView = powerbi.DataView;

export const LOADING_NOTICE = "残りの行を読み込んでいます…";
/** 月の欄を入れていない表の、1 つだけの期間の列の名前 */
export const ALL_PERIODS_LABEL = "全期間";
/** 指標の値を読む小計の設定を保存して、読み直しを待つあいだの知らせ */
export const INDICATOR_NOTICE = "指標の値を読み込んでいます…";
/** 小計を入れる設定を保存したのに、指標の値（組織 × イベント × 月の小計）が届かないとき */
export const INDICATOR_SUBTOTAL_WARNING =
    "指標の値（セグメント × シナリオ × 月の小計）が届いていない。ビジュアルが小計を入れる設定を保存したが、まだ届かない。作り手が Desktop でレポートを開き直して保存する（閲覧だけの場では設定を保存できない）";
/** 読み込みの上限。期間を絞るなら、表に出す月とその前年同期の月がそろう範囲にする */
export const TRUNCATED_NOTICE =
    "行が多すぎて、すべてを読み込めませんでした。表の合計が小さく出ています。ビジュアルのフィルターで行やセグメントを絞るか、期間を絞るときは表に出す月とその前年同期の月がそろう範囲にしてください";

export type Tone = "good" | "bad" | null;

export interface Cell {
    parts?: ValueParts;
    subParts?: ValueParts;
    text: string;
    /** 差の列の 2 段目（比・率） */
    sub?: string;
    tone: Tone;
    /** ツールチップ：その期間の主・比較・差・比の正確な値（表は表示単位で丸める） */
    tooltip?: TooltipItem[];
}

/** ツールチップの 1 行（Power BI の VisualTooltipDataItem と同じ形） */
export interface TooltipItem {
    displayName: string;
    value: string;
}

/** 列の種類：主、比較の値、差・比・率（見せ方は view） */
export type ValueColumnKind = "main" | "compare" | "diff";

export interface ValueColumn {
    periodKey: string;
    periodKind: PeriodKind;
    kind: ValueColumnKind;
    /** 比較の列のシナリオと見せ方と、列 1〜3 のどれか（0 始まり）。主の列には無い */
    compare?: string;
    view?: CompareView;
    slot?: number;
    /** 期間の名前（年月つき。見る人の「この期間だけ」の比較対象の保存に使う） */
    periodName?: string;
    header: string;
    /** 見出しの 2 行目（差の列は「修正予算比」など、セルの下の段の名前） */
    sub: string;
    /** 見出しの月ごとのイベント（4〜5月=実績 / 6月=見通し）。最新見込みの主・落とした比較のときだけ */
    title?: string;
    /** 同じ内訳をツールチップの行にしたもの（4〜5月 → 実績） */
    tooltip?: TooltipItem[];
}

/** 囲みの帯の 1 本。level は集計先の行の深さ、head はその集計先の行そのもの */
export interface Band {
    level: number;
    head: boolean;
    first: boolean;
    last: boolean;
    /** すぐ上の行がこのまとまりの集計先（上囲みで塗った行の真下） */
    afterHead: boolean;
}

/** 表の上の、見る人が主を選ぶメニューと、比較の列の ▾ の選択肢 */
export interface EventMenu {
    main: string;
    /**
     * 作り手が書式ペインで決めた主（保存値のまま。決めていなければ null）。見る人の選択と一緒に保存し、作り手が決め直したら選択を捨てる。
     * データで決まる既定（最新見込み・比較順の一番小さいイベント）は入れない（データが変わっただけで見る人の選択を捨てないため）
     */
    mainBase: string | null;
    mainItems: Array<{ value: string; label: string }>;
    /** シナリオの選択肢（今の主に合わせる。主と同じイベント・表の年度にデータの無いイベントは入れない） */
    compareItems: Array<{ value: string; label: string }>;
    /** 全期間の列 1〜3（まだ選んでいなければ既定の 1 列）と、まだ選んでいないときの既定のシナリオと見せ方（chooseCompare に渡す） */
    compares: { all: SlotView[]; fallback: string; view: CompareView };
}

/** ダイアログに出す比較の列 1 つ：シナリオ（null は「なし」）、見せ方、今選べるか（主と同じ・表の年度にデータが無ければ false）、名前 */
export interface SlotView {
    compare: string | null;
    view: CompareView;
    available: boolean;
    label: string;
}

export interface PeriodHeader {
    key: string;
    label: string;
    kind: PeriodKind;
    span: number;
    /** 期間の月（見出しを押したときに絞る月） */
    months: MonthIndex[];
    /** この期間の列 1〜3 と、この期間だけの列を持っているか（「全期間と同じ」なら false） */
    slots: SlotView[];
    own: boolean;
}

/** 行を押したときに絞るもの：科目（区分・中分類の行はその科目すべて）と組織の道筋（組織のブロックの中なら） */
export interface RowSelect {
    codes: string[];
    org: string | null;
}

export interface TableRow {
    aggregatePosition?: "top" | "bottom";
    hierarchyCells?: HierarchyCell[];
    nameColumn?: number;
    nameSpan?: number;
    /** 表の中で 1 つに決まるキー（組織のブロックごとに同じ行が並ぶので、組織の道筋とコード） */
    key: string;
    code: string;
    name: string;
    /** 単位を行の名前の横に置くときの単位（（百万円）（台））。名前のあとに小さく薄く出す */
    unit?: string;
    type: RowType;
    /** 途中の段階の行（後の段階が続く。上の線は小計の線） */
    continued?: boolean;
    depth: number;
    /** 左の帯（外側から） */
    bands: Band[];
    cells: Cell[];
    /** 開き閉じできる行（区分・中分類の小計・見出し）なら、開いているか。できない行は undefined */
    open?: boolean;
    /** 開き閉じを覚える名前（セグメントのブロックがあればブロックの道筋とコード、無ければコード） */
    fold?: string;
    /** 開いているときの印を上向き（▴）にするか：小計を下に置く表（集計の行が中身の下にある） */
    toggleAbove?: boolean;
    /** この行から始まる、左の組織の列のセル（組織の欄を入れたときだけ） */
    orgCells?: OrgCell[];
    /** 押したときに絞るもの。計算行・指標・見出しの無い行など、科目に結びつかない行は undefined（選べない） */
    select?: RowSelect;
    /** 組織のブロックの道筋（組織を選んだときに薄くするかを決める。組織の欄が無い・全社は null） */
    orgPath?: string | null;
}

export interface ViewModel {
    layout?: LayoutPlan;
    /** 入れるべきフィールドが足りないときの案内。空なら表を出す */
    landing: string[];
    title: string;
    /** 表の年度（期首の月の「年-月」）。呼び方を変えても変わらない */
    yearKey: string;
    unitCaption: string;
    mainLabel: string;
    /** 全期間の 1 本目の比較対象の名前（無ければ空） */
    compareLabel: string;
    periods: PeriodHeader[];
    columns: ValueColumn[];
    rows: TableRow[];
    /**
     * 見せる科目を選ぶダイアログ：科目を選べる区分・中分類と、見る人の今の選択。選べる親が無ければ null（ボタンを出さない）
     */
    picks: { parents: PickParent[]; current: RowPick[] } | null;
    /**
     * 期間を選ぶダイアログ：データのある年度ごとの列の候補、今出している列（名前）と合計の列、見る人が選んでいるか、
     * 今の書式ペインの期間の設定（見る人の選択と組にする）。データが無ければ null
     */
    periodPicker: { years: PeriodChoiceYear[]; current: string[]; total: boolean; picked: boolean; base: string } | null;
    /** 見る人が行を開き閉じしたときに書く並び：行を既定で閉じる表（組織のブロックがある）は開いた行、既定で開く表は閉じた行 */
    rowStateKey: "openRows" | "closedRows";
    /** 左の組織の列の見出し（組織の段の名前）。組織の欄を入れていなければ空 */
    orgColumns: string[];
    /** 選択 ID を作る節点（visual.ts が使う） */
    selectionNodes: SelectionNodes | null;
    /**
     * すべて開く・すべて閉じる：開き閉じできる組織の道筋と行のコード、押したときに保存する値（開いた組織・既定から切り替えた行）。
     * どちらも無ければ null（ボタンを出さない）
     */
    folds: {
        orgs: string[];
        rows: string[];
        openAll: Partial<VisualState>;
        closeAll: Partial<VisualState>;
    } | null;
    warnings: string[];
    notice: string | null;
    truncated: boolean;
    format: DataDrivenFormat;
    menu: EventMenu | null;
    /** 列がはみ出すときの最初の位置（start・end） */
    scrollStart: string;
    /** マウスを乗せたとき「画像としてコピー」のボタンを出すか（書式ペインの「表示」） */
    copyButton: boolean;
    style: TableStyle;
    /** 2 段の比較のセルがあるか（あれば 1 段の行も 2 段の高さにそろえる） */
    twoLines: boolean;
    /** 行の名前の列を出さないか：科目の欄を入れない表（合計の 1 行）で、合計行を「非表示」にしたとき */
    hideNames?: boolean;
    /**
     * 指標の欄にメジャーがあるのに、組織 × イベント × 月の小計（科目の最初の段の小計）が届いていないときの、その段の欄（queryName）と、
     * 小計を入れる設定がもう保存されているか。届いていれば null。visual.ts が設定を保存して読み直す（保存してもまだ届かなければ知らせる）
     */
    indicatorSubtotals: { queryName: string; enabled: boolean } | null;
}

/** 書式の保存値を生で読む（データ次第の選択肢は、populate の時点では items に無いので value に入らない） */
function savedText(dataView: DataView | undefined, object: string, property: string): string | null {
    return savedString(dataView?.metadata?.objects?.[object]?.[property]);
}

/** 行別の数値書式の枠（書式ペイン）の保存値を生で読む。対象の選択肢はデータ次第なので、populate では値が入らない */
export function readRowNumbersSaved(dataView: DataView | undefined): RowNumberSaved[] {
    return Array.from(
        { length: ROW_NUMBER_SLOTS },
        (_, i) => Object.fromEntries(ROW_NUMBER_PROPS.map((prop) => [prop, savedText(dataView, "rowNumbers", rowNumberProp(prop, i + 1)) ?? ""])) as RowNumberSaved
    );
}

/** 計算行（書式ペイン）の保存値を生で読む。置く場所・分子・分母の選択肢はデータ次第なので、populate では値が入らない */
export function readCalcSaved(dataView: DataView | undefined): CalcSaved[] {
    return Array.from(
        { length: CALC_ROW_SLOTS },
        (_, i) => Object.fromEntries(CALC_PROPS.map((prop) => [prop, savedText(dataView, "calcRows", calcProp(prop, i + 1)) ?? ""])) as CalcSaved
    );
}

/**
 * 指標のメジャーごとの設定（書式ペインの保存値。無ければ既定：表の最後の後ろ・合計・金額・色を付けない）。
 * 横持ち（金額にメジャーを 2 本以上）では、指標の値のイベントを書式ペインで選ぶ（既定は比較順の一番大きいメジャー）
 */
export function indicatorSettings(input: InputData, eventModel: EventModel): IndicatorSetting[] {
    const horizontal = input.eventSource === "measures";
    const measures = eventModel.events.filter((e) => e.measure).map((e) => ({ value: e.measure!.queryName, displayName: e.name }));
    const strongest = eventModel.events.find((e) => e.name === (eventModel.latest[0] ?? eventModel.events.at(-1)?.name))?.measure?.queryName ?? measures.at(-1)?.value ?? "";
    return input.indicators.map((indicator) => {
        const saved = indicator.saved;
        const event = saved.event === INDICATOR_ALL_EVENTS || measures.some((m) => m.value === saved.event) ? saved.event : strongest;
        return {
            name: indicator.name,
            queryName: indicator.queryName,
            code: indicator.code,
            after: saved.after || CALC_TABLE_END,
            placement: saved.placement === INDICATOR_MODES.under || saved.placement === INDICATOR_MODES.hidden ? saved.placement : INDICATOR_MODES.after,
            aggregation: saved.aggregation === "stock" ? "stock" : "flow",
            format: saved.format ?? "",
            good: saved.good === "up" || saved.good === "down" ? saved.good : "neutral",
            event: horizontal ? event : null,
            eventChoices: horizontal ? measures : [],
        };
    });
}

/** 指標の設定から、行を組む指定を作る（置く場所の「表の最後」は null） */
export function indicatorSpecs(settings: IndicatorSetting[], warn: (message: string) => void = () => {}): IndicatorSpec[] {
    return settings.map((s) => ({
        code: s.code,
        name: s.name,
        target: s.after === CALC_TABLE_END ? null : s.after,
        mode: s.placement === INDICATOR_MODES.under ? "under" : s.placement === INDICATOR_MODES.hidden ? "hidden" : "after",
        aggregation: s.aggregation === "stock" ? "stock" : "flow",
        format: indicatorFormat(s.format, s.name, warn),
        good: s.good === "up" ? 1 : s.good === "down" ? -1 : 0,
    }));
}

/**
 * 指標の値。横持ち（イベントが "" で届く）は、書式ペインのメジャーごとのイベントに割り当てる（すべてのイベントなら、どのイベントにも
 * 同じ値）。縦持ちはイベントの列で分かれて届いたまま
 */
export function indicatorFacts(input: InputData, settings: IndicatorSetting[], eventModel: EventModel): Fact[] {
    const eventNames = eventModel.events.map((e) => e.name);
    const assigned = new Map<string, string[]>();
    for (const s of settings) {
        if (s.event === null) continue;
        const name = eventModel.events.find((e) => e.measure?.queryName === s.event)?.name;
        assigned.set(s.code, s.event === INDICATOR_ALL_EVENTS ? eventNames : name !== undefined ? [name] : []);
    }
    return input.indicatorFacts.flatMap((f) => (f.event !== "" ? [f] : (assigned.get(f.code) ?? []).map((event) => ({ ...f, event }))));
}

/** 指標の書式：空か「金額」は金額（表示単位で割る）。字の無い形（#,0）も数として出す（表示単位で割らない）。0.0% は比率 */
function indicatorFormat(text: string, name: string, warn: (message: string) => void): RowFormat {
    const trimmed = text.trim();
    if (trimmed === "" || trimmed === "値" || trimmed === "金額") return AMOUNT_FORMAT;
    const parsed = parseRowFormat(text);
    if (parsed.ok) return parsed.format.kind === "amount" ? { ...parsed.format, kind: "number" } : parsed.format;
    warn(`指標「${name}」の書式「${text}」が読めない。値の欄と同じ書式で出した`);
    return AMOUNT_FORMAT;
}

/** 書式ペインの計算行の保存値から、行を組む指定を作る（種類が「使わない」の行は入れない） */
export function calcSpecs(saved: CalcSaved[]): CalcRowSpec[] {
    return saved.flatMap((s, i): CalcRowSpec[] =>
        s.kind === "subtotal" || s.kind === "ratio"
            ? [
                  {
                      slot: i + 1,
                      kind: s.kind,
                      name: s.name.trim(),
                      from: s.from === "" || s.from === CALC_TABLE_START ? null : s.from,
                      after: s.after === "" || s.after === CALC_TABLE_END ? null : s.after,
                      numerator: s.numerator || null,
                      denominator: s.denominator || null,
                      format: s.format.trim() || null,
                      good: s.good === "down" ? -1 : s.good === "neutral" ? 0 : 1,
                  },
              ]
            : []
    );
}

const yearKey = (start: MonthIndex) => `${yearOf(start)}-${String(monthOf(start)).padStart(2, "0")}`;

function landingMessages(input: InputData): string[] {
    const missing: string[] = [];
    // 科目は無くてもよい（値をすべて足した「合計」の 1 行）。区分・貸方フラグも任意（向きが分からなければ届いた値のまま足す）
    if (!input.has.amount) missing.push("値");
    return missing;
}

/**
 * 表を組む。input は続きの読み込みを足し合わせた受け取り（visual.ts）。無ければ dataView から読む（1 回分で読み切れるとき・テスト）
 */
export function transform(
    dataView: DataView | undefined,
    settings: VisualFormattingSettingsModel,
    viewer: VisualState = EMPTY_VISUAL_STATE,
    given?: InputData,
    theme: Theme = DEFAULT_THEME
): ViewModel {
    const input = given ?? readInput(dataView);
    const style = styleOf(settings, theme);
    const ratioCap = ratioCapOf(settings);
    const unitByName = style.unitPlace === UNIT_PLACES.name;
    // 単位を数字のあとに置く（1,234百万円）。表の上には出さない
    const unitInCell = style.unitPlace === UNIT_PLACES.cell;
    const empty = (landing: string[]): ViewModel => ({
        landing,
        title: "",
        yearKey: "",
        unitCaption: "",
        mainLabel: "",
        compareLabel: "",
        periods: [],
        columns: [],
        rows: [],
        orgColumns: [],
        selectionNodes: null,
        folds: null,
        rowStateKey: "closedRows",
        picks: null,
        periodPicker: null,
        warnings: [],
        notice: null,
        truncated: input.truncated,
        format: {
            mainItems: [],
            main: "",
            measureSettings: [],
            years: [],
            fiscalYear: LATEST_YEAR,
            calcChoices: { sections: [], refs: [], places: [] },
            calcSaved: readCalcSaved(dataView),
            indicatorSettings: [],
        },
        menu: null,
        scrollStart: SCROLL_STARTS.start,
        copyButton: false,
        style,
        twoLines: false,
        indicatorSubtotals: null,
    });

    const missing = landingMessages(input);
    if (missing.length > 0) return empty(missing);

    const warnings: string[] = [];
    if (input.dropped.unreadableMonth > 0) {
        // 年だけが届くのは、日付の階層（年・四半期・月・日）のまま入れたとき
        const yearsOnly = input.dropped.unreadableMonthSamples.every((s) => /^\d{4}$/.test(s));
        const hint = yearsOnly ? "日付の階層ではなく、日付の列（または月の列）を入れる" : "日付か 2025-04 の形の列を入れる";
        warnings.push(`月が読めない行が ${input.dropped.unreadableMonth} 行あり、表に入れていない（${input.dropped.unreadableMonthSamples.join("・")}）。${hint}`);
    }
    if (input.columnsCut) {
        // 列（月の値）には続きの読み込みが無く、60 で切られる（5 年分の月）。日付の列を日ごとに入れると、60 日で届く
        warnings.push(`月の列の値が ${COLUMN_LIMIT.toLocaleString()} に届いたので、それより後の期間が届いていないおそれがある。月の列は ${COLUMN_LIMIT} まで（月ごとなら 5 年分）。日付の列なら年月の列を入れるか、ビジュアルのフィルターで期間を絞る`);
    }
    if (input.dropped.notNumber > 0) warnings.push(`値の欄に数でない値が ${input.dropped.notNumber} 個あり、表に入れていない。値のメジャーは FORMAT などで文字にしない`);
    if (input.dropped.indicatorNotNumber > 0) {
        warnings.push(`指標のメジャーが数でない値を ${input.dropped.indicatorNotNumber} 個返したので、表に入れていない。指標のメジャーは FORMAT などで文字にしない`);
    }
    // 指標の値が科目の葉にだけ届き、組織 × イベント × 月の小計では空：メジャーが科目の段でだけ値を返している（表は小計を読むので空欄になる）
    if (input.indicatorLevel.subtotals) {
        for (const indicator of input.indicators) {
            if (input.indicatorPresence.leaf.has(indicator.code) && !input.indicatorPresence.subtotal.has(indicator.code)) {
                warnings.push(
                    `指標「${indicator.name}」の値が、行の欄を外した値（セグメント × シナリオ × 月）では空。表は行の欄を外した値を読むので空欄になる。メジャーが行の段でだけ値を返していないか確かめる（ISINSCOPE の条件が逆など）`
                );
            }
        }
    }
    if (input.dropped.noAccount > 0) warnings.push(`名前の空の行が ${input.dropped.noAccount} 行あり、表に入れていない`);
    if (input.orderGuessed.length > 0) {
        const places = input.orderGuessed.map((label) => (label === ORDER_OF_SECTIONS ? "区分どうし" : `区分「${label === "" ? "（区分なし）" : label}」`));
        warnings.push(
            `セグメント・シナリオごとに行の顔ぶれが違い、届いた順だけでは並びが決めきれない（${places.slice(0, 3).join("・")}${places.length > 3 ? " ほか" : ""}）。行の並びの数値を「行の順序」の欄に入れる`
        );
    }
    if (input.dropped.noEvent > 0) warnings.push(`シナリオの空の行が ${input.dropped.noEvent} 行あり、表に入れていない`);
    if (input.duplicateMeasures.length > 0) {
        warnings.push(`値の欄に同じ名前のメジャーが 2 本以上ある（${input.duplicateMeasures.slice(0, 3).join("・")}）。1 つのシナリオとして足した。ビジュアルの欄で名前を変える`);
    }
    for (const code of input.conflicts.slice(0, 5)) warnings.push(`コード「${code}」の定義が行によって違う。最初の行の定義を使った`);

    // イベント：比較順で並べる（events.ts）。種類（実績・見通し・計画）は持たない
    const eventModel = resolveEvents(input);
    warnings.push(...eventModel.warnings);
    const { latest } = eventModel;
    // 最新見込みの見出しの決まり（書式の「最新見込みの見出し」）
    const latestRule = latestHeaderRuleOf(settings, latest);
    const eventNames = eventModel.events.map((e) => e.name);

    // 指標：書式ペインのメジャーごとの設定。横持ちは、指標の値を書式ペインで選んだイベントに割り当てる
    const indicatorSaved = indicatorSettings(input, eventModel);
    const indicators = indicatorFacts(input, indicatorSaved, eventModel);
    // 指標の値は組織 × イベント × 月の小計で読む。小計が届いていなければ visual.ts が小計を入れる設定を保存する
    const indicatorSubtotals =
        input.indicators.length > 0 && input.facts.length > 0 && !input.indicatorLevel.subtotals && input.indicatorLevel.queryName !== null
            ? { queryName: input.indicatorLevel.queryName, enabled: input.indicatorLevel.enabled }
            : null;

    // 区分の向きの多数は金額で見る（科目の金額の絶対値の和）
    const magnitude = new Map<string, number>();
    for (const fact of input.facts) magnitude.set(fact.code, (magnitude.get(fact.code) ?? 0) + Math.abs(fact.value));
    // 計算行：書式ペインで足す小計と比率
    const calcSaved = readCalcSaved(dataView);
    const built = buildRows(input.accounts, {
        magnitude: (code) => magnitude.get(code) ?? 0,
        position: dropdownValue(settings.rows.accountTotal, "bottom") === "top" ? "above" : "below",
        calcRows: calcSpecs(calcSaved),
        indicators: indicatorSpecs(indicatorSaved, (message) => warnings.push(message)),
        sectionMaster: input.has.sectionMaster,
        flags: { credit: input.has.creditFlag, balance: input.has.balanceFlag },
        // 根の合計行：自動は、科目の段が 1 つ以下（科目名だけ・科目なし）なら出す
        totalRow: ((mode) => (mode === TOTAL_ROW.auto ? input.has.accountLevels <= 1 : mode === TOTAL_ROW.on))(dropdownValue(settings.rows.totalRow, TOTAL_ROW.auto)),
        totalName: settings.rows.totalName.value.trim(),
    });
    warnings.push(...built.warnings);
    // 見る人が選んだ科目だけを出し、残りを「その他」かうちにする。表示の並びだけを組み直し、合計は変えない
    const model = applyPicks(built, viewer.picks);
    const pickChoices = pickParents(built);
    const allFacts = indicators.length > 0 ? [...input.facts, ...indicators] : input.facts;
    // 最新見込みを出すのは、比較順のあるイベントが 2 つ以上あるとき。イベントの列と金額のメジャー 2 本以上を一緒に入れたときは、
    // イベントが「実績・数量」のようにメジャーと組になるので出さない
    // 月の欄が無い表は、月ごとにシナリオを切り替えられないので出さない（全期間の値どうしを比べる）
    const latestAvailable = latest.length >= 2 && !(input.eventSource === "column" && input.measureCount > 1) && input.has.period;
    // 主の選択肢：最新見込みと、比較順のあるイベント（比較順の無いイベントは比較にだけ使う）。個々のイベントは落とさず、無い月は空欄
    const mainValues = latestAvailable ? [LATEST, ...eventModel.mainEvents] : eventModel.mainEvents;
    const validMain = (value: string | null): value is string => value !== null && mainValues.includes(value);
    // 作り手の既定（書式ペイン）。無ければ、出せれば最新見込み、出せなければ比較順の一番大きいイベント（系列があれば最初の系列）
    const savedMain = savedText(dataView, "comparison", "main");
    const authorMain = validMain(savedMain) ? savedMain : latestAvailable ? LATEST : (eventModel.strongest ?? mainValues[0] ?? "");
    // 見る人の選択（表の上のメニュー）。選んだときに作り手が書式ペインで決めていた値が、今も同じときだけ効く
    const main = viewer.main !== null && viewer.mainBase === savedMain && validMain(viewer.main) ? viewer.main : authorMain;
    // 値を取る側：イベントの候補の並び（組織 × まとまり × 月ごとに、データのある最初の候補を使う）
    const mainRef: EventRef = { events: main === LATEST ? latest : [main], shift: 0 };

    const calendar: FiscalCalendar = {
        startMonth: Number(dropdownValue(settings.periods.fiscalStartMonth, "4")) || 4,
        yearLabel: (dropdownValue(settings.periods.yearLabel, "start") === "end" ? "end" : "start") as YearLabel,
    };
    // 符号の持ち方：表全体の設定（数値のカード）。自動はシナリオごとに見分けて借方プラスにそろえる（2.0 でシナリオ・メジャーごとの設定はやめた）
    const tableSign = dropdownValue(settings.numbers.amountSign, "auto") as AmountSign;
    const signs = resolveSigns(allFacts, model.traits, eventNames, () => tableSign);
    warnings.push(...signs.warnings);
    const debitFacts = toDebitPlus(allFacts, model.traits, signs.signs);
    const calc = new Calculator(model, debitFacts);
    // 月の欄が無い表：値はすべて 1 つの月（全期間）。年度・四半期・前年同期・期間の選択は出さない
    const noPeriods = !input.has.period;
    const starts = noPeriods ? [] : Array.from(new Set(calc.months().map((m) => fiscalYearStart(m, calendar)))).sort((a, b) => b - a);
    const years = starts.map((start) => ({ value: yearKey(start), displayName: fiscalYearName(start, calendar) }));
    const savedYear = savedText(dataView, "periods", "fiscalYear");
    // 年度の既定と累計の終わりの月：比較順の一番大きいイベント（系列があれば、主と同じ系列の中で）に値のある最後の月。
    // 比較順が無ければ主の最後の月
    const seriesOf = (name: string) => eventModel.events.find((e) => e.name === name)?.series?.name;
    const strongest = latest.find((e) => seriesOf(e) === seriesOf(main));
    const lastStrong = (strongest !== undefined ? calc.lastMonth({ events: [strongest], shift: 0 }) : null) ?? calc.lastMonth(mainRef);
    const selectedStart =
        starts.find((s) => savedYear !== null && savedYear !== LATEST_YEAR && yearKey(s) === savedYear) ??
        (lastStrong !== null ? fiscalYearStart(lastStrong, calendar) : starts[0]);
    const fiscalYear = savedYear && years.some((y) => y.value === savedYear) ? savedYear : LATEST_YEAR;

    // 比較の選択肢：最新見込み（主でなければ）と、表の年度に数字のあるイベント（主のほか）と前年同期。
    // 実績のように主と同じになる月があるイベントも選べる（差が 0 の月が混ざるだけ）。
    // 系列（イベントの列と金額のメジャー 2 本以上）があれば、主と同じ系列のイベントだけ（数量を金額と比べない）
    // 表の期間：書式ペインの年度と出す列、見る人が選んでいればその列（年度をまたげる）。選んだときの書式ペインの期間の設定が
    // 今も同じときだけ効く（主・比較と同じ）
    const authorPeriods: PeriodOptions = {
        showMonths: settings.periods.showMonths.value,
        showQuarters: settings.periods.showQuarters.value,
        showHalves: settings.periods.showHalves.value,
        showYear: settings.periods.showYear.value,
        showYtd: settings.periods.showYtd.value,
    };
    // 組にするのは書式ペインの保存値（年度は「最新」のまま）。解いた年度で組にすると、スライサーやほかのビジュアルの絞り込みで年度が
    // 動いただけで見る人の選択が捨てられた。見る人の選択は年月で持つので、データが次の年度に進んでも選んだ期間のまま
    const periodBase = JSON.stringify([savedYear ?? "", calendar.startMonth, ...Object.values(authorPeriods)]);
    const periodPick = viewer.periods !== null && viewer.periods.base === periodBase ? viewer.periods : null;
    const table = noPeriods
        ? { columns: [{ key: "all", kind: "sum" as const, label: ALL_PERIODS_LABEL, months: [ALL_PERIODS_MONTH] }], starts: [], title: "", fromPick: false }
        : selectedStart === undefined
          ? null
          : tablePeriods(starts, selectedStart, authorPeriods, lastStrong, calendar, periodPick);
    // 表に出す年度の月（比較の候補と表示単位）。年度をまたげば、どちらの年度も
    const yearMonths = noPeriods ? [ALL_PERIODS_MONTH] : (table?.starts ?? []).flatMap((start) => Array.from({ length: 12 }, (_, i) => start + i));
    const inYear = new Set(eventNames.filter((e) => calc.hasData(e, yearMonths)));
    // 書式ペインの選択肢と比較の既定は、書式ペインの年度で決める（見る人が年度をまたいでも、既定の相手や作り手の選択肢を変えない）
    const authorMonths = noPeriods ? [ALL_PERIODS_MONTH] : selectedStart === undefined ? [] : Array.from({ length: 12 }, (_, i) => selectedStart + i);
    const authorInYear = new Set(eventNames.filter((e) => calc.hasData(e, authorMonths)));
    const compareValues = (against: string, pool: Set<string> = inYear) => [
        ...(latestAvailable && against !== LATEST ? [LATEST] : []),
        ...eventNames.filter((e) => e !== against && pool.has(e) && seriesOf(e) === seriesOf(against)),
        ...(noPeriods ? [] : [PRIOR_YEAR]),
    ];
    const validCompare = (value: string | null, against: string, pool: Set<string> = inYear): value is string =>
        value !== null && compareValues(against, pool).includes(value);
    // 比較の既定：比較順の一番小さい（弱い）イベント。無ければ比較順の無いイベント、前年同期
    // 作り手の既定（書式ペイン）は書式ペインの年度で、表で使う相手が表の年度に無いとき（見る人が書式ペインの年度を含まない期間を選んだ）は
    // 表の年度の候補で決め直す（書式ペインの年度で決めると、比較の列が空欄になり、上のバーの表示と食い違った）
    const defaultCompareFor = (against: string, pool: Set<string> = authorInYear) => {
        const candidates = compareValues(against, pool);
        const events = eventModel.events.filter((e) => candidates.includes(e.name));
        return (events.find((e) => e.order !== null) ?? events.at(0))?.name ?? PRIOR_YEAR;
    };
    // 表で使う既定：書式ペインの年度で決めた既定が表の年度で使えれば、それ（見る人が年度をまたいでも替わらない）。使えなければ表の年度の候補で
    // 決め直す（表の年度の候補で決めると、見る人が主を替えて年度をまたいだとき、もう一方の年度にだけあるイベントが既定になった）
    const tableDefaultFor = (against: string) => {
        const author = defaultCompareFor(against);
        return validCompare(author, against) ? author : defaultCompareFor(against, inYear);
    };
    // 比較の列：見る人が列 1〜3 ごとにシナリオと見せ方を選ぶ（表の上のバーは全期間、期間の見出しの ▾ はその期間だけ）。
    // まだ選んでいなければ既定の 1 列（比較順の一番小さいイベント）。列の並びは列 1〜3 の順
    // 比較の個々のイベントは、組織 × 月にそのイベントが無ければ、同じ系列で比較順が下のイベントへ落とす（上のイベントへは落とさない）
    const refOf = (value: string): EventRef =>
        value === PRIOR_YEAR ? { events: mainRef.events, shift: -12 } : { events: value === LATEST ? latest : fallbackOf(value, eventModel), shift: 0 };
    const label = (value: string) => (value === LATEST ? LATEST_LABEL : value === PRIOR_YEAR ? PRIOR_YEAR_LABEL : value);
    // 選択肢の名前：ビジュアルが組むもの（最新見込み・前年同期）は、データのシナリオと見分けられるようにかっこで囲む。列の見出しは囲まない
    const choiceLabel = (value: string) => (value === LATEST || value === PRIOR_YEAR ? `（${label(value)}）` : value);
    const partnerOf = (value: string): Partner => ({ compare: value, ref: refOf(value), label: label(value) });
    const viewValues = Object.values(COMPARE_VIEWS) as string[];
    const candidates = compareValues(main);
    const chosen = viewer.compares;
    // 表に出すのは見る人の列 1〜3 のうちシナリオが今選べる列（visualState.ts の決まり）。まだ選んでいなければ既定の 1 列
    const fallback = tableDefaultFor(main);
    const viewOf = (view: string): CompareView => (viewValues.includes(view) ? (view as CompareView) : DEFAULT_COMPARE_VIEW);
    const slotsAt = (period: string | null) => {
        const { slots, own } = slotsOf(chosen, period, fallback, DEFAULT_COMPARE_VIEW);
        const shown = shownSlots(slots, candidates).map((slot) => ({ ...slot, view: viewOf(slot.view) }));
        const views = slots.map(
            (slot): SlotView => ({
                compare: slot.compare,
                view: viewOf(slot.view),
                available: slot.compare !== null && candidates.includes(slot.compare),
                label: slot.compare === null ? "" : choiceLabel(slot.compare),
            })
        );
        return { shown, views, own };
    };
    const allSlots = slotsAt(null);
    // 全期間の 1 列目の比較対象（比較の名前）
    const compare = allSlots.shown[0]?.compare ?? null;

    const format: DataDrivenFormat = {
        mainItems: mainValues.map((value) => ({ value, displayName: choiceLabel(value) })),
        main: authorMain,
        measureSettings: eventModel.measureSettings,
        rowNumbers: { targets: rowNumberTargets(built), saved: readRowNumbersSaved(dataView) },
        years,
        fiscalYear,
        calcChoices: model.calcChoices,
        calcSaved,
        indicatorSettings: indicatorSaved,
        hasOrgs: input.orgLevelNames.length > 0,
        hasPeriods: !noPeriods,
    };
    const menu: EventMenu = {
        main,
        mainBase: savedMain,
        mainItems: mainValues.map((value) => ({ value, label: choiceLabel(value) })),
        compareItems: candidates.map((value) => ({ value, label: choiceLabel(value) })),
        compares: { all: allSlots.views, fallback, view: DEFAULT_COMPARE_VIEW },
    };
    if ((selectedStart === undefined && !noPeriods) || table === null) return { ...empty([]), warnings, format, menu, landing: [], indicatorSubtotals };

    const periodColumns = table.columns;

    // 率のメジャー（利益率など、% の書式）を指標として表に出すと、四半期・通期などの値は月の値の合計か期末の月の値になり、
    // 期間の率にならない（ビジュアルには月の値しか届かず、期間で計算し直せない）。月より長い列があるときだけ知らせる
    if (periodColumns.some((period) => period.months.length > 1)) {
        const rates = indicatorSaved.filter(
            (s, i) =>
                s.placement !== INDICATOR_MODES.hidden &&
                (isPercentFormat(input.indicators[i]?.formatString) || parseRowFormat(s.format).format.kind === "percent")
        );
        if (rates.length > 0) {
            warnings.push(
                `指標「${rates.map((s) => s.name).join("」「")}」は率のメジャー（% の書式）なので、四半期・通期などの列は月の値の合計か期末の月の値になり、期間の率にならない。` +
                    "分子と分母のメジャーを指標に入れて配置を「表に出さない（計算行で使う）」にし、計算行の「比率」で作ると、どの期間も分子の合計 ÷ 分母の合計になる"
            );
        }
    }

    // 期間ごとの比較の列：見る人が「この期間だけ」の列を持っていればそれ、無ければ全期間と同じ。列 1〜3 の順
    const partners = new Map(candidates.map((value): [string, Partner] => [value, partnerOf(value)]));
    const periodSlots = new Map(periodColumns.map((period) => [period.key, slotsAt(periodNameOf(period))]));
    const periodSpecs = new Map(
        periodColumns.map((period): [string, CompareColumnSpec[]] => [
            period.key,
            periodSlots.get(period.key)!.shown.map((slot) => ({ compare: slot.compare, partner: partners.get(slot.compare)!, view: slot.view, slot: slot.index })),
        ])
    );
    const specsAt = (key: string) => periodSpecs.get(key) ?? [];
    // どこかの期間で使う比較対象（組織のブロックの名乗りを調べる）
    const usedCompares = candidates.filter((value) => Array.from(periodSpecs.values()).some((specs) => specs.some((spec) => spec.compare === value)));
    const compareLabel = compare === null ? "" : label(compare);
    const swap = settings.comparison.diffSwap.value;
    const toneMode = dropdownValue(settings.comparison.toneMode, TONE_MODES.both);
    const subtotalsBelow = dropdownValue(settings.rows.accountTotal, "bottom") !== "top";

    // 表示単位：金額の行の、年度の通期の値（主と比較）の最大で、表全体で 1 つに決める。
    // 列ごとに単位が動くと横に比べられない。出す列の出し入れで単位が変わらないよう、出している列には寄らない
    // 見る人が科目を選んでも単位が変わらないよう、組み直す前の行で決める
    const valueRows = built.display.filter((d) => !["heading", "blank"].includes(d.def.type));
    // 選べる比較対象を全部含める（見る人が比較の列を出し入れしても単位が変わらない）
    const unitRefs = [mainRef, ...candidates.map(refOf)];
    // 行別の数値書式（書式ペインの枠）。対象は段か行で、表全体 → 段 → 上の行から自分の行の順に重ねる（rowNumbers.ts）
    const rowNumberSaved = readRowNumbersSaved(dataView);
    const overrides = rowNumberOverrides(
        rowNumberSaved,
        rowNumberTargets(built).map((t) => t.value),
        (message) => warnings.push(message)
    );
    // 組み直す前の行（単位を決める。見る人がその他・うちにして表から消えた科目も入れる。入れないと、科目を選ぶだけで表全体の単位が
    // 変わった）と、組み直したあとの行（見る人がまとめた「その他」・小計のうち）の両方で
    const rowOverride = overrides.size > 0 ? new Map([...resolveRowNumbers(built, overrides), ...resolveRowNumbers(model, overrides)]) : new Map<string, RowNumberOverride>();
    /** 行の書式の上書き（#,0h など）。読めなければ知らせて金額のまま */
    const formatCache = new Map<string, RowFormat | null>();
    const ownFormatOf = (code: string): RowFormat | null => {
        const text = rowOverride.get(code)?.format?.trim();
        if (!text) return null;
        if (!formatCache.has(text)) {
            const parsed = parseRowFormat(text);
            if (!parsed.ok) warnings.push(`行別の数値書式の書式「${text}」が読めない。金額のまま出した`);
            formatCache.set(text, parsed.ok ? (parsed.format.kind === "amount" ? { ...parsed.format, kind: "number" } : parsed.format) : null);
        }
        return formatCache.get(text) ?? null;
    };
    // 表の単位は、固定の単位か書式で上書きした行を除いて決める（除かないと、千円で見せる区分の大きさで表全体の単位が決まる）。
    // 上書きごとの最大は、単位だけ変えて桁を表全体（自動）に合わせるときの桁に使う
    const overrideMax = new Map<string, number>();
    let maxAbs = 0;
    for (const { def } of valueRows) {
        if (def.format.kind !== "amount" || ownFormatOf(def.code)) continue;
        const override = rowOverride.get(def.code);
        const ownUnit = override?.unitType !== undefined ? overrideKey(override) : null;
        for (const ref of unitRefs) {
            for (const months of [...table.starts.map((start) => Array.from({ length: 12 }, (_, i) => start + i)), ...yearMonths.map((m) => [m])]) {
                const v = calc.aggregate(def.code, ref, months);
                if (v === null) continue;
                if (ownUnit !== null) overrideMax.set(ownUnit, Math.max(overrideMax.get(ownUnit) ?? 0, Math.abs(v)));
                else maxAbs = Math.max(maxAbs, Math.abs(v));
            }
        }
    }
    const unitType = dropdownValue(settings.numbers.unitType, "auto");
    const notation = dropdownValue(settings.numbers.unitNotation, "japanese");
    const precision = dropdownValue(settings.numbers.precision, "auto");
    const unitDef = resolveUnit(unitType === "auto" ? tableAutoUnitKey(maxAbs, notation) : unitType, maxAbs, notation, precision);
    const unit: AmountUnit = {
        divisor: unitDef.divisor,
        decimals: precision === "auto" ? (maxAbs / unitDef.divisor < 100 && unitDef.divisor > 1 ? 1 : 0) : Math.max(0, parseInt(precision, 10) || 0),
    };
    const currency = settings.numbers.currency.value?.trim() || DEFAULT_CURRENCY;
    // 表の上の単位は、表全体の単位で出す金額の行があるときだけ（どの行も単位か書式を上書きし、計算行も無ければ出さない）
    const hasAmountRows = valueRows.some(
        (d) => d.def.format.kind === "amount" && !ownFormatOf(d.def.code) && rowOverride.get(d.def.code)?.unitType === undefined && rowOverride.get(d.def.code)?.currency === undefined
    );
    const mixedUnits = valueRows.some((d) => d.def.format.kind !== "amount" || ownFormatOf(d.def.code) !== null);

    const sign: SignOptions = {
        negative: dropdownValue(settings.numbers.negativeStyle, NEGATIVE_STYLES.triangle),
        zero: dropdownValue(settings.numbers.zeroStyle, ZERO_STYLES.zero),
        diffZero: dropdownValue(settings.numbers.diffZeroStyle, ZERO_STYLES.plusMinus),
        negativeZero: settings.numbers.negativeZero.value,
    };
    interface RowNumbers {
        unit: AmountUnit;
        sign: SignOptions;
        /** 表全体と違う単位の行の名前に添える（（千円））。添えるのは、親と単位の違う行だけ（rowSuffix） */
        suffix: string;
        word: string;
        currency: string;
        format: RowFormat | null;
    }
    /** 行の数値の設定：上書きしていればその設定、ほかは表全体 */
    const numbersByKey = new Map<string, RowNumbers>();
    const tableNumbers: RowNumbers = { unit, sign, suffix: "", word: unitDef.unitWord, currency, format: null };
    const numbersOf = (code: string): RowNumbers => {
        const override = rowOverride.get(code);
        if (!override) return tableNumbers;
        const key = overrideKey(override);
        const known = numbersByKey.get(key);
        if (known) return known;
        const own = override.unitType !== undefined;
        const ownMax = overrideMax.get(key) ?? 0;
        const def = own ? resolveUnit(override.unitType!, ownMax, notation, precision) : unitDef;
        const decimals =
            override.precision !== undefined
                ? Math.max(0, parseInt(override.precision, 10) || 0)
                : own && precision === "auto"
                  ? ownMax / def.divisor < 100 && def.divisor > 1
                      ? 1
                      : 0
                  : unit.decimals;
        const rowCurrency = override.currency?.trim() || currency;
        const value: RowNumbers = {
            unit: { divisor: def.divisor, decimals },
            sign: {
                negative: override.negativeStyle ?? sign.negative,
                zero: override.zeroStyle ?? sign.zero,
                diffZero: override.diffZeroStyle ?? sign.diffZero,
                negativeZero: override.negativeZero !== undefined ? override.negativeZero === "on" : sign.negativeZero,
            } as SignOptions,
            // 表全体と違う単位・通貨の行に添える。表の上の単位を出していない（どの金額の行も単位を上書きした）なら、同じ単位の行にも添える
            // （添えないと、単位がどこにも出ない表になった）
            suffix: (own || rowCurrency !== currency) && (def.divisor !== unitDef.divisor || rowCurrency !== currency || !hasAmountRows) ? `（${def.unitWord}${rowCurrency}）` : "",
            word: def.unitWord,
            currency: rowCurrency,
            format: ownFormatOf(code),
        };
        numbersByKey.set(key, value);
        return value;
    };
    /** 単位を名前に添える行：親と単位・通貨の違う行（親の行の名前に添えてあれば、子には添えない） */
    const parentOfRow = parentsOf(model);
    const rowSuffix = (code: string): string => {
        const own = numbersOf(code);
        if (!own.suffix) return "";
        const parent = parentOfRow.get(code);
        const above = parent !== undefined ? numbersOf(parent) : null;
        return above && above.suffix === own.suffix ? "" : own.suffix;
    };

    // 見出しは 2 行。2 段の列はセルの上下と同じ順に「修正予算差」「修正予算比」
    const columns: ValueColumn[] = [];
    const periods: PeriodHeader[] = [];
    const titled = (items: TooltipItem[] | undefined) => {
        const title = items?.map((item) => `${item.displayName}=${item.value}`).join(" / ");
        return title && items ? { title, tooltip: items } : {};
    };
    for (const period of periodColumns) {
        const specs = specsAt(period.key);
        const here = periodSlots.get(period.key)!;
        periods.push({ key: period.key, label: period.label, kind: period.kind, span: 1 + specs.length, months: period.months, slots: here.views, own: here.own });
        // 最新見込みの主は、期間に使ったイベントが 1 つならその名前、混ざれば「最新見込み」を名乗る
        const mainUsed = main === LATEST ? calc.eventsByMonth(mainRef, period.months) : null;
        // データの無い期間は「―」（空の見出しだと、列がずれたように見える）
        const mainHeader = mainUsed ? latestName(mainUsed, latestRule) || "―" : main;
        columns.push({ periodKey: period.key, periodKind: period.kind, kind: "main", header: mainHeader, sub: "", ...titled(breakdownItems(mainUsed)) });
        for (const spec of specs) {
            const at = spec.partner;
            // 比較の月ごとの内訳：下のイベントへ落とした月があるとき（期間を通して落としたときも）と、最新見込み・前年同期でイベントが混ざるとき。
            // 差の列も比較の名前を名乗るので、比較の値の列を出していなくても内訳が分かるよう、同じツールチップを付ける
            const used = at.ref.events.length > 1 ? calc.compareEventsByMonth(mainRef, at.ref, period.months) : null;
            const items = at.compare === PRIOR_YEAR || at.compare === LATEST ? breakdownItems(used) : breakdownItems(used, at.compare);
            const [header, sub] = compareHeaderLines(at.label, spec.view, swap);
            columns.push({
                periodKey: period.key,
                periodKind: period.kind,
                kind: spec.view === COMPARE_VIEWS.value ? "compare" : "diff",
                compare: spec.compare,
                view: spec.view,
                slot: spec.slot,
                periodName: periodNameOf(period),
                header,
                sub,
                ...titled(items),
            });
        }
    }

    // 組織のブロックの木。組織の欄が無ければ null（表は 1 つ）
    const tree = buildOrgTree(input.orgs, input.orgLevelNames.length, settings.segments.rootName.value.trim() || DEFAULT_ROOT_NAME);
    format.hasRoot = tree?.root.path === null;
    const hierarchy = accountTree(model, settings.rows.stepParents.value);
    const hierarchySaved = readHierarchyOverrides(dataView.metadata.objects);
    const hierarchyTargets: powerbi.IEnumMember[] = [];
    const hierarchyLevels = { account: new Set<number>(), org: new Set<number>() };
    const accountTargets = (nodes: AccountNode[], path: string[]) => nodes.forEach(n => {
        const names = [...path, n.row.def.name];
        // 合計行は、うちだけ（見る人が区分をうちにした）でも対象にする
        if (n.children.length || (n.row.def.code === TOTAL_KEY && n.following.length)) {
            hierarchyLevels.account.add(path.length);
            hierarchyTargets.push({ value: `account:node:${n.row.def.code}`, displayName: `行：${names.join(" / ")}` });
        }
        accountTargets(n.children, names);
    });
    if (hierarchy) {
        accountTargets(hierarchy.roots, []);
        warnings.push(...hierarchy.warnings);
    }
    if (tree) for (const node of allNodes(tree.root)) {
        hierarchyLevels.org.add(node.column + (tree.root.column < 0 ? 1 : 0));
        hierarchyTargets.push({ value: `org:node:${node.path ?? ROOT_FOLD}`, displayName: `セグメント：${node.path?.split(ORG_SEPARATOR).join(" / ") ?? node.label}` });
    }
    hierarchyTargets.unshift(...(["account", "org"] as const).flatMap(kind => Array.from(hierarchyLevels[kind]).sort((a, b) => a - b)
        .map(level => ({ value: `${kind}:level:${level}`, displayName: `${kind === "account" ? "行" : "セグメント"} 表示階層 ${level + 1}` }))));
    format.hierarchy = { targets: hierarchyTargets, saved: hierarchySaved };
    // 折りたたみ：閉じた行の下の行を隠す（どの組織のブロックでも同じ）。囲みの帯は見えている行で引き直す。
    // 組織のブロックがあれば、行は既定で閉じる。組織の欄が無ければ既定で開く。見る人の保存は既定から切り替えた行
    // セグメントのブロックがあれば、行の開き閉じはブロックごとに覚える（保存はブロックの道筋とコードの組 = 行のキー）。1 つのブロックで科目を
    // 開いても、ほかのブロックは開かない
    const foldableCodes = hierarchy?.parents ?? model.display.filter(foldable).map((d) => d.def.code);
    const rowsClosedByDefault = tree !== null;
    const openRows = new Set(viewer.openRows);
    const foldKey = (prefix: string, code: string) => (tree ? `${prefix}\u0001${code}` : code);
    // その他・小計をうちにした行は、既定で閉じる。行を既定で開く表（セグメントなし）では、見る人の「閉じた行」の記録を反転して読む（押すと開く）
    const closedByDefault = new Set(foldableCodes.filter((code) => code.startsWith(OTHERS_KEY) || code.startsWith(UNDER_KEY)));
    const closedRowsSet = new Set(viewer.closedRows);
    const closedIn = (prefix: string) =>
        new Set(rowsClosedByDefault
            ? foldableCodes.filter((code) => !openRows.has(foldKey(prefix, code)))
            : foldableCodes.filter((code) => closedRowsSet.has(code) !== closedByDefault.has(code)).concat(viewer.closedRows.filter((code) => !foldableCodes.includes(code))));
    const shownIn = (closedSet: Set<string>) => {
        const hiddenSet = hiddenRows(model.display, model.attached, closedSet, model.following);
        return model.display.filter((d) => !hiddenSet.has(d.def.code));
    };
    const closed = closedIn("");
    const shown = shownIn(closed);
    const bands = shown.map((): Band[] => []);
    /**
     * 表の行（1 つの組織のブロック、組織の欄が無ければ表全体）。org はツールチップに出す組織の名前。
     * 組織のブロックの最新見込みは、その組織で使ったイベントを名乗る（列の見出しは全社で使ったイベント。事業B に見通しが無い月は実績）。
     * 比較も、下のイベントへ落としたり最新見込みで混ざったりしたら、その組織で使ったイベントを名乗る（事業B の 4 月の修正予算は期初予算。
     * ）。期間の中で混ざれば、比較の最新見込みは「最新見込み」、個々のイベントは使ったイベントを比較順の候補の並びで
     * 並べる（修正予算・期初予算）
     */
    /**
     * 行を押したときに絞る科目：科目の行はその科目、区分・中分類の小計・見出しはその下の科目すべて、段階の行（営業利益・合計行など）は足す区分の科目、
     * 計算行は式が引く行の科目。科目の無い行（指標だけを引く計算行・指標）は選べない。
     * 段階の行・計算行も選べないと、押すと空いた所を押したことになって選択が解け、ほかの組織のブロックまで明るくなった（2026-10-08 ユーザー）
     */
    const selectCodes = new Map<string, string[] | null>();
    const codesOf = (code: string): string[] | null => {
        const known = selectCodes.get(code);
        if (known !== undefined) return known;
        const def = model.rows.get(code);
        let result: string[] | null = null;
        if (isAccountRow(def)) result = [code];
        // その他はまとめた科目
        // その他はまとめた科目。小計をうちにした行は、その小計の科目すべて
        else if (def?.type === "others") result = def.summands.flatMap((s) => codesOf(s.code) ?? [s.code]);
        else if (def?.type === "breakdown" && def.summands.length > 0) result = def.summands.flatMap((s) => codesOf(s.code) ?? []);
        else if (def && (def.type === "subtotal" || def.type === "heading")) {
            // 子と足す行（見る人がその他・うちにして子から外した科目も、親の合計には入っている）
            const codes = Array.from(new Set([...def.children, ...def.summands.map((s) => s.code)].flatMap((child) => codesOf(child) ?? [])));
            result = codes.length > 0 ? codes : null;
        } else if (def && (def.type === "step" || def.type === "calc")) {
            selectCodes.set(code, null);
            const codes = Array.from(new Set([...def.summands.map((s) => s.code), ...def.refs].flatMap((ref) => codesOf(ref) ?? [])));
            result = codes.length > 0 ? codes : null;
        }
        selectCodes.set(code, result);
        return result;
    };
    const tableRows = (
        calc: Calculator,
        prefix: string,
        org: string | null,
        whole: boolean,
        rowsIn: DisplayRow[] = shown,
        bandsIn: Band[][] = bands,
        orgPath: string | null = null,
        has: (code: string) => boolean = () => true,
        closedSet: Set<string> = closed
    ): TableRow[] => {
        // 「その他」にまとめた行が小計（中分類）なら、その下の行にデータがあるか（小計そのものはファクトを持たない）
        const hasRow = (code: string): boolean => has(code) || (model.rows.get(code)?.summands ?? []).some((s) => s.code !== code && hasRow(s.code));
        const ownHeaders = new Map(
            periodColumns.map((period): [string, string | null] => [period.key, !whole && main === LATEST ? latestName(calc.eventsByMonth(mainRef, period.months), latestRule) || null : null])
        );
        const ownCompares = new Map(
            usedCompares.map((value): [string, Map<string, string | null>] => [
                value,
                new Map(
                    periodColumns.map((period): [string, string | null] => {
                        const at = partners.get(value)!;
                        if (whole || at.compare === PRIOR_YEAR || at.ref.events.length <= 1) return [period.key, null];
                        const used = new Set(Array.from(calc.compareEventsByMonth(mainRef, at.ref, period.months).values()).flat());
                        if (used.size <= 1) return [period.key, used.size === 1 ? Array.from(used)[0] : null];
                        return [period.key, at.compare === LATEST ? latestName(used, latestRule) : at.ref.events.filter((e) => used.has(e)).join("・")];
                    })
                ),
            ])
        );
        return rowsIn.map((row, index) => {
            const { depth } = row;
            const { unit: rowUnit, sign: rowSign, word, currency: rowCurrency, format: ownFormat } = numbersOf(row.def.code);
            const suffix = rowSuffix(row.def.code);
            const plain = othersNamed(row.def, hasRow, settings.rows.othersCount.value);
            // 行別の数値書式で書式を上書きした行（#,0h）は、その書式で出す（金額の行だけ）
            const formatted = ownFormat && plain.format.kind === "amount" ? { ...plain, format: ownFormat } : plain;
            // 単位を数字のあとに置くとき、金額の行は数字に単位の字を付ける（ツールチップは円まで出すので付けない）
            const named = unitInCell && formatted.format.kind === "amount" ? { ...formatted, format: { ...formatted.format, suffix: `${word}${rowCurrency}` } } : formatted;
            const blankRow = named.type === "heading" || named.type === "blank";
            // 単位を行の名前の横に置くとき：値の行は（百万円）、単位の字のある指標は数字から字を外して（台）。比率の % は数字に残す
            const nameUnit = !unitByName || blankRow ? "" : named.format.kind === "amount" ? `（${word}${rowCurrency}）` : named.format.kind === "number" && named.format.suffix.trim() ? `（${named.format.suffix.trim()}）` : "";
            // ツールチップの正確な値は単位の字を付けたまま（def.format）。表のセルだけ外す
            const def = nameUnit && named.format.kind === "number" ? { ...named, format: { ...named.format, suffix: "" } } : named;
            const exact = formatted.format;
            // 表全体と違う単位の区分は、区分の行（一番上の段）の名前に単位を添える
            // ツールチップは名前だけ（値は円まで出すので、単位の添えは付けない）
            const shownName = unitByName || unitInCell ? named.name : suffix ? named.name + suffix : named.name;
            const cells: Cell[] = [];
            for (const period of periodColumns) {
                if (blankRow) {
                    columns.filter((c) => c.periodKey === period.key).forEach(() => cells.push({ text: "", tone: null }));
                    continue;
                }
                const mainValue = calc.aggregate(def.code, mainRef, period.months);
                const specs = specsAt(period.key);
                const results = specs.map((spec) => calc.compare(def.code, mainRef, spec.partner.ref, period.months));
                // ツールチップ：表は単位で丸めるので、その期間の主・比較・差・比の正確な値をまとめて出す（どのセルでも同じ）。
                // 指標の行はメジャー 1 本ずつのまとまりでイベントを採るので、列の見出しと違うイベントの値のことがある。そのイベントを名乗る
                // （見通しに人数が無い月の人数は修正予算）
                const columnMain = ownHeaders.get(period.key) ?? columns.find((c) => c.periodKey === period.key && c.kind === "main")?.header ?? "";
                // 混ざったときの名乗りは列の見出しと同じ：主と最新見込みの相手は「最新見込み」、個々のイベントの相手は使ったイベントを候補の並びで
                // （見通し・修正予算。どちらも「最新見込み」を名乗ると、比較の列が複数のときツールチップで見分けられない）
                const ownName = (ref: EventRef, fallback: string, latestLike = true) => {
                    if (!def.code.startsWith(INDICATOR_KEY) || ref.events.length <= 1) return fallback;
                    const used = calc.rowEventsByMonth(def.code, ref, period.months);
                    if (latestLike) return latestName(used, latestRule) || fallback;
                    const all = new Set(Array.from(used.values()).flat());
                    return all.size === 0 ? fallback : ref.events.filter((e) => all.has(e)).join("・");
                };
                const mainName = ownName(mainRef, columnMain);
                // 比較の相手ごとに、比較の値・差と、比か率（その相手の列が率を出していれば率、比を出していれば比、どちらも無ければ比）。
                // 同じ名前で同じ値の相手（同じシナリオを 2 つの列に選んだとき、組織のブロックで落とした先が同じになった相手）は 1 度だけ。
                // 名前だけでまとめない：別々の相手が期間の中でイベントが混ざって同じ「最新見込み」を名乗っても、値が違えば両方出す
                const partners: TooltipItem[] = [];
                const keys = specs.map((spec, i) => {
                    const at = spec.partner;
                    const name = at.compare === PRIOR_YEAR ? at.label : ownName(at.ref, ownCompares.get(spec.compare)?.get(period.key) ?? at.label, at.compare === LATEST);
                    return { name, key: `${name}\u0001${results[i].compare}` };
                });
                specs.forEach((_, i) => {
                    const { name, key } = keys[i];
                    if (keys.findIndex((k) => k.key === key) !== i) return;
                    const result = results[i];
                    const views = specs.filter((_, j) => keys[j].key === key).map((o) => o.view);
                    const rate = views.some((v) => v === COMPARE_VIEWS.rate || v === COMPARE_VIEWS.diffRate);
                    const ratio = views.some((v) => v === COMPARE_VIEWS.ratio || v === COMPARE_VIEWS.diffRatio) || !rate;
                    partners.push(
                        { displayName: name, value: formatExact(result.compare, exact, rowCurrency, rowSign) },
                        { displayName: `${name}差`, value: formatExact(result.diff, exact, rowCurrency, rowSign, true) }
                    );
                    // 比率の行の比・率は意味が無い（差のポイントだけ）
                    if (def.format.kind === "percent") return;
                    if (ratio) partners.push({ displayName: `${name}比`, value: formatRatio(result.main, result.compare, 1 + EXACT_EXTRA_DECIMALS, ratioCap) });
                    if (rate) partners.push({ displayName: `${name}率`, value: formatRate(result.main, result.compare, rowSign, 1 + EXACT_EXTRA_DECIMALS, ratioCap) });
                });
                const values: TooltipItem[] = [{ displayName: mainName, value: formatExact(mainValue, exact, rowCurrency, rowSign) }, ...partners].filter(
                    (item) => item.value !== ""
                );
                const tooltip: TooltipItem[] | undefined =
                    mainValue === null && results.every((result) => result.compare === null)
                        ? undefined
                        : [
                              ...(org !== null ? [{ displayName: "セグメント", value: org }] : []),
                              { displayName: "行", value: def.name },
                              ...(def.type === "others" ? [{ displayName: "まとめた行", value: othersList(def, hasRow, (code) => model.rows.get(code)?.name ?? code) }] : []),
                              { displayName: "期間", value: period.label },
                              ...values,
                          ];
                const withTooltip = (cell: Cell): Cell => (tooltip ? { ...cell, tooltip } : cell);
                cells.push(withTooltip({ text: formatRowValue(mainValue, def.format, rowUnit, rowSign, ratioCap), tone: null }));
                specs.forEach((spec, i) => {
                    const result = results[i];
                    if (spec.view === COMPARE_VIEWS.value) {
                        cells.push(withTooltip({ text: formatRowValue(result.compare, def.format, rowUnit, rowSign, ratioCap), tone: null }));
                        return;
                    }
                    // 色は表に出る差の符号で決める（丸めて 0 は塗らない、▲0 は悪い向きなら塗る）。比・率だけの列も差の向きで塗る
                    const tone = toneOf(diffSign(result.diff, def.format, rowUnit, rowSign), def.good, toneMode);
                    cells.push(
                        withTooltip(
                            diffCell(result.main, result.compare, tone, def.format.kind === "percent", spec.view, swap, rowSign, ratioCap, () =>
                                formatDiff(result.diff, def.format, rowUnit, rowSign, ratioCap)
                            )
                        )
                    );
                });
            }
            return {
                key: `${prefix}\u0001${def.code}`,
                code: def.code,
                name: shownName,
                ...(nameUnit ? { unit: nameUnit } : {}),
                type: def.type,
                ...(def.continued ? { continued: true } : {}),
                depth,
                bands: bandsIn[index],
                cells: cells.map(cell => ({ ...cell, parts: valueParts(cell.text, [def.format.suffix, `%${def.format.suffix}`, `pt${def.format.suffix}`, "%", "pt"]),
                    ...(cell.sub !== undefined ? { subParts: valueParts(cell.sub, ["%", "pt"]) } : {}) })),
                ...(foldable(row) ? { open: !closedSet.has(def.code), fold: foldKey(prefix, def.code), ...(subtotalsBelow ? { toggleAbove: true } : {}) } : {}),
                ...(codesOf(def.code) ? { select: { codes: codesOf(def.code)!, org: orgPath } } : {}),
                orgPath,
            };
        });
    };

    // 組織のブロック：開いた組織は子のブロックと計のブロック、閉じた組織は計の表（子を隠す）。ブロックの値はその組織より下のファクトで
    const rows: TableRow[] = [];
    let layout: LayoutPlan | undefined;
    {
        const options = (kind: "account" | "org", id: string, level: number) => hierarchyOptions({ ...settings.rows, ...settings.segments } as unknown as HierarchyDefaults, hierarchySaved, kind, id, level);
        const codes = tree && settings.segments.hideEmptyAccounts.value ? codesByOrg(tree.root, input.facts) : null;
        // 表に出す月（前年同期を比べるなら 12 か月前も）。「値の無い行を隠す」「0 だけの行を隠す」がどちらも切なら null
        const hideZero = settings.rows.hideZeroRows.value ?? false;
        const shownMonths = settings.rows.hideBlankRows.value || hideZero
            ? new Set(periodColumns.flatMap(p => usedCompares.includes(PRIOR_YEAR) ? [...p.months, ...p.months.map(m => m - 12)] : p.months))
            : null;
        const shownRefs = [mainRef, ...usedCompares.map(refOf)];
        const block = (node?: OrgNode): LayoutPlan => {
            const prefix = node ? node.path ?? "\u0000" : "";
            const ownCalc = !node || node === tree?.root ? calc : new Calculator(model, debitFacts.filter(f => containsOrg(node, f.org)));
            const present = node && codes ? codes.get(node.path) ?? new Set<string>() : null;
            const empty = present ? emptyAccountRows(model.display, model.attached, code => present.has(code), model.rows) : new Set<string>();
            // 「行」カードの「値の無い行を隠す」：表に出す月にファクトの無い行を外す（区分・中分類・その他・うちはセグメントの空の行と同じ扱い）
            // 「0 だけの行を隠す」：0 の値は無い値とみなす（科目も指標も）
            if (shownMonths) {
                const counted = (f: Fact) => shownMonths.has(f.month) && (!node || containsOrg(node, f.org)) && (!hideZero || f.value !== 0 || f.last !== 0);
                const inShown = new Set([...input.facts, ...input.indicatorFacts].filter(counted).map(f => f.code));
                // 表に出したシナリオ（基準と、使っている比較）の、月ごとに採ったイベントの数字だけを数える。月だけで見ると、表に出していない
                // シナリオ（期初予算など）にだけ値のある行が残った（2026-10-08 ユーザーが会社で見つけた）
                const valued = (code: string) => ownCalc.hasValue(code, shownRefs, shownMonths, hideZero);
                for (const code of emptyAccountRows(model.display, model.attached, code => inShown.has(code) && valued(code), model.rows, true)) empty.add(code);
            }
            const shown = model.display.filter(r => !empty.has(r.def.code));
            const full = shown.length ? shown : model.display;
            const values = tableRows(ownCalc, prefix, node?.label ?? null, !node || node === tree?.root, full,
                full.map((): Band[] => []), node?.path ?? null, present ? code => present.has(code) : undefined, new Set());
            return accountPlan(hierarchy.roots, new Map(values.map(r => [r.code, r])), closedIn(prefix),
                (n, level) => options("account", n.row.def.code, level), code => foldKey(prefix, code), `accounts:${prefix}`, String(settings.rows.accountRoot.value.value));
        };
        const orgPlan = (node: OrgNode, level: number): LayoutPlan => {
            const id = node.path ?? ROOT_FOLD;
            const fold = node === tree!.root ? ROOT_FOLD : id;
            const open = node.children.length > 0 && (node === tree!.root ? !viewer.openOrgs.includes(ROOT_FOLD) : viewer.openOrgs.includes(id));
            const opt = options("org", id, level);
            const header = { key: `org:${id}`, kind: "org" as const, label: node.label, org: node.path,
                ...(node.children.length ? { open, fold } : {}) };
            let content = block(node);
            if (open) {
                const children = stack(`${id}:children`, node.children.map(n => orgPlan(n, level + 1)), opt.style === "split" ? opt.direction : "vertical");
                const aggregate: LayoutPlan = { kind: "frame", key: `${id}:total`, style: opt.style === "split" ? "split" : "plain",
                    header: opt.style === "split" ? { ...header, above: opt.total === "bottom" }
                        : { ...header, key: `${header.key}:total`, open: undefined, fold: undefined }, content };
                // 横積みは親の名前のセルが子の行をまたぐので、開いたときは合計の表を出さない（閉じれば合計の表。科目の横積みと同じ）
                const total = opt.style === "columns" ? "none" : opt.total;
                content = stack(`${id}:content`, total === "none" ? [children] : total === "top" ? [aggregate, children] : [children, aggregate]);
                // 集計帳票がある分割では、その見出しに開閉操作をまとめる。
                if (opt.style === "split" && opt.total !== "none") {
                    aggregate.header = { ...header, above: opt.total === "bottom" };
                    return { ...content, preserveTables: true } as LayoutPlan;
                }
            }
            return { kind: "frame", key: `org:${id}`, style: opt.style, header, content };
        };
        layout = tree ? compactOrgBoxes(orgPlan(tree.root, 0)) : block();
        rows.push(...planRows(layout));
    }

    return {
        landing: [],
        title: table.title,
        yearKey: table.starts.map(yearKey).join(","),
        // 単位を行の名前の横に置くときは、表の上には出さない
        // 金額でない行（率・時間・人数など）が混ざる表は「金額：」（「単位：」だと率や時間にも掛かって読める）
        unitCaption: hasAmountRows && !unitByName && !unitInCell ? `${mixedUnits ? "金額" : "単位"}：${unitDef.unitWord}${currency}` : "",
        mainLabel: label(main),
        compareLabel,
        periods,
        columns,
        rows,
        ...(layout ? { layout } : {}),
        // セグメントは 1 列。見出しは段の名前を並べる（事業部・事業）
        orgColumns: tree && !layout ? [input.orgLevelNames.slice(0, tree.columns).join("・")] : [],
        selectionNodes: input.selection,
        // すべて開くは、どのブロックのどの行も開く（行はブロックごとに覚える）
        folds: foldsOf(
            tree ? toggleablePaths(tree.root) : [],
            tree ? allNodes(tree.root).flatMap((node) => foldableCodes.map((code) => foldKey(node.path ?? "\u0000", code))) : foldableCodes,
            rowsClosedByDefault,
            closedByDefault
        ),
        rowStateKey: rowsClosedByDefault ? "openRows" : "closedRows",
        picks: pickChoices.length > 0 ? { parents: pickChoices, current: viewer.picks } : null,
        periodPicker: noPeriods ? null : {
            years: periodChoices(starts, lastStrong, calendar),
            current: periodColumns.filter((c) => c.kind !== "sum").map(periodIdOf),
            // 合計の列のチェックは保存値で示す（月の列が無くて合計の列が出ていなくても）
            total: table.fromPick ? periodPick!.total : false,
            picked: table.fromPick,
            base: periodBase,
        },
        warnings,
        notice: null,
        truncated: input.truncated,
        format,
        menu,
        scrollStart: dropdownValue(settings.periods.scrollStart, SCROLL_STARTS.start),
        copyButton: settings.table.copyButton.value,
        style,
        twoLines: rows.some((row) => row.cells.some((cell) => cell.sub !== undefined)),
        ...(rows.length > 0 && rows.every((row) => row.code === TOTAL_ACCOUNT) && dropdownValue(settings.rows.totalRow, TOTAL_ROW.auto) === TOTAL_ROW.off
            ? { hideNames: true }
            : {}),
        indicatorSubtotals,
    };
}

/** 見た目の値（書式ペイン）。文字サイズ（pt）は 6〜40 に収め、色は空なら既定 */
export interface TableStyle {
    fontFamily: string;
    /** 表の上のバーとダイアログの文字サイズ */
    fontSize: number;
    /** 表のすぐ上の年度と単位の文字の色。空なら表の文字と同じ */
    headingColor: string;
    /** 表の中の要素ごとの文字サイズ */
    sizes: typeof DEFAULT_TEXT_SIZES;
    good: string;
    bad: string;
    /** 段ごとの囲みの色（外側から）。深い段は最後の色 */
    bandColors: string[];
    headBackground: string;
    /** 見出しの字の色（背景の明るさで濃い字か白）と、薄い字（列の見出しの 2 行目・単位・組織の列の見出し）の色 */
    headText: string;
    headMuted: string;
    /** 字・副文字・地（テーマの色） */
    text: string;
    muted: string;
    background: string;
    lines: typeof DEFAULT_LINE_COLORS;
    /** 期間のあいだに縦の線を引くか */
    periodLines: boolean;
    /** 集計の列（四半期・半期・通期・累計・合計）の前の区切りを二重線にするか */
    aggregateDouble: boolean;
    /** 「単位：百万円」を置く所 */
    unitPlace: UnitPlace;
    /** 年度（表の題）を左上の角に置くか（既定は表の上） */
    titleInCorner: boolean;
    /** 囲み（区分・中分類の箱）を塗るか、セグメントの箱を塗るか、セグメントの段ごとの色（外側から） */
    bandFill: boolean;
    /** 区分・中分類の行の数字も太字にするか（既定は計算行の数字だけ太字） */
    boldAggregates: boolean;
    /** 行の高さ（「表全体」カード） */
    rowHeight: RowHeight;
    /** 数字のあとの単位の字の幅を列の中でそろえるか（「数値」カード） */
    alignTails: boolean;
    /** 文字の要素ごとの太さ（各カードの「…の太字」） */
    bold: Record<BoldKey, BoldMode>;
    /** 列の余白（「表全体」カード） */
    columnPadding: ColumnPadding;
    /** うちの行の見せ方：基準の数字をかっこで囲む・灰色にする・「うち」の字を付ける（「行」カードの「うちの行」） */
    breakdown: { brackets: boolean; muted: boolean; tag: boolean };
    segmentFill: boolean;
    segmentColors: string[];
    /** 列の見出しの文字：セグメントの列（横積み）と、行の名前の列 */
    headers: { segment: string; row: string };
    /** 列の見出し（セグメント・行の名前の列）を太字にするか */
    headerBold: { segment: boolean; row: boolean };
}

export type UnitPlace = (typeof UNIT_PLACES)[keyof typeof UNIT_PLACES];

function styleOf(settings: VisualFormattingSettingsModel, theme: Theme): TableStyle {
    const size = (value: number | undefined, fallback: number) => Math.min(40, Math.max(6, value || fallback));
    const color = (slice: { value: { value: string } | undefined }, fallback: string) => slice.value?.value || fallback;
    const { comparison, periods, rows, segments, table } = settings;
    const headBackground = color(table.headBackground, DEFAULT_HEAD_BACKGROUND);
    return {
        fontFamily: table.fontFamily.value,
        fontSize: size(table.fontSize.value, DEFAULT_HEADING_SIZE),
        headingColor: table.headingColor.value?.value ?? "",
        sizes: {
            org: size(segments.orgSize.value, DEFAULT_TEXT_SIZES.org),
            name: size(rows.nameSize.value, DEFAULT_TEXT_SIZES.name),
            period: size(periods.periodSize.value, DEFAULT_TEXT_SIZES.period),
            column: size(comparison.columnSize.value, DEFAULT_TEXT_SIZES.column),
            compareHead: size(comparison.compareHeadSize.value, DEFAULT_TEXT_SIZES.compareHead),
            main: size(comparison.mainSize.value, DEFAULT_TEXT_SIZES.main),
            compare: size(comparison.compareSize.value, DEFAULT_TEXT_SIZES.compare),
            sub: size(comparison.subSize.value, DEFAULT_TEXT_SIZES.sub),
        },
        good: comparison.good.value?.value ?? DEFAULT_GOOD_COLOR,
        bad: comparison.bad.value?.value ?? DEFAULT_BAD_COLOR,
        bandColors: settings.rows.bandColors(),
        headBackground,
        // 見出しの背景を濃くしても字が読めるように、囲みの塗りと同じく明るさで字の色を選ぶ（明るい背景ならテーマの字の色）
        headText: textOn(headBackground, theme),
        headMuted: readableText(headBackground, theme.muted, 4.5),
        // 字・副文字・地（書式ペインに無いので、いつもテーマの色）
        text: theme.text,
        muted: theme.muted,
        background: theme.background,
        lines: {
            row: color(rows.rowLine, DEFAULT_LINE_COLORS.row),
            subtotal: color(rows.subtotalLine, DEFAULT_LINE_COLORS.subtotal),
            total: color(rows.totalLine, DEFAULT_LINE_COLORS.total),
            head: color(table.headLine, DEFAULT_LINE_COLORS.head),
            block: color(segments.blockLine, DEFAULT_LINE_COLORS.block),
            period: color(periods.periodLine, DEFAULT_LINE_COLORS.period),
            name: color(table.nameLine, DEFAULT_LINE_COLORS.name),
            band: color(rows.bandLine, DEFAULT_LINE_COLORS.band),
            outer: color(table.outerLine, DEFAULT_LINE_COLORS.outer),
        },
        periodLines: periods.periodLines.value,
        aggregateDouble: periods.aggregateDouble.value,
        unitPlace: unitPlaceOf(dropdownValue(settings.numbers.unitPlace, UNIT_PLACES.right)),
        titleInCorner: dropdownValue(periods.titlePlace, TITLE_PLACES.top) === TITLE_PLACES.corner,
        bandFill: settings.rows.bandFill.value,
        boldAggregates: settings.rows.boldAggregates.value,
        rowHeight: rowHeightOf(dropdownValue(settings.table.rowHeight, "normal")),
        alignTails: settings.numbers.alignTails.value,
        bold: Object.fromEntries(BOLD_KEYS.map((key) => {
            // 行の名前と数字の太さは「行」カードの「合計行を太字」1 つで決める（入れれば行の種類で太さを変える、切れば太字にしない）
            // 「行名をすべて太字」「数字をすべて太字」を入れれば、行の種類によらず太字
            if (key === "name" && settings.rows.allNamesBold.value) return [key, "on"];
            if ((key === "main" || key === "compare" || key === "sub") && settings.comparison.valueBold.value) return [key, "on"];
            if (key === "name" || key === "main" || key === "compare" || key === "sub") return [key, (settings.rows.totalBold.value ?? true) ? "auto" : "off"];
            const card = {"org":"segments","period":"periods","column":"comparison","compareHead":"comparison"}[key as "org" | "period" | "column" | "compareHead"] as "segments" | "periods" | "comparison";
            const value = dropdownValue((settings[card] as unknown as Record<string, Parameters<typeof dropdownValue>[0]>)[`${key}Bold`], "auto");
            return [key, value === "on" || value === "off" ? value : "auto"];
        })) as Record<BoldKey, BoldMode>,
        columnPadding: columnPaddingOf(dropdownValue(settings.table.columnPadding, "normal")),
        breakdown: { brackets: settings.rows.breakdownBrackets.value, muted: settings.rows.breakdownMuted.value, tag: settings.rows.breakdownTag.value },
        segmentFill: settings.segments.segmentFill.value,
        segmentColors: settings.segments.segmentColors(),
        headers: {
            segment: (settings.segments.segmentHeader.value ?? "").trim() || "セグメント",
            row: (settings.rows.rowHeader.value ?? "").trim() || "項目",
        },
        headerBold: { segment: settings.segments.segmentHeaderBold.value ?? false, row: settings.rows.rowHeaderBold.value ?? false },
    };
}

/** 薄い字の既定の色（visual.less の #605e5c と同じ） */
/** 塗った所の上の字の色：明るい塗りならテーマの字の色、暗い塗りなら白（contrastingText と同じ明るさの境） */
export function textOn(fill: string, theme: Pick<Theme, "text">): string {
    return contrastingText(fill) === "#FFFFFF" ? "#FFFFFF" : theme.text;
}

function rowHeightOf(value: string): RowHeight {
    return value in ROW_HEIGHTS ? (value as RowHeight) : "normal";
}

function columnPaddingOf(value: string): ColumnPadding {
    return value in COLUMN_PADDINGS ? (value as ColumnPadding) : "normal";
}

function unitPlaceOf(value: string): UnitPlace {
    return value === UNIT_PLACES.corner || value === UNIT_PLACES.name || value === UNIT_PLACES.cell ? value : UNIT_PLACES.right;
}

/** 比・率の上限（倍）。書式ペインは % で、100〜99999 に収める（0 と負の値も 100）。空なら既定 */
function ratioCapOf(settings: VisualFormattingSettingsModel): number {
    const percent = Number(settings.numbers.ratioCap.value);
    return Math.min(99999, Math.max(100, Number.isFinite(percent) ? Math.round(percent) : DEFAULT_RATIO_CAP)) / 100;
}

/** 「その他」の行の名前（まとめた科目の数。組織のブロックでは、その組織にデータのある科目だけを数える） */
function othersNamed(def: RowDef, has: (code: string) => boolean, count = true): RowDef {
    if (def.type !== "others") return def;
    // 件数は「行」カードの「その他の件数を出す」で切れる
    return { ...def, name: count ? `その他（${def.summands.filter((s) => has(s.code)).length}件）` : "その他" };
}

/** 「その他」にまとめた科目の名前（ツールチップ。多ければ先頭だけ） */
const OTHERS_LIST_MAX = 10;
function othersList(def: RowDef, has: (code: string) => boolean, nameOf: (code: string) => string): string {
    const names = def.summands.filter((s) => has(s.code)).map((s) => nameOf(s.code));
    return names.slice(0, OTHERS_LIST_MAX).join("・") + (names.length > OTHERS_LIST_MAX ? ` ほか${names.length - OTHERS_LIST_MAX}件` : "");
}

/** 木のすべての節（根を含む） */
function allNodes(node: OrgNode): OrgNode[] {
    return [node, ...node.children.flatMap(allNodes)];
}

/**
 * すべて開く・すべて閉じるの保存値（行は既定から切り替えた行なので、既定が閉じなら開くときに全部を入れる）。
 * 行を既定で開く表でも、その他・うちにした小計（closedByDefault）は既定で閉じ、「閉じた行」の記録を反転して読むので、開くときに入れ、閉じるときに外す
 */
function foldsOf(orgs: string[], rows: string[], rowsClosedByDefault: boolean, closedByDefault: Set<string>): ViewModel["folds"] {
    if (orgs.length === 0 && rows.length === 0) return null;
    return {
        orgs,
        rows,
        openAll: { openOrgs: orgs, ...(rowsClosedByDefault ? { openRows: rows } : { closedRows: rows.filter((code) => closedByDefault.has(code)) }) },
        closeAll: { openOrgs: [], ...(rowsClosedByDefault ? { openRows: [] } : { closedRows: rows.filter((code) => !closedByDefault.has(code)) }) },
    };
}

/** 書式ペインの「最新見込みの見出し」と文字の欄から、見出しの決まりを作る */
function latestHeaderRuleOf(settings: VisualFormattingSettingsModel, latest: string[]): LatestHeaderRule {
    const c = settings.comparison;
    const mode = dropdownValue(c.latestHeader, LATEST_HEADERS.label);
    if (mode === LATEST_HEADERS.scenario) return { newestFirst: latest };
    if (mode !== LATEST_HEADERS.custom) return {};
    return {
        confirmed: (c.confirmedScenario.value ?? "")
            .split(/[,、]/)
            .map((name) => name.trim())
            .filter((name) => name !== ""),
        confirmedLabel: (c.confirmedLabel.value ?? "").trim(),
        forecastLabel: (c.forecastLabel.value ?? "").trim(),
    };
}

/** 書式の「最新見込みの見出し」の決まり。何も無ければ、混ざったときは「最新見込み」 */
export interface LatestHeaderRule {
    /** シナリオの名前：比較順の大きい順。混ざったときは使ったうち比較順のいちばん前（確度の低い）名前 */
    newestFirst?: string[];
    /** 文字で決める：確定とみなすシナリオ。これだけなら confirmedLabel、ほかが入れば forecastLabel */
    confirmed?: string[];
    confirmedLabel?: string;
    forecastLabel?: string;
}

/**
 * 最新見込みの期間の名乗り：使ったイベントが 1 つならその名前、月か組織でイベントが混ざれば「最新見込み」。
 * 「シナリオの名前」なら、混ざったときは確度の低い名前（実績と見通しなら見通し、予定と実績なら予定）。
 * 「文字で決める」なら、確定とみなすシナリオだけの期間は確定の見出し（空ならそのシナリオの名前）、ほかが入れば別の見出し（空なら「最新見込み」）。
 * 月ごとの内訳は見出しの点線とツールチップ
 */
export function latestName(used: Map<MonthIndex, string[]> | Set<string>, rule: LatestHeaderRule = {}): string {
    const all = used instanceof Set ? used : new Set(Array.from(used.values()).flat());
    if (all.size === 0) return "";
    if (rule.confirmed) {
        const confirmed = rule.confirmed;
        if (Array.from(all).every((e) => confirmed.includes(e))) return rule.confirmedLabel || (confirmed.find((e) => all.has(e)) ?? "");
        return rule.forecastLabel || LATEST_LABEL;
    }
    if (all.size === 1) return Array.from(all)[0];
    return rule.newestFirst ? ([...rule.newestFirst].reverse().find((e) => all.has(e)) ?? LATEST_LABEL) : LATEST_LABEL;
}

/**
 * 月ごとの内訳（4〜5月 → 実績、6月 → 見通し）。expected を渡すと、どの月もそのイベントだけのときに出さない（比較の個々のイベント）。
 * 渡さなければ、1 か月に 1 つのイベントだけで、期間を通して同じなら出さない（最新見込み・前年同期）
 */
export function breakdownItems(used: Map<MonthIndex, string[]> | null, expected?: string): TooltipItem[] | undefined {
    if (!used || used.size === 0) return undefined;
    const entries = Array.from(used.entries()).sort((a, b) => a[0] - b[0]);
    const lists = entries.map(([, events]) => events.join("・"));
    if (expected !== undefined ? lists.every((l) => l === expected) : new Set(lists).size === 1 && entries[0][1].length === 1) return undefined;
    const parts: TooltipItem[] = [];
    let i = 0;
    while (i < entries.length) {
        let j = i;
        while (j + 1 < entries.length && lists[j + 1] === lists[i] && entries[j + 1][0] === entries[j][0] + 1) j++;
        const from = monthOf(entries[i][0]);
        const to = monthOf(entries[j][0]);
        parts.push({ displayName: from === to ? `${from}月` : `${from}〜${to}月`, value: lists[i] });
        i = j + 1;
    }
    return parts;
}

/**
 * 囲み：集計先（小計・見出し）と、その下の行（子・孫・内訳）をひとまとまりにして、左に段ごとの帯を引く。
 * 親の行は帯を横に伸ばして（下囲みなら L 字の底、上囲みなら上の辺）、どの行がまとまりを締めるかを見せる
 */
export function bandsOf(display: DisplayRow[]): Band[][] {
    const bands: Band[][] = display.map((): Band[] => []);
    const position = new Map(display.map((d, i) => [d.def.code, i]));
    const byCode = new Map(display.map((d) => [d.def.code, d.def]));
    // すぐ後ろに続く、自分より深い「うち」の行（親の下に置いた内訳）
    const breakdownsAfter = (index: number, depth: number) => {
        let to = index;
        while (to + 1 < display.length && display[to + 1].def.type === "breakdown" && display[to + 1].depth > depth) to++;
        return to;
    };
    display.forEach(({ def, depth }, index) => {
        // 小計・見出しの行（子のまとまり）と、うちを持つ行を箱にする
        const hasBreakdowns = breakdownsAfter(index, depth) > index;
        if (!hasBreakdowns && (def.children.length === 0 || (def.type !== "subtotal" && def.type !== "heading"))) return;
        // 子孫は、並びの中で自分の前か後ろに続けて置かれている（rows.ts の木の並べ方）。その端から端まで
        let from = index;
        let to = index;
        const walk = (code: string) => {
            for (const child of byCode.get(code)?.children ?? []) {
                const at = position.get(child);
                if (at !== undefined) {
                    from = Math.min(from, at);
                    to = Math.max(to, at);
                }
                walk(child);
            }
        };
        walk(def.code);
        // 最後の子のうちも、このまとまりの箱に入れる
        to = breakdownsAfter(to, depth);
        for (let i = from; i <= to; i++) bands[i].push({ level: depth, head: i === index, first: i === from, last: i === to, afterHead: i === index + 1 });
    });
    bands.forEach((list) => list.sort((a, b) => a.level - b.level));
    return bands;
}

/** 表の「自動」の単位の候補（小さい順）。日本語は 千・万・百万・億・兆、英語は K・M・bn・T */
const TABLE_AUTO_KEYS: Record<string, string[]> = {
    japanese: ["0", "3", "4", "6", "8", "12"],
    standard: ["0", "3", "6", "9", "12"],
};

/**
 * 表の「自動」の単位：一番大きい値が 4 けた以上残る、いちばん大きい単位（77 億なら百万で 7,730）。
 * グラフの自動（数字を小さくまとめる）と違い、表は月の小さな行も読めるよう、けたを残す
 */
export function tableAutoUnitKey(maxAbs: number, notation: string): string {
    const keys = TABLE_AUTO_KEYS[notation === "standard" ? "standard" : "japanese"];
    let chosen = keys[0];
    for (const key of keys) {
        if (maxAbs / Math.pow(10, Number(key)) >= 1000) chosen = key;
    }
    return chosen;
}

/**
 * 比較の列の見出しの 2 行。2 段の列はセルの上下と同じ順に「修正予算差」「修正予算比」と書く（前年同期差・前年同期比）。
 * 率は「修正予算率」（値に + と ▲ が付くので達成率の 101.2% とは見分けられる）。比較の値の列は相手の名前
 */
export function compareHeaderLines(compareLabel: string, view: CompareView, swap: boolean): [string, string] {
    const diff = `${compareLabel}差`;
    const ratio = `${compareLabel}比`;
    const rate = `${compareLabel}率`;
    switch (view) {
        case COMPARE_VIEWS.value:
            return [compareLabel, ""];
        case COMPARE_VIEWS.diff:
            return [diff, ""];
        case COMPARE_VIEWS.ratio:
            return [ratio, ""];
        case COMPARE_VIEWS.rate:
            return [rate, ""];
        default: {
            const second = view === COMPARE_VIEWS.diffRate ? rate : ratio;
            return swap ? [second, diff] : [diff, second];
        }
    }
}

function diffCell(
    main: number | null,
    compare: number | null,
    tone: Tone,
    percentRow: boolean,
    view: CompareView,
    swap: boolean,
    sign: SignOptions,
    cap: number,
    formatDiffValue: () => string
): Cell {
    const diffText = formatDiffValue();
    // 比率の行の比・率は意味が無い（差のポイントだけ出す。比・率だけの列でも差のポイント）
    if (percentRow || view === COMPARE_VIEWS.diff) return { text: diffText, tone };
    // 比・率だけの列で比・率が出せない（「―」）なら、差の向きの色は付けない（色の意味が「良い・悪い」なので）
    const only = (text: string): Cell => ({ text, tone: text === NOT_AVAILABLE ? null : tone });
    if (view === COMPARE_VIEWS.ratio) return only(formatRatio(main, compare, 1, cap));
    if (view === COMPARE_VIEWS.rate) return only(formatRate(main, compare, sign, 1, cap));
    const second = view === COMPARE_VIEWS.diffRate ? formatRate(main, compare, sign, 1, cap) : formatRatio(main, compare, 1, cap);
    if (!second) return { text: diffText, tone };
    return swap ? { text: second, sub: diffText, tone } : { text: diffText, sub: second, tone };
}

/** 比較の相手：値（イベント・最新見込み・前年同期）、値を取る側、名前 */
interface Partner {
    compare: string;
    ref: EventRef;
    label: string;
}

/** 期間の比較の列 1 本：シナリオ、相手、見せ方、列 1〜3 のどれか */
interface CompareColumnSpec {
    compare: string;
    partner: Partner;
    view: CompareView;
    slot: number;
}

/**
 * 期間の名前（見る人の期間ごとの相手の保存）。年月つきで、年度を替えると別の期間になる
 */
export function periodNameOf(period: { kind: PeriodKind; key: string; months: MonthIndex[] }): string {
    return periodIdOf(period);
}

