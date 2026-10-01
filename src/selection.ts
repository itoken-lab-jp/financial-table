/**
 * 行・列・セル・組織を押したときの選択。Profitbase の見比べで足りないと分かった挙動。
 *
 * - 行を押すと、その科目（区分・中分類の行はその下の科目すべて）でほかのビジュアルを絞る。組織のブロックの中なら組織 × 科目
 * - セルを押すと、行 × 列の月（四半期・通期ならその月すべて）で絞る。列の見出しを押すと、その月で絞る
 * - 左の列の組織の名前を押すと、その組織で絞る
 * - Ctrl（Mac は ⌘）で足し引き。同じものだけを選んでいるときにもう一度押すと解く。表の空いた所を押しても解く
 * - 形の違うもの（行と組織の名前、組織の段の深さが違うものなど）は Ctrl でも足さず、押したものに置き換える。選択 ID の列の組み合わせが
 *   混ざると、ほかのビジュアルが「サポートされていないクロス強調表示」になった（2026-09-27 Desktop）
 * - 選んだもののほかを薄くする（計算の行・指標の行は選べない。薄くもしない＝Profitbase の合計の行と同じ）
 *
 * 選択 ID（Power BI）は visual.ts が SelectionNodes（data.ts）から作る。ここは何を選んだかと、どこを明るく残すかだけ
 */
import { ORG_SEPARATOR } from "./data";
import { MonthIndex } from "./periods";
import type { ViewModel, TableRow } from "./viewModel";
import { planHeaders } from "./hierarchy";

/** 表の行だけでなく、分割帯と横積みの結合セルも選択可能な表示として扱う。 */
export function selectableView(model: ViewModel): SelectableView {
    const view: SelectableView = { rows: new Map(), periods: new Map(model.periods.map(p => [p.key, p.months])), orgs: new Set() };
    const addRow = (row: TableRow): void => {
        if (row.select) view.rows.set(row.key, row.select);
        for (const cell of row.orgCells ?? []) {
            if (cell.path !== null) view.orgs.add(cell.path);
            for (const head of cell.heads ?? []) if (head.path !== null) view.orgs.add(head.path);
        }
        for (const cell of row.hierarchyCells ?? []) addRow(cell.row);
    };
    model.rows.forEach(addRow);
    if (model.layout) for (const header of planHeaders(model.layout)) {
        if (header.org != null) view.orgs.add(header.org);
        if (header.row) addRow(header.row);
    }
    return view;
}

export type SelectTarget =
    | { kind: "row"; rowKey: string; codes: string[]; org: string | null }
    | { kind: "cell"; rowKey: string; periodKey: string; codes: string[]; org: string | null; months: MonthIndex[] }
    | { kind: "period"; periodKey: string; months: MonthIndex[] }
    | { kind: "org"; path: string };

/** 同じものかを見分けるキー */
export function targetKey(target: SelectTarget): string {
    switch (target.kind) {
        case "row":
            return `row\u0001${target.rowKey}`;
        case "cell":
            return `cell\u0001${target.rowKey}\u0001${target.periodKey}`;
        case "period":
            return `period\u0001${target.periodKey}`;
        case "org":
            return `org\u0001${target.path}`;
    }
}

/** 選択 ID の形（どの段の節点で作るか）。種類と、組織の段の深さ（組織の無い行は 0） */
export function targetShape(target: SelectTarget): string {
    const depth = (org: string | null) => (org === null ? 0 : org.split(ORG_SEPARATOR).length);
    switch (target.kind) {
        case "row":
        case "cell":
            return `${target.kind}:${depth(target.org)}`;
        case "period":
            return "period";
        case "org":
            return `org:${depth(target.path)}`;
    }
}

/** 押したあとの選択。multi は Ctrl（⌘）を押しながら。形の違うものは足さずに置き換える */
export function nextSelection(current: SelectTarget[], target: SelectTarget, multi: boolean): SelectTarget[] {
    const key = targetKey(target);
    const has = current.some((t) => targetKey(t) === key);
    if (multi && current.length > 0 && targetShape(current[0]) !== targetShape(target)) return [target];
    if (multi) return has ? current.filter((t) => targetKey(t) !== key) : [...current, target];
    return has && current.length === 1 ? [] : [target];
}

const inOrg = (org: string | null, path: string) => org !== null && (org === path || org.startsWith(path + ORG_SEPARATOR));

/** セルを明るく残すか（選択が無ければ全部明るい）。org はそのセルの行の組織のブロックの道筋 */
export function cellActive(selection: SelectTarget[], rowKey: string, periodKey: string, org: string | null): boolean {
    if (selection.length === 0) return true;
    return selection.some((t) =>
        t.kind === "row"
            ? t.rowKey === rowKey
            : t.kind === "cell"
              ? t.rowKey === rowKey && t.periodKey === periodKey
              : t.kind === "period"
                ? t.periodKey === periodKey
                : inOrg(org, t.path)
    );
}

/** 行の見出しを明るく残すか（列だけを選んでいれば、どの行も明るい） */
export function rowActive(selection: SelectTarget[], rowKey: string, org: string | null): boolean {
    if (selection.length === 0) return true;
    return selection.some((t) => (t.kind === "row" || t.kind === "cell" ? t.rowKey === rowKey : t.kind === "period" ? true : inOrg(org, t.path)));
}

/** 列の見出しを明るく残すか（行・組織だけを選んでいれば、どの列も明るい） */
export function periodActive(selection: SelectTarget[], periodKey: string): boolean {
    if (selection.length === 0) return true;
    return selection.some((t) => (t.kind === "period" || t.kind === "cell" ? t.periodKey === periodKey : true));
}

/** 選べない行（計算の行・指標）の見出しは薄くしない。ただし組織を選んでいれば、ほかの組織のブロックの行は薄くする */
export function plainRowActive(selection: SelectTarget[], org: string | null): boolean {
    const orgs = selection.filter((t): t is Extract<SelectTarget, { kind: "org" }> => t.kind === "org");
    return orgs.length === 0 || orgs.some((t) => inOrg(org, t.path));
}

/** 選べない行のセル：組織のほかに、列を選んでいれば選んだ列だけを明るく残す（行によって列の明るさが縞にならないように） */
export function plainCellActive(selection: SelectTarget[], periodKey: string, org: string | null): boolean {
    if (!plainRowActive(selection, org)) return false;
    const periods = selection.filter((t) => t.kind === "period");
    return periods.length === 0 || periods.some((t) => t.kind === "period" && t.periodKey === periodKey);
}

/** 今の表の、選べるもの（行のキーと絞るもの・期間のキーと月・組織の道筋） */
export interface SelectableView {
    rows: Map<string, { codes: string[]; org: string | null }>;
    periods: Map<string, MonthIndex[]>;
    orgs: Set<string>;
}

/**
 * 表が変わったとき（行・組織を閉じた、年度を変えた、データが変わった）に選択を合わせ直す。表に無くなったもの・期間の月が変わったもの
 * （年度が替わって 1Q が別の月になった、累計が伸びた）は外し、行が絞る科目は今の表のものにする
 */
export function reconcileSelection(selection: SelectTarget[], view: SelectableView): SelectTarget[] {
    const sameMonths = (a: MonthIndex[], b: MonthIndex[] | undefined) => b !== undefined && a.length === b.length && a.every((m, i) => m === b[i]);
    return selection.flatMap((t): SelectTarget[] => {
        switch (t.kind) {
            case "row": {
                const row = view.rows.get(t.rowKey);
                return row ? [{ ...t, codes: row.codes, org: row.org }] : [];
            }
            case "cell": {
                const row = view.rows.get(t.rowKey);
                return row && sameMonths(t.months, view.periods.get(t.periodKey)) ? [{ ...t, codes: row.codes, org: row.org }] : [];
            }
            case "period":
                return sameMonths(t.months, view.periods.get(t.periodKey)) ? [t] : [];
            case "org":
                return view.orgs.has(t.path) ? [t] : [];
        }
    });
}

/** 合わせ直しで変わったか */
export function sameSelection(a: SelectTarget[], b: SelectTarget[]): boolean {
    return a.length === b.length && a.every((t, i) => JSON.stringify(t) === JSON.stringify(b[i]));
}

/** 選ばれた組織の名前か（左の列の見出しを枠で見せる） */
export function orgSelected(selection: SelectTarget[], path: string | null): boolean {
    return path !== null && selection.some((t) => t.kind === "org" && t.path === path);
}
