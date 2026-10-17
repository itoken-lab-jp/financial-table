/**
 * 行別の数値書式（「数値」カードの設定を、段か行で上書きする。2.0 で「区分ごとの数値」を広げた）。書式ペインの枠に対象と設定を書く。
 *
 * - 対象は段（その深さの行すべて）か行（区分・中分類・科目。その行と下の行）。上書きの順は 表全体 → 段 → 行 で、行は上の行から継いで
 *   自分に近い行ほど後に効く。設定ごとに継ぐ（単位だけ区分で変え、マイナスの書き方は科目で変える、ができる）
 * - 段と、上の行から継ぐ上書きが効くのは科目の木の行（区分・中分類の小計、科目、科目のうち、見る人がまとめた「その他」）。計算行と、
 *   後ろに置いた指標の行は、その行を対象に選んだときだけ効く（区分をまたぐ行なので、段・上の行からは継がない。2.1.7.15 から。それまでは
 *   表全体のままで、書式・良し悪し・太字は計算行・指標のカードにあった）
 * - 選べる表示単位は固定の単位だけ。表全体の単位と違えば、親と単位の違う行の名前に単位を添える
 * - 保存のキーは行のコード（区分は「§sec:区分」、中分類は「§grp:区分:中分類」、科目はコード）と段（「§level:深さ」）。枠の番号や並びでは覚えない
 */
import { INDICATOR_KEY } from "./data";
import { RowDef, RowModel, TOTAL_KEY } from "./rows";
import { ROW_NUMBER_TABLE, RowNumberProp, RowNumberSaved, rowNumberAllowed } from "./settings";

/** 段の対象の値の頭（後ろは表示の深さ。一番上が 0） */
export const LEVEL_KEY = "§level:";

/** 上書きの中身（表全体に合わせる設定は入れない） */
export type RowNumberOverride = Partial<Omit<RowNumberSaved, "target">>;

/** 書式ペインの対象の選択肢 */
export interface RowNumberTarget {
    value: string;
    displayName: string;
    /** 計算行・後ろに置いた指標の行（太字が効くのはこの行だけ） */
    own?: boolean;
}

/** 科目の木の行か（段と行の上書きが効く行）。指標の行のうちは親の行と同じ見せ方なので入れ、後ろに置いた指標（measure）は入れない */
function inTree(def: RowDef): boolean {
    return def.code === TOTAL_KEY || def.type === "detail" || def.type === "subtotal" || def.type === "heading" || def.type === "others" || def.type === "breakdown";
}

/** 計算行・後ろに置いた指標の行か（その行を選んだときだけ上書きが効く行） */
export function ownRow(def: RowDef): boolean {
    return def.type === "calc" || def.type === "step" || def.type === "measure";
}

/** 行 → 親の行（子とうちから引く） */
export function parentsOf(model: RowModel): Map<string, string> {
    const parentOf = new Map<string, string>();
    for (const [code, def] of model.rows) for (const kid of [...def.children, ...(model.attached.get(code) ?? [])]) if (!parentOf.has(kid)) parentOf.set(kid, code);
    return parentOf;
}

/** 書式ペインの対象の選択肢：段（深さの順）と行（表の並び。名前は上の行からの道筋） */
export function rowNumberTargets(model: RowModel): RowNumberTarget[] {
    const parentOf = parentsOf(model);
    const rows = model.display.filter((d) => inTree(d.def) && d.def.type !== "others" && d.def.type !== "heading" && !d.def.code.startsWith(INDICATOR_KEY));
    const depths = Array.from(new Set(rows.map((d) => d.depth))).sort((a, b) => a - b);
    const pathOf = (code: string): string[] => {
        const def = model.rows.get(code);
        const parent = parentOf.get(code);
        return [...(parent !== undefined ? pathOf(parent) : []), def?.name ?? code];
    };
    const named = rows.filter((d) => d.def.type !== "breakdown").map((d) => ({ value: d.def.code, displayName: pathOf(d.def.code).join(" / ") }));
    // 同じ名前の行（コードの違う同じ名前の行）は、コードを添えて見分ける
    const count = new Map<string, number>();
    for (const t of named) count.set(t.displayName, (count.get(t.displayName) ?? 0) + 1);
    // 計算行と後ろに置いた指標の行（表の並び）。名前の頭に種類を付ける
    const own = model.display
        // 合計行は科目の木の行として上で並べたので、ここでは外す（「合計」と「計算行：合計」が 2 回並んだ。通し確認の 12）
        .filter((d) => ownRow(d.def) && d.def.code !== TOTAL_KEY)
        .map((d) => ({ value: d.def.code, displayName: `${d.def.type === "measure" ? "指標" : "計算行"}：${d.def.name || d.def.code}`, own: true }));
    return [
        ...depths.map((depth) => ({ value: LEVEL_KEY + depth, displayName: `${depth + 1} 段目の行すべて` })),
        ...named.map((t) => (count.get(t.displayName)! > 1 ? { ...t, displayName: `${t.displayName}（${t.value}）` } : t)),
        ...own,
    ];
}

/** 枠の保存値から、対象 → 上書き。表に無い対象・2 つの枠の同じ対象は知らせる（同じ対象は番号の小さい枠） */
export function rowNumberOverrides(
    saved: RowNumberSaved[],
    targets: string[],
    warn: (message: string) => void,
    allowed: (prop: RowNumberProp) => string[] | null = rowNumberAllowed,
): Map<string, RowNumberOverride> {
    const result = new Map<string, RowNumberOverride>();
    saved.forEach((slot, i) => {
        if (slot.target === "") return;
        const override: RowNumberOverride = {};
        for (const [key, value] of Object.entries(slot) as Array<[keyof RowNumberSaved, string]>) {
            if (key === "target" || value === "" || value === ROW_NUMBER_TABLE) continue;
            // 書式ペインの選択肢に無い値は使わない（ペインは「表全体と同じ」と見せるので、見えない上書きになる。「自動」の単位など）。
            // 書式・通貨の字は字を書く欄なので、どの字でも使う
            const choices = allowed(key);
            if (choices === null || choices.includes(value)) override[key] = value;
        }
        // 何も上書きしない枠は何もしない（対象を取らず、知らせない）
        if (Object.keys(override).length === 0) return;
        if (!targets.includes(slot.target)) {
            warn(`「数値」カードの適用先「${slot.target}」が表に無い。効かせていない`);
            return;
        }
        if (result.has(slot.target)) {
            warn(`「数値」カードの適用先に同じ行が 2 つある。先に足した方を使った`);
            return;
        }
        result.set(slot.target, override);
    });
    return result;
}

/**
 * 行のコード → 効く上書き（表全体 → 段 → 上の行から自分の行の順に重ねたもの）。上書きの無い行は入れない。
 * 計算行・後ろの指標の行は、その行の上書きだけ
 */
export function resolveRowNumbers(model: RowModel, overrides: ReadonlyMap<string, RowNumberOverride>): Map<string, RowNumberOverride> {
    const result = new Map<string, RowNumberOverride>();
    if (overrides.size === 0) return result;
    const parentOf = parentsOf(model);
    const chainOf = (code: string): string[] => {
        const chain: string[] = [];
        for (let at: string | undefined = code; at !== undefined && !chain.includes(at); at = parentOf.get(at)) chain.unshift(at);
        return chain;
    };
    for (const { def, depth } of model.display) {
        if (ownRow(def)) {
            const own = overrides.get(def.code);
            if (own && Object.keys(own).length > 0) result.set(def.code, { ...own });
            continue;
        }
        if (!inTree(def)) continue;
        const merged: RowNumberOverride = { ...(overrides.get(LEVEL_KEY + depth) ?? {}) };
        for (const code of chainOf(def.code)) {
            // 見る人が小計をうちにした行（picks.ts の UNDER_KEY。足す行が元の小計 1 つ）は、元の小計の行の上書き
            const row = model.rows.get(code);
            const source = row?.type === "breakdown" && row.summands.length === 1 ? row.summands[0].code : code;
            Object.assign(merged, overrides.get(source) ?? {});
        }
        if (Object.keys(merged).length > 0) result.set(def.code, merged);
    }
    return result;
}

/** 上書きの中身を、同じものどうしで同じになるキーにする（単位の最大や、行の数値の設定をまとめて持つ） */
export function overrideKey(override: RowNumberOverride): string {
    return JSON.stringify(Object.entries(override).sort(([a], [b]) => a.localeCompare(b)));
}
