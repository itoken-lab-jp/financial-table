/**
 * 見せる科目を選ぶ。区分・中分類ごとに、見る人が選んだ科目だけを出し、選ばなかった科目は「その他」の 1 行にまとめるか、
 * 選んだ科目を親の「うち」にして残りを出さない。
 *
 * - どちらでも合計は変わらない。親の小計は、選んだかにかかわらず区分の科目すべてを足したまま（summands は変えない）
 * - 選べるのは、区分・中分類の小計の行の子で、その小計に足す科目だけ（区分と残高・フローの違う科目は選ばずにいつも出す）
 * - 保存するのは選んだ科目。あとから出てきた科目は選ばなかった側に入る（見たい科目だけの形を崩さない）
 * - 行の組み立て（rows.ts）のあとに、表示の並びだけを組み直す。書式ペインの選択肢（計算の行・指標の置く場所）は選択で変えない
 */
import { INDICATOR_KEY } from "./data";
import { DisplayRow, RowDef, RowModel } from "./rows";

/** others：選ばなかった科目を「その他」にまとめる。under：選んだ科目を親のうちにし、残りは出さない */
export type PickMode = "others" | "under";

export interface RowPick {
    /** 親の行（区分・中分類の小計の行）のコード */
    parent: string;
    mode: PickMode;
    /** 選んだ科目のコード */
    codes: string[];
}

/** 「その他」の行のコード（親の行のコードに付ける） */
export const OTHERS_KEY = "§others:";

/** 科目を選べる親（見る人のダイアログに並べる） */
export interface PickParent {
    code: string;
    name: string;
    depth: number;
    /** 選べる科目（表の並び） */
    accounts: Array<{ code: string; name: string }>;
}

/** 科目の行か（選んでうちにした科目も。指標のうちは入れない） */
export function isAccountRow(def: RowDef | undefined): boolean {
    return def !== undefined && (def.type === "detail" || (def.type === "breakdown" && !def.code.startsWith(INDICATOR_KEY)));
}

/** 親の子のうち、選べる科目（親の小計に足す科目） */
function pickableOf(parent: RowDef, rows: ReadonlyMap<string, RowDef>): string[] {
    const summed = new Set(parent.summands.map((s) => s.code));
    return parent.children.filter((code) => rows.get(code)?.type === "detail" && summed.has(code));
}

/** 選べる科目が 2 つ以上の区分・中分類か（1 つなら選ぶことが無い）。ダイアログに出す親と、選択を効かせる親をそろえる（出ない親に効くと戻せない） */
const PICK_MIN = 2;

/**
 * 科目を選べる親。並びは区分 → その中分類（小計を下に置く表でも、区分を先に）。中分類の名前には区分の名前を添える
 * （区分がダイアログに出ないとき、別の区分の同じ名前の中分類を見分ける）。組み直す前の行で決める
 */
export function pickParents(model: RowModel): PickParent[] {
    const position = new Map(model.display.map((d, i): [string, number] => [d.def.code, i]));
    /** 行の中で一番上の位置（子・孫を含む） */
    const top = (code: string): number => Math.min(position.get(code) ?? Number.POSITIVE_INFINITY, ...(model.rows.get(code)?.children ?? []).map(top));
    const parentOf = (code: string) => model.display.find((d) => d.def.type === "subtotal" && d.def.children.includes(code))?.def;
    return model.display
        .flatMap(({ def, depth }): Array<PickParent & { at: number }> => {
            if (def.type !== "subtotal") return [];
            const accounts = pickableOf(def, model.rows).map((code) => ({ code, name: model.rows.get(code)!.name }));
            if (accounts.length < PICK_MIN) return [];
            const section = depth > 0 ? parentOf(def.code) : undefined;
            return [{ code: def.code, name: section ? `${def.name}（${section.name}）` : def.name, depth, accounts, at: top(def.code) }];
        })
        .sort((a, b) => a.at - b.at || a.depth - b.depth)
        .map((p): PickParent => ({ code: p.code, name: p.name, depth: p.depth, accounts: p.accounts }));
}

/** 見る人の選択で、表示の並びを組み直す。選択が無ければ同じものを返す */
export function applyPicks(model: RowModel, picks: readonly RowPick[]): RowModel {
    if (picks.length === 0) return model;
    const rows = new Map(model.rows);
    const attached = new Map(model.attached);
    const following = new Map(model.following);
    let display: DisplayRow[] = model.display.slice();

    /** 行と、その下の行（子・孫・うち・後ろの指標） */
    const subtree = (code: string, into = new Set<string>()): Set<string> => {
        if (into.has(code)) return into;
        into.add(code);
        for (const next of [...(rows.get(code)?.children ?? []), ...(attached.get(code) ?? []), ...(following.get(code) ?? [])]) subtree(next, into);
        return into;
    };

    const done = new Set<string>();
    for (const pick of picks) {
        const parent = rows.get(pick.parent);
        // 同じ親が 2 度あれば最初のものだけ（読むときにもまとめる）
        if (!parent || parent.type !== "subtotal" || done.has(parent.code)) continue;
        done.add(parent.code);
        const pickable = pickableOf(parent, rows);
        if (pickable.length < PICK_MIN) continue;
        const chosen = new Set(pick.codes);
        const selected = pickable.filter((code) => chosen.has(code));
        const hidden = pickable.filter((code) => !chosen.has(code));
        const parentAt = display.findIndex((d) => d.def.code === parent.code);
        if (parentAt < 0) continue;
        const depth = display[parentAt].depth + 1;
        const removed = new Set<string>();
        for (const code of hidden) subtree(code, removed);

        if (pick.mode === "others") {
            if (hidden.length === 0) continue;
            const code = OTHERS_KEY + parent.code;
            const others: RowDef = {
                code,
                name: "その他",
                type: "others",
                sign: parent.sign,
                good: parent.good,
                aggregation: parent.aggregation,
                format: parent.format,
                formula: null,
                refs: [],
                children: [],
                summands: hidden.map((c) => ({ code: c, weight: 1 })),
            };
            rows.set(code, others);
            // 区分の中の最後（子と、その下の行の後ろ）
            const inBlock = new Set<string>();
            for (const child of parent.children) subtree(child, inBlock);
            let end = -1;
            display.forEach((d, i) => {
                if (inBlock.has(d.def.code)) end = i;
            });
            display.splice(end + 1, 0, { def: others, depth });
            rows.set(parent.code, { ...parent, children: [...parent.children.filter((c) => !removed.has(c)), code] });
        } else {
            // うち：選んだ科目（とその下の行）を親の行の下へ移し、うちの行にする。親の子からは科目を外す（囲みの帯を引かない）。
            // 親に残った子（中分類など）があれば、その後ろ（小計を上に置く表で、うちが親の帯の中に入らないように。指標のうちと同じ）
            const moved: DisplayRow[] = [];
            for (const code of selected) {
                const own = subtree(code);
                for (const d of display) {
                    if (!own.has(d.def.code)) continue;
                    if (d.def.code === code) {
                        const def: RowDef = { ...d.def, type: "breakdown" };
                        rows.set(code, def);
                        moved.push({ def, depth: d.depth });
                    } else moved.push(d);
                    removed.add(d.def.code);
                }
            }
            const kept = display.filter((d) => !removed.has(d.def.code));
            const rest = new Set<string>();
            for (const child of parent.children) if (!pickable.includes(child)) subtree(child, rest);
            let at = kept.findIndex((d) => d.def.code === parent.code);
            kept.forEach((d, i) => {
                if (rest.has(d.def.code) && i > at) at = i;
            });
            kept.splice(at + 1, 0, ...moved);
            display = kept;
            attached.set(parent.code, [...selected, ...(attached.get(parent.code) ?? [])]);
            rows.set(parent.code, { ...parent, children: parent.children.filter((c) => !pickable.includes(c)) });
            continue;
        }
        display = display.filter((d) => !removed.has(d.def.code));
    }
    // 並びの行は、組み直した定義を指す
    display = display.map((d) => (rows.get(d.def.code) === d.def ? d : { ...d, def: rows.get(d.def.code) ?? d.def }));
    return { ...model, rows, display, attached, following };
}
