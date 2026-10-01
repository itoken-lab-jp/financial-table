/**
 * 区分ごとの数値（「数値」カードの設定を区分で上書きする）。書式ペインの枠に区分の名前と設定を書く。
 *
 * - 効くのはその区分の小計・中分類の小計・科目・科目のうち（と、見る人がまとめた「その他」）。計算の行と、後ろに置いた指標の行は表全体のまま
 *   （区分をまたぐ行なので）
 * - 区分で選べる表示単位は固定の単位だけ。表全体の単位と違えば、区分の行の名前に単位を添える
 */
import { AccountRecord } from "./data";
import { NO_CATEGORY, RowModel } from "./rows";
import { SECTION_TABLE, SectionNumberProp, SectionNumberSaved, sectionNumberAllowed } from "./settings";

/** 枠の保存値から、区分 → 上書き（表全体に合わせる設定は入れない）。表に無い区分・2 つの枠の同じ区分は知らせる（同じ区分は番号の小さい枠） */
export function sectionOverrides(
    saved: SectionNumberSaved[],
    sections: string[],
    warn: (message: string) => void,
    allowed: (prop: SectionNumberProp) => string[] = sectionNumberAllowed,
): Map<string, Partial<SectionNumberSaved>> {
    const result = new Map<string, Partial<SectionNumberSaved>>();
    saved.forEach((slot, i) => {
        if (slot.section === "") return;
        const override: Partial<SectionNumberSaved> = {};
        for (const [key, value] of Object.entries(slot) as Array<[keyof SectionNumberSaved, string]>) {
            // 書式ペインの選択肢に無い値は使わない（ペインは「表全体に合わせる」と見せるので、見えない上書きになる。区分の「自動」の単位など）
            if (key !== "section" && value !== "" && value !== SECTION_TABLE && allowed(key).includes(value)) override[key] = value;
        }
        // 何も上書きしない枠は何もしない（区分を取らず、知らせない）
        if (Object.keys(override).length === 0) return;
        if (!sections.includes(slot.section)) {
            warn(`区分ごとの数値の設定 ${i + 1} の区分「${slot.section}」が表に無い。効かせていない`);
            return;
        }
        if (result.has(slot.section)) {
            warn(`区分ごとの数値の設定 ${i + 1} の「${slot.section}」は、番号の小さい設定でも決めている。番号の小さい設定を使った`);
            return;
        }
        result.set(slot.section, override);
    });
    return result;
}

/**
 * 行のコード → 区分。科目は科目の区分。小計・見出しは、その下の科目（子・うち）がみな同じ区分ならその区分。見る人がまとめた「その他」は
 * まとめた科目の区分。科目のうち（指標のうち）は親の行の区分。計算の行と、後ろに置いた指標は区分を持たない（子の親にならないので）
 */
export function rowSections(model: RowModel, accounts: ReadonlyMap<string, AccountRecord>): Map<string, string> {
    const kidsOf = (code: string) => [...(model.rows.get(code)?.children ?? []), ...(model.attached.get(code) ?? [])];
    const parentOf = new Map<string, string>();
    for (const code of model.rows.keys()) for (const kid of kidsOf(code)) parentOf.set(kid, code);
    const down = new Map<string, string | null>();
    /** 下の行（と自分）から決まる区分 */
    const fromBelow = (code: string): string | null => {
        const known = down.get(code);
        if (known !== undefined) return known;
        down.set(code, null);
        let result: string | null = null;
        const account = accounts.get(code);
        const def = model.rows.get(code);
        if (account && (def?.type === "detail" || def?.type === "breakdown")) result = account.category ?? NO_CATEGORY;
        else {
            // 「その他」はまとめた科目。区分・中分類の小計は、子とうちに加えて足す行も（見る人が「うち」で全部外すと子もうちも無くなる）。
            // 計算の行（step）は足す行が区分をまたぐので見ない
            const below =
                def?.type === "others"
                    ? def.summands.map((s) => s.code)
                    : def?.type === "subtotal"
                      ? [...kidsOf(code), ...def.summands.map((s) => s.code)]
                      : kidsOf(code);
            const labels = new Set(below.map(fromBelow).filter((label): label is string => label !== null));
            if (labels.size === 1) result = Array.from(labels)[0];
        }
        down.set(code, result);
        return result;
    };
    const result = new Map<string, string>();
    for (const { def } of model.display) {
        let label = fromBelow(def.code);
        // 下から決まらない行（指標のうち・子の無い行）は、親の行の区分
        for (let parent = parentOf.get(def.code); label === null && parent !== undefined; parent = parentOf.get(parent)) label = fromBelow(parent);
        if (label !== null) result.set(def.code, label);
    }
    return result;
}
