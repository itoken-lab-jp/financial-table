/**
 * 金額の符号の持ち方。
 *
 * メジャーが返す値の持ち方は 3 通り。表はイベントごとに借方プラス（借方の科目が正、貸方の科目が負）にそろえてから、
 * 区分の向きで見せる（rows.ts の sign）。
 * - 借方プラス：会計システムの試算表。そのまま
 * - 貸方プラス：足すと利益になる管理会計の出力（営業利益のメジャー＝売上 − 費用 を科目で割ったもの）。すべて反転
 * - すべてプラス：Excel の予算。貸方の科目だけ反転
 * 「自動」はイベントごとに、貸方の科目の合計が負ならそのまま（正なら反転）、借方の科目の合計が正ならそのまま（負なら反転）で決める。
 * 全月の合計で見るので、返品で売上が負の月があっても揺れにくい。実績は借方プラス、予算はすべてプラス、のような混ざり方も読む。
 * 金額の科目（rows.ts の traits.amount）だけに効く。人数・時間のような金額でない科目は届いた値のまま。
 * 見分けには区分の合計に足す科目だけを使う（損益の区分の残高フラグの科目・内訳は数えない）
 */
import { Fact } from "./data";
import { yearOf } from "./periods";
import { AccountTraits } from "./rows";

export type AmountSign = "auto" | "debit" | "credit" | "positive";

/** イベントの、貸方の科目・借方の科目に掛ける係数（借方プラスにそろえる） */
export interface EventSign {
    credit: 1 | -1;
    debit: 1 | -1;
}

const FIXED: Record<Exclude<AmountSign, "auto">, EventSign> = {
    debit: { credit: 1, debit: 1 },
    credit: { credit: -1, debit: -1 },
    positive: { credit: -1, debit: 1 },
};

/** 合計がこの割合より 0 に近ければ、向きを決めない（貸方の科目どうしが打ち消し合っているなど） */
const NEAR_ZERO = 1e-6;
/**
 * 組織・まとまり・年ごとの合計が、絶対値の和のこの割合以上でイベントと逆なら（金額の 3/4 以上が逆の符号）、持ち方が混ざっていると見る。
 * 本社費の配賦（費用のマイナス）や、本社の営業外の損益のように、打ち消し合う科目の多い所を混ざりと取り違えないため
 */
const CLEAR = 0.5;

interface Sums {
    credit: number;
    creditAbs: number;
    debit: number;
    debitAbs: number;
}

const emptySums = (): Sums => ({ credit: 0, creditAbs: 0, debit: 0, debitAbs: 0 });

/** 合計の符号で係数を決める。normal はその向きの科目の借方プラスでの符号（貸方は −1、借方は +1）。その向きの科目が無ければ 1 */
function decide(total: number, abs: number, normal: -1 | 1): 1 | -1 | null {
    if (abs === 0) return 1;
    if (Math.abs(total) <= abs * NEAR_ZERO) return null;
    return Math.sign(total) === normal ? 1 : -1;
}

/**
 * イベントごとの符号の持ち方。setting はイベントの設定（横持ちのメジャーごと。無ければ表全体の「金額の持ち方」）。
 * 自動で決めきれないときは借方プラスとして読み、知らせる。同じイベントの中で組織・損益と貸借・年によって持ち方が違えば、
 * イベントの合計の多いほうに合わせて読み、知らせる
 */
export function resolveSigns(
    facts: Fact[],
    traits: Map<string, AccountTraits>,
    events: string[],
    setting: (event: string) => AmountSign
): { signs: Map<string, EventSign>; warnings: string[] } {
    const warnings: string[] = [];
    const sums = new Map<string, Sums>();
    /** イベント → 組織 × まとまり（フロー・残高）× 年 → 合計（持ち方が混ざっていないかを見る。年は途中で持ち方を変えたデータ） */
    const parts = new Map<string, Map<string, Sums>>();
    /** 金額の科目の値が 1 つでもあるイベント */
    const hasAmount = new Set<string>();
    for (const fact of facts) {
        const trait = traits.get(fact.code);
        if (trait?.amount) hasAmount.add(fact.event);
        if (!trait?.inTotal || trait.credit === null) continue;
        const add = (sum: Sums) => {
            if (trait.credit) {
                sum.credit += fact.value;
                sum.creditAbs += Math.abs(fact.value);
            } else {
                sum.debit += fact.value;
                sum.debitAbs += Math.abs(fact.value);
            }
        };
        let sum = sums.get(fact.event);
        if (!sum) sums.set(fact.event, (sum = emptySums()));
        add(sum);
        let byPart = parts.get(fact.event);
        if (!byPart) parts.set(fact.event, (byPart = new Map()));
        const key = `${fact.org}\u0000${trait.stock ? "stock" : "flow"}\u0000${yearOf(fact.month)}`;
        let part = byPart.get(key);
        if (!part) byPart.set(key, (part = emptySums()));
        add(part);
    }

    const signs = new Map<string, EventSign>();
    const undecided: string[] = [];
    const inverted: string[] = [];
    const mixed: string[] = [];
    for (const event of events) {
        const chosen = setting(event);
        if (chosen !== "auto") {
            signs.set(event, FIXED[chosen]);
            continue;
        }
        const sum = sums.get(event);
        // 貸方の科目の合計が負なら借方プラスの形（そのまま）、正なら反転。借方の科目は正ならそのまま、負なら反転
        const credit = sum ? decide(sum.credit, sum.creditAbs, -1) : null;
        const debit = sum ? decide(sum.debit, sum.debitAbs, 1) : null;
        const resolved: EventSign = { credit: credit ?? 1, debit: debit ?? 1 };
        signs.set(event, resolved);
        // 見分けに使える科目が無い（内訳だけなど）か、合計がほぼ 0 なら決めきれない
        if (credit === null || debit === null) {
            if (hasAmount.has(event)) undecided.push(event);
            continue;
        }
        // 貸方の科目はそのまま・借方の科目は反転（どちらもマイナス）は 3 通りのどれでもない
        if (resolved.credit === 1 && resolved.debit === -1 && sum!.creditAbs > 0 && sum!.debitAbs > 0) inverted.push(event);
        // 組織・損益と貸借・年ごとに見て、イベント全体とはっきり逆の向きの所があれば、持ち方が混ざっている
        const opposite = (total: number, abs: number, normal: -1 | 1, factor: 1 | -1) =>
            abs > 0 && Math.abs(total) >= abs * CLEAR && decide(total, abs, normal) !== factor;
        for (const part of parts.get(event)?.values() ?? []) {
            if (opposite(part.credit, part.creditAbs, -1, resolved.credit) || opposite(part.debit, part.debitAbs, 1, resolved.debit)) {
                mixed.push(event);
                break;
            }
        }
    }
    if (undecided.length > 0) {
        warnings.push(
            `符号の持ち方を決めきれないシナリオがある（${undecided.slice(0, 3).join("・")}。貸方か借方の行の合計がほぼ 0、か見分けに使える行が無い）。借方プラスとして読んだ。持ち方を書式ペインの「数値」の符号の持ち方で選ぶ`
        );
    }
    if (inverted.length > 0) {
        warnings.push(`値の符号が 3 通りのどれでもないシナリオがある（${inverted.slice(0, 3).join("・")}。貸方の行も借方の行もマイナス）。借方の行を反転して読んだ。データの符号を確かめる`);
    }
    if (mixed.length > 0) {
        warnings.push(
            `符号の持ち方が同じシナリオの中で混ざっている（${mixed.slice(0, 3).join("・")}。セグメントか、損益と貸借か、年で符号の持ち方が違う）。シナリオの合計の多いほうに合わせて読んだ。メジャーかデータで持ち方をそろえる`
        );
    }
    return { signs, warnings };
}

/** 金額の科目の値を、イベントごとの係数で借方プラスにそろえる。金額でない科目・向きの分からない科目はそのまま */
export function toDebitPlus(facts: Fact[], traits: Map<string, AccountTraits>, signs: Map<string, EventSign>): Fact[] {
    return facts.map((fact) => {
        const trait = traits.get(fact.code);
        const sign = signs.get(fact.event);
        if (!trait?.amount || trait.credit === null || !sign) return fact;
        const factor = trait.credit ? sign.credit : sign.debit;
        return factor === 1 ? fact : { ...fact, value: -fact.value || 0, last: -fact.last || 0 };
    });
}
