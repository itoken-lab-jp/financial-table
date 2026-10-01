/**
 * 科目マスタ（実在の科目だけ）から、表の行を組む。小計・率・見出し・空行はビジュアルが作る。
 *
 * - 名前に意味を持たせない。区分の名前から型（売上高・流動資産…）を
 *   当てることはしない。区分の並びは区分マスタの区分の並び、無ければ科目の並び（区分の最初の科目）
 * - 区分の中は、中分類ごとの小計 → 科目。科目が 1 つで区分と同じ名前なら、小計を付けずにその行を区分の行にする
 * - 計算の行：段階利益・合計・比率は、すべて書式ペインの「計算の行」で足す。小計はどこからの区分から置く場所の区分まで、
 *   比率は分子の行 ÷ 分母の行
 * - 指標の行：指標の欄のメジャーを、書式ペインのメジャーごとの置く場所に置く。行の後ろ（人数・時間・EBITDA
 *   のような会社独自の行。届いた値のまま、集計・書式・良し悪しはメジャーごとの設定）か、行の「うち」（親の行と同じ見せ方）。
 *   科目の表の「内訳の親」と「書式」の欄はやめた（ファクトが足せない形と、ディメンションに見せ方を持たせる形になるため）
 * - 科目の性質は 2 つのフラグ：貸方フラグ（正常な残高が貸方）と残高フラグ（期末を取る）。
 *   空の貸方フラグは区分の向きを継ぎ、空の残高フラグはフロー。収益・費用・資産…の分類は持たない
 * - 符号：データは借方プラスにそろえてから受け取る（signs.ts）。区分は自分の向きでプラスに見せる。区分の向きは区分の中の科目の
 *   貸方フラグの金額の多数（科目の数ではなく。値引の科目が多くても売上の向き）。向きの逆の科目（評価勘定：減価償却累計額・
 *   貸倒引当金・自己株式・売上値引）は ▲ で見せる（行を足すと小計になる）。小計の行は足す範囲の最初の区分の向き
 * - 良し悪し：貸方のフローは増えると良い、借方のフローは増えると悪い、残高は中立。評価勘定は区分の良し悪しのまま
 * - 比率の行は、分子と分母の行を期間で集計してから割る
 * - 並び：区分の中は科目の並び → コード。小計を科目の上か下に置くかは表全体で選ぶ。うちは親の行の下
 */
import { AccountRecord, INDICATOR_KEY } from "./data";
import { Expr, FormulaError, parseFormula } from "./formula";
import { AMOUNT_FORMAT, RowFormat, parseRowFormat } from "./numberFormat";

/** others は見る人が選ばなかった科目をまとめた「その他」の行（picks.ts） */
export type RowType = "detail" | "subtotal" | "step" | "calc" | "measure" | "breakdown" | "others" | "blank" | "heading";
export type Aggregation = "flow" | "stock" | "calc";
/** 1 = 増えて良い、-1 = 増えて悪い、0 = 色を付けない */
export type GoodDirection = 1 | -1 | 0;
export type ParentPosition = "below" | "above";

/** 足す行と重み（借方プラスのデータの値を足すときの係数） */
export interface Summand {
    code: string;
    weight: number;
}

export interface RowDef {
    code: string;
    name: string;
    type: RowType;
    /** 表示の値 = 足した値（借方プラス）× sign。貸方の向きの行は −1 */
    sign: 1 | -1;
    good: GoodDirection;
    aggregation: Aggregation;
    format: RowFormat;
    /** 率の行の式（[分子] / [分母]） */
    formula: Expr | null;
    /** 式が引く行のコード */
    refs: string[];
    /** 子（表示の並び順）。囲みの帯はここから引く */
    children: string[];
    /** 小計・段階の行で足す行 */
    summands: Summand[];
    /** 段階の行のうち、後の段階の行が足す範囲に自分の範囲を含むもの（売上総利益 → 営業利益 → …の途中）。線は二重線でなく小計の線 */
    continued?: boolean;
}

export interface DisplayRow {
    def: RowDef;
    /** 木の深さ。一番上の段が 0 */
    depth: number;
}

/** 科目の性質（2 つのフラグ。空の貸方フラグは区分の向き、空の残高フラグはフロー）。符号の持ち方をそろえるのに使う（signs.ts） */
export interface AccountTraits {
    /** 正常な残高が貸方か。分からなければ null */
    credit: boolean | null;
    /** 残高（期末を取る）か */
    stock: boolean;
    /** 金額の科目か：足さない区分でない。符号の持ち方はこの科目だけに効く */
    amount: boolean;
    /** 区分の合計に足す科目か（区分と残高かフローかの違う科目・うちは足さない）。符号の持ち方の見分けはこの科目だけで見る */
    inTotal: boolean;
}

/** 書式ペインの選択肢。value は保存する値 */
export interface CalcChoices {
    /** 計算の行の置く場所・どこからの区分（表の並び） */
    sections: Array<{ value: string; displayName: string }>;
    /**
     * 比率の分子・分母に使える行（区分は「§sec:区分」、書式ペインの小計は「§calc:本目」、指標は「§ind:queryName」、ほかは行のコード。
     * 比率の行は入れない）
     */
    refs: Array<{ value: string; displayName: string }>;
    /** 指標の置く場所（表の行。見出しの区分も「§sec:区分」。空行は入れない） */
    places: Array<{ value: string; displayName: string }>;
}

export interface RowModel {
    rows: Map<string, RowDef>;
    display: DisplayRow[];
    warnings: string[];
    /** 科目のコード → 性質（金額の行のうちも入れる） */
    traits: Map<string, AccountTraits>;
    calcChoices: CalcChoices;
    /** 行のうち（親の行のコード → うちの行のコード）。合計の行・計算の行の小計のうちは子（children）に入らないので、折りたたみはここも見る */
    attached: Map<string, string[]>;
    /** 行の後ろに置いた指標（置く場所の行のコード → 指標の行のコード）。折りたたみで、子の後ろに置いた指標も一緒に隠す */
    following: Map<string, string[]>;
}

/** 書式ペインで足す計算の行（「計算の行」カードの編集する行）。小計はどこからの区分から置く場所の区分までの合計、比率は分子の行 ÷ 分母の行 */
export interface CalcRowSpec {
    /** 書式ペインの何本目か（1 から。警告に出す） */
    slot: number;
    kind: "subtotal" | "ratio";
    name: string;
    /** 小計の足し始めの区分。null は表の最初 */
    from: string | null;
    /** この区分の後に置く（小計はここまで足す）。null は表の最後 */
    after: string | null;
    /** 比率の分子・分母の行（CalcChoices.refs の value） */
    numerator: string | null;
    denominator: string | null;
    /** 比率の書式。空なら 0.0% */
    format: string | null;
    /** 比率の良し悪し（小計は足す範囲の最初の区分の向き） */
    good: GoodDirection;
}

/** 指標の行。書式ペインの「指標」カードのメジャーごとの設定から作る（viewModel） */
export interface IndicatorSpec {
    /** 行のコード（INDICATOR_KEY + queryName） */
    code: string;
    name: string;
    /** 置く場所の行（CalcChoices.places の value）。null は表の最後 */
    target: string | null;
    /** after：置く場所の行（とその下の行）の後ろ。under：置く場所の行の「うち」（親の行と同じ見せ方で、集計・書式・良し悪しも親に合わせる） */
    mode: "after" | "under";
    /** 後ろに置くときの期間の集計（合計・期末） */
    aggregation: "flow" | "stock";
    /** 後ろに置くときの書式（空は金額） */
    format: RowFormat;
    good: GoodDirection;
}

/** 分子・分母・置く場所の行のキーで、区分の行を指すもの（区分の科目が 1 つだと行のコードが科目のコードになるので、区分の名前で持つ） */
export const SECTION_KEY = "§sec:";
/** 書式ペインで足した行のコード（何本目か）。名前から作ると、同じ名前の小計とぶつかって付け替わる */
export const CALC_KEY = "§calc:";

export interface BuildOptions {
    position: ParentPosition;
    /** 書式ペインで足す計算の行 */
    calcRows?: CalcRowSpec[];
    /** 指標の行（指標の欄の並び） */
    indicators?: IndicatorSpec[];
    /** 区分マスタの区分の並びの列を入れたか。入れたら、並びの空の区分を知らせる。無ければ値で見る */
    sectionMaster?: boolean;
    /** 貸方フラグ・残高フラグの列を入れたか。無ければ値で見る。列が無ければ知らせる（向きが分からない・残高が分からない） */
    flags?: { credit: boolean; balance: boolean };
    /** 科目の金額の大きさ（すべての月・イベントの絶対値の和）。区分の向きの多数を金額で見る。無ければ科目の数で見る */
    magnitude?: (code: string) => number;
}

const DEFAULT_OPTIONS: BuildOptions = { position: "below" };

/** 全角の数字・記号を半角に（Excel の日本語入力で入る －1・○ など） */
function halfWidth(value: string): string {
    return value.normalize("NFKC").replace(/[−–—ー‐]/g, "-");
}

/** フラグ。ON（1・true・はい・○・✓）は true、OFF（0・false・いいえ・×）は false、空と読めないものは null */
export function parseFlag(value: string | null | undefined): boolean | null {
    if (value === null || value === undefined) return null;
    const v = halfWidth(value).trim().toLowerCase();
    if (v === "") return null;
    if (/^(1|true|yes|y|on|はい|有|あり|○|◯|●|✓|✔|〇)$/.test(v)) return true;
    if (/^(0|false|no|n|off|いいえ|無|なし|×|✕|-)$/.test(v)) return false;
    return null;
}

/** コードの並び（並びが同じか無いとき）。数字どうしは数として比べる */
const compareCode = (a: string, b: string) => a.localeCompare(b, "ja", { numeric: true });

const PERCENT_FORMAT: RowFormat = parseRowFormat("0.0%").format;

/** 区分（読んだものをまとめたもの） */
interface Section {
    label: string;
    /** 区分の向き：正常な残高が貸方か（科目の貸方フラグの金額の多数）。分からなければ null */
    credit: boolean | null;
    /** 残高の区分か */
    stock: boolean;
    /** 足さない区分（区分なし・向きの分からない区分）。見出しの下に並べるだけ */
    other: boolean;
    /** 区分の科目 */
    accounts: AccountRecord[];
    /** 区分マスタの行（区分の科目から、空でない値を集めたもの） */
    master: AccountRecord["section"];
}

/** 区分の空の科目を置く区分。区分マスタの「その他」などとぶつからない名前 */
export const NO_CATEGORY = "（区分なし）";

/** 行の向きの符号：貸方の向きは −1（借方プラスのデータを反転して見せる） */
const signOf = (credit: boolean | null): 1 | -1 => (credit ? -1 : 1);
/** 良し悪し：貸方のフローは +、借方のフローは −、残高は色なし */
const goodOf = (credit: boolean | null, stock: boolean): GoodDirection => (stock || credit === null ? 0 : credit ? 1 : -1);

/** 値を持つ行か（見出し・空行でない） */
const hasValue = (type: RowType) => type !== "heading" && type !== "blank";

export function buildRows(accounts: Map<string, AccountRecord>, options: Partial<BuildOptions> = {}): RowModel {
    const opts: BuildOptions = { ...DEFAULT_OPTIONS, ...options };
    const warnings: string[] = [];
    const warn = (message: string) => {
        if (!warnings.includes(message)) warnings.push(message);
    };
    const list = (names: string[]) => names.slice(0, 3).join("・") + (names.length > 3 ? " ほか" : "");
    const rows = new Map<string, RowDef>();
    const display: DisplayRow[] = [];
    const traits = new Map<string, AccountTraits>();
    const add = (def: Omit<RowDef, "children" | "summands" | "refs" | "formula"> & Partial<RowDef>): RowDef => {
        const row: RowDef = { children: [], summands: [], refs: [], formula: null, ...def };
        rows.set(row.code, row);
        return row;
    };
    /** 同じコードが既にあれば番号を付ける（見出し・段階の行・率の行の名前がぶつかっても上書きしない） */
    const uniqueCode = (base: string) => {
        let code = base;
        for (let n = 2; rows.has(code); n++) code = `${base}:${n}`;
        return code;
    };
    const byOrder = (a: AccountRecord, b: AccountRecord) => {
        const oa = a.order ?? Number.POSITIVE_INFINITY;
        const ob = b.order ?? Number.POSITIVE_INFINITY;
        return oa !== ob ? oa - ob : compareCode(a.code, b.code);
    };
    const sorted = Array.from(accounts.values()).sort(byOrder);
    const flags = opts.flags ?? {
        credit: sorted.some((a) => (a.creditFlag ?? null) !== null),
        balance: sorted.some((a) => (a.balanceFlag ?? null) !== null),
    };

    // ---- 区分に分ける
    const sectionMap = new Map<string, Section>();
    const noCategory: string[] = [];
    for (const account of sorted) {
        const label = account.category ?? NO_CATEGORY;
        if (account.category === null) noCategory.push(account.name);
        let section = sectionMap.get(label);
        if (!section) {
            section = { label, credit: null, stock: false, other: false, accounts: [], master: { ...account.section } };
            sectionMap.set(label, section);
        }
        // 区分マスタの列は、区分の科目のどれかに入っていればよい
        const master = section.master as unknown as Record<string, unknown>;
        for (const [field, value] of Object.entries(account.section)) if (master[field] === null && value !== null) master[field] = value;
        section.accounts.push(account);
    }
    if (noCategory.length > 0) warn(`区分の空の科目がある（${list(noCategory)}）。「${NO_CATEGORY}」として最後に並べた`);

    // 区分の並びの列を入れたか。列の有無が分からないときは値で見る
    const useMaster = opts.sectionMaster ?? Array.from(accounts.values()).some((a) => a.section.order !== null);

    // ---- 科目のフラグを読む
    const creditFlag = new Map<string, boolean | null>();
    const balanceFlag = new Map<string, boolean | null>();
    for (const account of sorted) {
        const credit = parseFlag(account.creditFlag);
        if ((account.creditFlag ?? null) !== null && credit === null) warn(`貸方フラグ「${account.creditFlag}」が読めない（${account.name}）`);
        const balance = parseFlag(account.balanceFlag);
        if ((account.balanceFlag ?? null) !== null && balance === null) warn(`残高フラグ「${account.balanceFlag}」が読めない（${account.name}）`);
        creditFlag.set(account.code, credit);
        balanceFlag.set(account.code, balance);
    }

    // ---- 区分の性質は、科目のフラグの多数で決める（名前では決めない）
    /** 多数。null のフラグは数えない。同数は null */
    const majority = (
        accounts: AccountRecord[],
        flagOf: (a: AccountRecord) => boolean | null,
        weightOf: (a: AccountRecord) => number
    ): boolean | null => {
        let yes = 0;
        let no = 0;
        for (const a of accounts) {
            const v = flagOf(a);
            if (v === true) yes += weightOf(a);
            else if (v === false) no += weightOf(a);
        }
        return yes > no ? true : no > yes ? false : null;
    };
    /**
     * 向きの重みは金額（すべての月・イベントの絶対値の和）。科目の数だと、値引・戻り・割戻の科目が多い区分や、フィルターで顔ぶれが
     * 変わった区分が反転しうる。金額が無ければ（どの科目も 0）科目の数
     */
    const size = (a: AccountRecord) => opts.magnitude?.(a.code) ?? 0;
    const unknownDirection: string[] = [];
    const unknownStock: string[] = [];
    for (const section of sectionMap.values()) {
        const candidates = section.accounts;
        // 残高の区分か：科目の数の多数（空の残高フラグはフローに数え、同数はフロー）。金額で数えると、付け忘れた大きな科目 1 つで
        // 区分ごとフローになる。空を数えないと、フローの区分に残高の科目が 1 つ混ざっただけで区分ごと残高になる。
        // 区分と合わない科目は足さずに名指しで知らせる（emitSection）
        section.stock = majority(candidates, (a) => balanceFlag.get(a.code) ?? false, () => 1) === true;
        // 向きの票は、区分と残高かフローかの同じ科目だけ（違う科目は区分の合計に足さない）。空の貸方フラグは数えない。
        // 同数は、並びで最初にフラグを入れた科目
        const voters = candidates.filter((a) => (balanceFlag.get(a.code) ?? false) === section.stock);
        const flagOf = (a: AccountRecord) => creditFlag.get(a.code) ?? null;
        section.credit = majority(voters, flagOf, voters.some((a) => size(a) > 0) ? size : () => 1) ?? voters.map(flagOf).find((v) => v !== null) ?? null;
        // 足さない区分：区分なし・向きの分からない区分。名前（その他など）では決めない
        const noCategory = section.label === NO_CATEGORY;
        section.other = noCategory || section.credit === null;
        // 向きの分からない区分：科目があるのに貸方フラグがどれも空の区分
        if (!noCategory && section.credit === null) unknownDirection.push(section.label);
        if (!section.other && !flags.balance) unknownStock.push(section.label);
    }
    if (unknownDirection.length > 0) {
        warn(
            `区分「${unknownDirection.slice(0, 3).join("」「")}」の向き（正常な残高が貸方か借方か）が分からない（区分の科目の貸方フラグがどれも空）。区分を足さずに、見出しの下に並べた。${
                flags.credit ? "科目の表の貸方フラグを埋める" : "科目の表に貸方フラグの列を入れる"
            }（収益・負債・純資産は ON、費用・資産は OFF。空は OFF でなく、区分の向きを継ぐ）`
        );
    }
    if (unknownStock.length > 0) {
        warn(
            `残高フラグの列が無いので、区分（${list(unknownStock)}）をフローとして四半期・通期で足した。貸借対照表のような残高（期末の値）の科目があれば、科目の表に残高フラグの列を入れる（損益だけの表なら、列を入れて空にしておけば知らせない）`
        );
    }

    // ---- 科目の性質（signs.ts が符号の持ち方をそろえるのに使う）
    const blankCredit: string[] = [];
    for (const section of sectionMap.values()) {
        for (const account of section.accounts) {
            // 空のフラグは区分の向き・フロー（残高は区分の多数では決めない）
            const stock = balanceFlag.get(account.code) ?? false;
            const amount = !section.other;
            traits.set(account.code, {
                credit: creditFlag.get(account.code) ?? section.credit,
                stock,
                amount,
                inTotal: amount && stock === section.stock,
            });
        }
        // 区分の合計に足す金額の科目で、貸方フラグが空のもの（区分の向きを継いだ。名前から向きを当てないので、付け忘れなら区分の向きが崩れる）
        if (!section.other) {
            for (const a of section.accounts) if (traits.get(a.code)?.inTotal && creditFlag.get(a.code) === null) blankCredit.push(a.name);
        }
    }
    if (blankCredit.length > 0) {
        warn(`貸方フラグが空の科目は、区分の向きにした（${list(blankCredit)}）。区分と逆の科目（評価勘定）なら符号が逆になるので、科目の表の貸方フラグを埋める`);
    }

    /** 区分の値の行のコード（段階の行が足す行） */
    const sectionValue = new Map<string, string>();
    /** 足さない区分の見出しの行のコード（指標の置く場所） */
    const sectionHead = new Map<string, string>();
    /** 中分類の小計の行のコード → 区分（書式ペインの選択肢で、区分をまたいで同じ名前の中分類を見分ける） */
    const groupSection = new Map<string, string>();
    const nameIndex = new Map<string, string | null>();
    const noteName = (name: string, code: string) => nameIndex.set(name, nameIndex.has(name) && nameIndex.get(name) !== code ? null : code);
    const displayed = new Set<string>();

    /**
     * 科目の行。区分の合計に足す科目は区分の向きで見せる（評価勘定は ▲）。足さない区分の科目は届いた値のまま・色なし。
     * 区分の合計に足さない科目（損益の区分の残高フラグの科目など）は自分の向きで見せる
     */
    const accountRow = (account: AccountRecord, section: Section): RowDef => {
        const trait = traits.get(account.code) ?? { credit: section.credit, stock: section.stock, amount: false, inTotal: false };
        noteName(account.name, account.code);
        noteName(account.code, account.code);
        displayed.add(account.code);
        const asIs = section.other || !trait.amount;
        const own = !asIs && !trait.inTotal;
        return add({
            code: account.code,
            name: account.name,
            type: "detail",
            sign: asIs ? 1 : signOf(own ? trait.credit : section.credit),
            good: asIs || trait.stock ? 0 : own ? goodOf(trait.credit, false) : goodOf(section.credit, section.stock),
            aggregation: trait.stock ? "stock" : "flow",
            format: AMOUNT_FORMAT,
        });
    };

    const emitAccount = (account: AccountRecord, section: Section, depth: number): RowDef => {
        const row = accountRow(account, section);
        display.push({ def: row, depth });
        return row;
    };

    const heading = (section: Section) => {
        const row = add({
            code: uniqueCode(`§head:区分:${section.label}`),
            name: section.label,
            type: "heading",
            sign: 1,
            good: 0,
            aggregation: "flow",
            format: AMOUNT_FORMAT,
            children: section.accounts.map((a) => a.code),
        });
        sectionHead.set(section.label, row.code);
        return row;
    };

    /** 区分の行を組んで、表示の並びに足す。段階の行が足す値の行は sectionValue に残す（足さない区分は残さない） */
    /** 表に出した区分（表の並び）。計算の行の小計の範囲と、書式ペインの置く場所の選択肢 */
    const emitted: Section[] = [];
    const emitSection = (section: Section): void => {
        emitted.push(section);
        unflushed.push(section);
        const main = section.accounts;
        if (section.other) {
            // 足さない区分なので、見出しの下に並べるだけ。科目が 1 つで区分と同じ名前なら見出しも付けない
            if (main.length === 1 && main[0].name === section.label) {
                emitAccount(main[0], section, 0);
                return;
            }
            display.push({ def: heading(section), depth: 0 });
            for (const account of main) emitAccount(account, section, 1);
            return;
        }

        // 区分の合計に足さない科目：区分と残高・フローの違う科目（損益の区分の残高フラグの科目など。期末の値を月ごとに足すと嘘になる）
        const mismatch = main.filter((a) => traits.get(a.code)?.stock !== section.stock);
        if (mismatch.length > 0) {
            warn(
                section.stock
                    ? `残高フラグが OFF か空の科目は、残高の区分「${section.label}」の合計と段階の行に足さない（${list(mismatch.map((a) => a.name))}）`
                    : `残高フラグの科目は、区分「${section.label}」の合計と段階の行に足さない（${list(mismatch.map((a) => a.name))}）`
            );
        }
        const summable = (code: string) => !mismatch.some((a) => a.code === code);

        // 足せる科目が無い区分は、その他の区分と同じく見出しの下に並べるだけ（小計は 0 と出てしまう）
        if (!main.some((a) => summable(a.code))) {
            const bare = main.length === 1 && main[0].name === section.label;
            if (!bare) display.push({ def: heading(section), depth: 0 });
            for (const account of main) emitAccount(account, section, bare ? 0 : 1);
            return;
        }

        // 科目が 1 つで区分と同じ名前：小計を付けず、その科目を区分の行にする
        if (main.length === 1 && main[0].name === section.label && main[0].subCategory === null) {
            emitAccount(main[0], section, 0);
            sectionValue.set(section.label, main[0].code);
            noteName(section.label, main[0].code);
            return;
        }

        const sumRow = (code: string, name: string, children: string[]): RowDef =>
            add({
                code,
                name,
                type: "subtotal",
                sign: signOf(section.credit),
                good: goodOf(section.credit, section.stock),
                aggregation: section.stock ? "stock" : "flow",
                format: AMOUNT_FORMAT,
                children,
                summands: children.filter(summable).map((c) => ({ code: c, weight: 1 })),
            });

        // 中分類ごとの小計。並びは中分類の中の最初の科目の位置
        const groups: Array<{ name: string | null; accounts: AccountRecord[] }> = [];
        for (const account of main) {
            const group = groups.find((g) => g.name !== null && g.name === account.subCategory);
            if (group) group.accounts.push(account);
            else groups.push({ name: account.subCategory, accounts: [account] });
        }
        const sectionCode = uniqueCode(`§sec:${section.label}`);
        const children: string[] = [];
        const body: Array<() => void> = [];
        for (const group of groups) {
            if (group.name === null) {
                const account = group.accounts[0];
                children.push(account.code);
                body.push(() => emitAccount(account, section, 1));
                continue;
            }
            if (!group.accounts.some((a) => summable(a.code))) {
                // 足せる科目の無い中分類は小計を付けない（足すものが無い）
                for (const account of group.accounts) {
                    children.push(account.code);
                    body.push(() => emitAccount(account, section, 1));
                }
                continue;
            }
            const groupCode = uniqueCode(`§grp:${section.label}:${group.name}`);
            groupSection.set(groupCode, section.label);
            const groupChildren = group.accounts.map((a) => a.code);
            children.push(groupCode);
            body.push(() => {
                const groupRow = sumRow(groupCode, group.name!, groupChildren);
                noteName(group.name!, groupCode);
                if (opts.position === "above") display.push({ def: groupRow, depth: 1 });
                for (const account of group.accounts) emitAccount(account, section, 2);
                if (opts.position === "below") display.push({ def: groupRow, depth: 1 });
            });
        }
        const sectionRow = sumRow(sectionCode, section.label, children);
        if (opts.position === "above") display.push({ def: sectionRow, depth: 0 });
        body.forEach((emit) => emit());
        if (opts.position === "below") display.push({ def: sectionRow, depth: 0 });
        sectionValue.set(section.label, sectionCode);
        noteName(section.label, sectionCode);
    };

    /**
     * 小計の行（書式ペインの計算の行の小計）。借方プラスのデータの値を足し、足す範囲の最初の区分の向きで見せる
     * （営業利益 = −(売上高 + 売上原価 + 販管費)、資産合計 = 流動資産 + 固定資産）。期末かどうかは足す区分で決める
     */
    const emitSubtotal = (name: string, sections: Section[], codeBase: string): RowDef | null => {
        // 向きの分からない区分（科目があるのに貸方フラグがどれも空）は足せない。飛ばして足すと、売上原価を引かない売上総利益の
        // ような値になるので、行ごと出さない（収益にだけ ON を入れ、費用を空にした科目の表で起きる）
        const unknown = sections.filter((sec) => unknownDirection.includes(sec.label));
        if (unknown.length > 0) {
            warn(`計算の行「${name}」は、向きの分からない区分（${list(unknown.map((sec) => sec.label))}）を足す範囲に含むので出さなかった。区分の科目の貸方フラグを埋める`);
            return null;
        }
        const summed = sections.filter((sec) => !sec.other && sectionValue.has(sec.label));
        if (summed.length === 0) {
            warn(`計算の行「${name}」に足す区分が無い（その他の区分だけ）。出さなかった`);
            return null;
        }
        const stock = summed.every((sec) => sec.stock);
        if (!stock && summed.some((sec) => sec.stock)) {
            warn(
                `計算の行「${name}」は、フローの区分と残高の区分をまたいで足している。四半期・通期では残高も月ごとに足される。計算の行の「どこから」で、損益と貸借対照表の間を切る`
            );
        }
        const credit = summed[0].credit;
        const code = uniqueCode(codeBase);
        const row = add({
            code,
            name,
            type: "step",
            sign: signOf(credit),
            good: goodOf(credit, stock),
            aggregation: stock ? "stock" : "flow",
            format: AMOUNT_FORMAT,
            summands: summed.map((sec) => ({ code: sectionValue.get(sec.label)!, weight: 1 })),
        });
        noteName(name, code);
        display.push({ def: row, depth: 0 });
        return row;
    };

    /** 比率の行。分子・分母は行のキー（区分は「§sec:区分」、ほかは行のコード）。行を組んだあとで結ぶ */
    const pendingRatios: Array<{ row: RowDef; numerator: string; denominator: string }> = [];
    const emitRatio = (name: string, numerator: string, denominator: string, format: RowFormat, good: GoodDirection, codeBase: string): RowDef => {
        const row = add({ code: uniqueCode(codeBase), name, type: "calc", sign: 1, good, aggregation: "calc", format });
        display.push({ def: row, depth: 0 });
        pendingRatios.push({ row, numerator, denominator });
        return row;
    };
    /** 比率の書式。比率の値は表示単位で割らない（客単価 1,770 円を千円で割らない）。字の無い形は、その桁の数として出す */
    const ratioFormat = (text: string | null, name: string): RowFormat => {
        if (!text) return PERCENT_FORMAT;
        const parsed = parseRowFormat(text);
        if (parsed.ok) return parsed.format.kind === "amount" ? { ...parsed.format, kind: "number" } : parsed.format;
        warn(`計算の行「${name}」の書式「${text}」が読めない。0.0% にした`);
        return PERCENT_FORMAT;
    };

    // ---- 書式ペインの計算の行。置く場所の区分の後か表の最後に、書式ペインの並びで出す（次の区分を出す直前、flushCalc）
    const calcSpecs = (opts.calcRows ?? []).filter((spec) => {
        if (spec.name.trim() !== "") return true;
        warn(`計算の行 ${spec.slot} の名前が空。出さなかった`);
        return false;
    });
    const calcDone = new Set<CalcRowSpec>();
    const emitCalc = (spec: CalcRowSpec) => {
        calcDone.add(spec);
        if (spec.kind === "ratio") {
            if (!spec.numerator || !spec.denominator) {
                warn(`計算の行「${spec.name}」の分子か分母が空。出さなかった`);
                return;
            }
            emitRatio(spec.name, spec.numerator, spec.denominator, ratioFormat(spec.format, spec.name), spec.good, CALC_KEY + spec.slot);
            return;
        }
        // 小計：どこからの区分から、置く場所の区分まで（表の並び。どこからの既定は表の最初、置く場所は表の最後も選べる）
        const end = spec.after === null ? emitted.length - 1 : emitted.findIndex((sec) => sec.label === spec.after);
        const start = spec.from === null ? 0 : emitted.findIndex((sec) => sec.label === spec.from);
        if (start < 0 || start > end) {
            warn(
                spec.from !== null && sectionMap.has(spec.from)
                    ? `計算の行「${spec.name}」のどこからの区分「${spec.from}」が、置く場所より後にある。出さなかった`
                    : `計算の行「${spec.name}」のどこからの区分「${spec.from}」が表に無い。出さなかった`
            );
            return;
        }
        emitSubtotal(spec.name, emitted.slice(start, end + 1), CALC_KEY + spec.slot);
    };
    /** 出したが、まだ計算の行を出していない区分 */
    let unflushed: Section[] = [];
    const flushCalc = () => {
        const done = unflushed;
        unflushed = [];
        for (const section of done) for (const spec of calcSpecs) if (!calcDone.has(spec) && spec.after === section.label) emitCalc(spec);
    };

    let blanks = 0;
    const blank = () => display.push({ def: add({ code: `§blank:${blanks++}`, name: "", type: "blank", sign: 1, good: 0, aggregation: "flow", format: AMOUNT_FORMAT }), depth: 0 });

    // ---- 区分を並べる：区分マスタの区分の並び → 区分の最初の科目の並び。並びの空の区分（区分の並びの列を入れたとき）と区分なしは最後
    const sections = Array.from(sectionMap.values());
    const unlinked = useMaster ? sections.filter((sec) => sec.master.order === null && sec.label !== NO_CATEGORY) : [];
    if (unlinked.length > 0) {
        warn(`区分の並びが空の区分がある（${list(unlinked.map((sec) => sec.label))}）。最後に並べた。区分マスタの区分の名前を科目の区分とそろえる`);
    }
    const last = (sec: Section) => (unlinked.includes(sec) ? 1 : 0) + (sec.label === NO_CATEGORY ? 2 : 0);
    sections.sort(
        (a, b) =>
            last(a) - last(b) || (a.master.order ?? Number.POSITIVE_INFINITY) - (b.master.order ?? Number.POSITIVE_INFINITY) || byOrder(a.accounts[0], b.accounts[0])
    );
    for (const section of sections) {
        flushCalc();
        emitSection(section);
    }
    flushCalc();

    // 表の最後に置く計算の行。置く場所の区分が表に無い計算の行は出さない（フィルターで区分が消えたときも）。
    // 表が空（フィルターで科目が無い）なら、どの行も出さず、知らせもしない
    for (const spec of emitted.length > 0 ? calcSpecs : []) {
        if (calcDone.has(spec)) continue;
        if (spec.after === null) emitCalc(spec);
        else warn(`計算の行「${spec.name}」の置く場所の区分「${spec.after}」が表に無い。出さなかった`);
    }

    // ---- 表に置けなかった科目は無いはずだが、落とさずに最後に出して知らせる
    const missing = sorted.filter((a) => !displayed.has(a.code));
    if (missing.length > 0) {
        warn(`表に置けなかった科目がある（${list(missing.map((a) => a.name))}）。足さずに最後に並べた`);
        const other: Section = { label: "（置けなかった科目）", credit: null, stock: false, other: true, accounts: missing, master: missing[0].section };
        blank();
        display.push({ def: add({ code: uniqueCode("§head:置けなかった科目"), name: other.label, type: "heading", sign: 1, good: 0, aggregation: "flow", format: AMOUNT_FORMAT }), depth: 0 });
        for (const account of missing) display.push({ def: accountRow(account, other), depth: 1 });
    }

    // ---- 指標の行：書式ペインのメジャーごとの置く場所に置く。置く場所が別の指標なら、その指標を置いてから置く
    const indicatorNames = new Map((opts.indicators ?? []).map((spec) => [spec.code, spec.name]));
    /** 置く場所のキー → 行のコード。区分は区分の値の行（足さない区分は見出し） */
    const placeOf = (key: string): string | undefined => {
        if (key.startsWith(SECTION_KEY)) {
            const label = key.slice(SECTION_KEY.length);
            return sectionValue.get(label) ?? sectionHead.get(label);
        }
        return rows.has(key) ? key : undefined;
    };
    /** 置く場所の名前（警告に出す。書式ペインの選択肢の名前にそろえる） */
    const placeLabel = (key: string) =>
        key.startsWith(SECTION_KEY)
            ? key.slice(SECTION_KEY.length)
            : key.startsWith(CALC_KEY)
              ? `計算の行 ${key.slice(CALC_KEY.length)}`
              : key.startsWith(INDICATOR_KEY)
                ? (indicatorNames.get(key) ?? key.slice(INDICATOR_KEY.length))
                : (rows.get(key)?.name ?? key);
    /**
     * 行のうち（親の行のコード → うちの行のコード）。区分・中分類の合計の行と計算の行の小計のうちは、親の子（children）にしない
     * （子にすると囲みの帯が合計の行の下まで伸び、下囲みの底が合計の行でなくなる）ので、ここで持つ
     */
    const attached = new Map<string, string[]>();
    /** 行と、その下の行（子・孫・うち）の並びの最後の位置 */
    const blockEnd = (code: string): number => {
        const positions = new Map(display.map((d, i): [string, number] => [d.def.code, i]));
        let end = positions.get(code) ?? -1;
        const walk = (parent: string) => {
            for (const child of [...(rows.get(parent)?.children ?? []), ...(attached.get(parent) ?? [])]) {
                const at = positions.get(child);
                if (at !== undefined && at > end) end = at;
                walk(child);
            }
        };
        walk(code);
        return end;
    };
    /** 行の後ろに置いた指標（置く場所の行のコード → 指標の行のコード） */
    const following = new Map<string, string[]>();
    /** 置く場所の行の後ろに最後に置いた指標（同じ行の後ろに置いた指標を、書式ペインの並びで並べる） */
    const afterTail = new Map<string, string>();
    /** 行のうちに最後に置いた指標（同じ行のうちを、書式ペインの並びで並べる） */
    const underTail = new Map<string, string>();
    const placeIndicator = (spec: IndicatorSpec, target: string | null) => {
        const parent = target === null ? undefined : rows.get(target);
        let mode = spec.mode;
        if (mode === "under" && parent && (parent.type === "calc" || parent.type === "heading")) {
            warn(`指標「${spec.name}」は「${parent.name}」のうちにできない（比率・見出しの行）。後ろに置いた`);
            mode = "after";
        }
        if (mode === "under" && parent) {
            // うち：親の行の下（親のうちの後ろ）に 1 段深く。見せ方は親に合わせる
            let row: RowDef;
            if (parent.type === "measure" || (parent.type === "breakdown" && !traits.has(parent.code))) {
                // 指標の行のうち：値は届いたまま。集計・書式・良し悪しは親の指標と同じ
                row = add({ code: spec.code, name: spec.name, type: "breakdown", sign: 1, good: parent.good, aggregation: parent.aggregation, format: parent.format });
            } else {
                // 金額の行のうち：値は金額と同じ持ち方（CALCULATE([金額], 製品[区分] = "新製品") の形）なので、親の性質で借方プラスにそろえ
                // （signs.ts）、親の行と同じ向きで見せる。区分の合計にも親の合計にも足さない
                const parentTrait = traits.get(parent.code);
                const stock = parent.aggregation === "stock";
                traits.set(spec.code, { credit: parentTrait ? parentTrait.credit : parent.sign === -1, stock, amount: parentTrait ? parentTrait.amount : true, inTotal: false });
                row = add({ code: spec.code, name: spec.name, type: "breakdown", sign: parent.sign, good: parent.good, aggregation: stock ? "stock" : "flow", format: parent.format });
            }
            const depth = display.find((d) => d.def.code === parent.code)!.depth + 1;
            display.splice(blockEnd(underTail.get(parent.code) ?? parent.code) + 1, 0, { def: row, depth });
            underTail.set(parent.code, row.code);
            attached.set(parent.code, [...(attached.get(parent.code) ?? []), row.code]);
            // 科目・指標の行のうちは子にする（区分の囲みの帯に入る）。合計の行・計算の行の小計のうちは子にしない
            if (parent.type !== "subtotal" && parent.type !== "step") parent.children.push(row.code);
            noteName(spec.name, spec.code);
            return;
        }
        const row = add({ code: spec.code, name: spec.name, type: "measure", sign: 1, good: spec.good, aggregation: spec.aggregation, format: spec.format });
        noteName(spec.name, spec.code);
        if (!parent) {
            display.push({ def: row, depth: 0 });
            return;
        }
        // 後ろ：置く場所の行とその下の行（子・うち）の後ろに、置く場所の行と同じ深さで
        const depth = display.find((d) => d.def.code === parent.code)!.depth;
        display.splice(blockEnd(afterTail.get(parent.code) ?? parent.code) + 1, 0, { def: row, depth });
        afterTail.set(parent.code, row.code);
        following.set(parent.code, [...(following.get(parent.code) ?? []), row.code]);
    };
    // 表が空（フィルターで科目が無い）なら、指標も出さず、知らせない（計算の行と同じ）
    let pending = emitted.length > 0 ? [...(opts.indicators ?? [])] : [];
    while (pending.length > 0) {
        const waiting: IndicatorSpec[] = [];
        for (const spec of pending) {
            if (rows.has(spec.code)) {
                warn(`指標「${spec.name}」が 2 つある。最初のものを出した`);
                continue;
            }
            if (spec.target === null) {
                placeIndicator(spec, null);
                continue;
            }
            const target = placeOf(spec.target);
            if (target !== undefined) placeIndicator(spec, target);
            // 置く場所の指標をまだ置いていない（後ろの指標のうち・後ろに置いた）なら、次の回で置く
            else if (pending.some((p) => p !== spec && p.code === spec.target)) waiting.push(spec);
            else warn(`指標「${spec.name}」の置く場所「${placeLabel(spec.target)}」が表に無い。出さなかった`);
        }
        if (waiting.length === pending.length) {
            // どれも置けない：置く場所が輪になっている
            warn(`指標の置く場所が輪になっている（${list(waiting.map((spec) => spec.name))}）。出さなかった`);
            break;
        }
        pending = waiting;
    }

    // ---- 比率の行の分子・分母を行に結ぶ。キーは区分（§sec:区分）か行のコード。見つからなければ名前でも探す。比率の行は指さない（輪になる）
    for (const { row, numerator, denominator } of pendingRatios) {
        const resolve = (key: string): string | null => {
            const isSection = key.startsWith(SECTION_KEY);
            // 警告に出す名前は書式ペインと同じ（区分は区分の名前、書式ペインの行は「計算の行 N」、指標はメジャーの名前、科目はコード）
            const label = isSection || key.startsWith(CALC_KEY) || key.startsWith(INDICATOR_KEY) ? placeLabel(key) : key;
            const found = (isSection ? sectionValue.get(key.slice(SECTION_KEY.length)) : undefined) ?? (rows.has(key) ? key : nameIndex.has(key) ? nameIndex.get(key) : undefined);
            if (found !== undefined && found !== null && rows.get(found)?.type === "calc") {
                warn(`比率の行「${row.name}」の「${label}」は比率の行なので、分子・分母にできない。空欄にした`);
                return null;
            }
            if (found === undefined) warn(`比率の行「${row.name}」の「${label}」が表に無い。空欄にした`);
            else if (found === null) warn(`比率の行「${row.name}」の「${label}」は同じ名前の行が 2 つ以上あり決められない`);
            return found ?? null;
        };
        const num = resolve(numerator);
        const den = resolve(denominator);
        if (num === null || den === null) continue;
        try {
            row.formula = parseFormula(`[${num}] / [${den}]`);
            row.refs = [num, den];
        } catch (error) {
            warn(`比率の行が組めない（${row.name}：${error instanceof FormulaError ? error.message : String(error)}）`);
        }
    }

    // ---- 段階の行の二重線は、続く段階の無い最後の合計だけ（損益計算書の当期純利益、貸借対照表の資産合計と負債純資産合計）。
    // 後の段階の行が自分の足す区分をすべて足していれば途中の段階
    const steps = display.filter(({ def }) => def.type === "step").map(({ def }) => def);
    steps.forEach((step, i) => {
        const later = steps.slice(i + 1);
        if (later.some((next) => step.summands.every((s) => next.summands.some((n) => n.code === s.code)))) step.continued = true;
    });

    // ---- 書式ペインに出すもの：計算の行と指標の選択肢
    const valueLabel = new Map(Array.from(sectionValue, ([label, code]) => [code, label] as const));
    const headLabel = new Map(Array.from(sectionHead, ([label, code]) => [code, label] as const));
    const nameCount = new Map<string, number>();
    for (const { def } of display) if (hasValue(def.type)) nameCount.set(def.name, (nameCount.get(def.name) ?? 0) + 1);
    /** 同じ名前の行に添える、行の出どころ（科目は科目コード、指標はメジャーの queryName） */
    const sourceOf = (code: string) =>
        code.startsWith(CALC_KEY)
            ? `計算の行 ${code.slice(CALC_KEY.length)}`
            : code.startsWith(INDICATOR_KEY)
              ? code.slice(INDICATOR_KEY.length)
              : groupSection.has(code)
                ? `${groupSection.get(code)}の中分類`
                : code;
    const choiceOf = (def: RowDef) => {
        const label = valueLabel.get(def.code) ?? headLabel.get(def.code);
        if (label !== undefined) return { value: SECTION_KEY + label, displayName: label };
        return { value: def.code, displayName: (nameCount.get(def.name) ?? 0) > 1 ? `${def.name}（${sourceOf(def.code)}）` : def.name };
    };
    const calcChoices: CalcChoices = {
        sections: emitted.map((sec) => ({ value: sec.label, displayName: sec.label })),
        refs: display.filter(({ def }) => hasValue(def.type) && def.type !== "calc").map(({ def }) => choiceOf(def)),
        places: display.filter(({ def }) => def.type !== "blank").map(({ def }) => choiceOf(def)),
    };

    return { rows, display, warnings, traits, calcChoices, attached, following };
}
