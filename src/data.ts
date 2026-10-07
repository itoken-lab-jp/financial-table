/**
 * DataView（matrix）から、科目・金額・イベント・指標を取り出す。
 *
 * 届き方（スター スキーマ）：行の段は組織（最大 3 段）→ イベント → 比較順 → 科目（区分 → 中分類 → 科目名を 1 つの欄に）→ 科目コード →
 * 並び → 残高フラグ・貸方フラグ。列は月（値のメジャーが 2 本以上なら、月の下にメジャーの段）。
 * 値は金額と指標。行の葉の values は列の葉の並びの番号がキーで、valueSourceIndex がどのメジャーか（省けば列の葉のメジャー）。
 * Power BI は値の無い組を落とすので、金額の葉はどれも金額を持つ実在の科目（小計・段階利益・率の行はビジュアルが組む）。金額が 1 つも無い
 * 葉（指標だけが科目ごとに届いた葉）の科目は受け取らない（1 つのファクトに単位の列を持つ形で、量の科目が空の行として出ないように）。同じ科目の定義が食い違うと別の葉で届く。最初の定義を使い、食い違えば知らせる。
 *
 * 指標は、いつも組織 × イベント × 月の小計を読む。科目とつながらないファクトの人数も、科目を外した値で正しく届く。小計は宣言の既定で切ってあり、指標があるときだけ
 * ビジュアルが科目の最初の段の小計を入れる設定を保存する。ほかの段の小計の節点は読まない。
 *
 * 続きの読み込みは 1 回分ずつ届く（fetchMoreData(false)）。1 回分ずつ readInput で読み、mergeInput で足し合わせる。続きの回は前の回と
 * 少し重なって届く（Microsoft の fetchMoreData の説明。table・categorical の lastMergeIndex にあたるものが matrix には無い）ので、
 * 行の葉と小計の節点を道筋（すべての段の値）で見分け、前の回までに読んだものは読まない（2026-09-27 の Desktop で、重なった葉を 2 度足していた）。
 * 科目の並びは、並びの欄があればその値、無ければ届いた順（モデルの「列で並べ替え」の順）。届いた順はセグメント・シナリオの組ごとにしか
 * 分からないので、組ごとの順を区分・中分類・科目の段ごとにつなぎ合わせる（assignOrder）。期間は月の値で並べる。
 */
import powerbi from "powerbi-visuals-api";

import { MonthIndex, toMonthDay } from "./periods";

import DataView = powerbi.DataView;
import DataViewHierarchyLevel = powerbi.DataViewHierarchyLevel;
import DataViewMatrixNode = powerbi.DataViewMatrixNode;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import PrimitiveValue = powerbi.PrimitiveValue;

export const ROLES = {
    /** 科目の階層 */
    account: "account",
    accountCode: "accountCode",
    /** 科目の並び（任意）。無ければ届いた順。セグメントで科目の顔ぶれが違うと届いた順では決めきれないので残した */
    accountOrder: "accountOrder",
    balanceFlag: "balanceFlag",
    creditFlag: "creditFlag",
    period: "period",
    event: "event",
    eventOrder: "eventOrder",
    organization: "organization",
    amount: "amount",
    indicator: "indicator",
} as const;

/** 科目の段より上の段（組織・イベント・比較順）。これより下の最初の段の小計が、組織 × イベント × 月の値 */
const ABOVE_ACCOUNT = new Set<string>([ROLES.organization, ROLES.event, ROLES.eventOrder]);

/** 月の欄を入れていないときの、すべての値を置く 1 つの月（表は「全期間」の 1 列。viewModel） */
export const ALL_PERIODS_MONTH: MonthIndex = 0;

/** 科目の欄も科目コードの欄も入れていないときの、値をすべて足す 1 つの科目のコードと名前 */
export const TOTAL_ACCOUNT = "§all";
export const TOTAL_ACCOUNT_NAME = "合計";

/** 指標の行のコードの頭（後ろはメジャーの queryName）。行のキーと、書式ペインの置く場所・分子・分母の保存値に使う */
export const INDICATOR_KEY = "§ind:";

/**
 * 区分マスタの 1 行（科目の行を通して届く。読んだまま）。区分の属性だけを持つ。段階の行・率の行は書式ペインの「計算行」
 */
export interface SectionRecord {
    order: number | null;
}

/** 科目マスタの 1 行（読んだまま。意味づけは rows.ts） */
export interface AccountRecord {
    code: string;
    name: string;
    order: number | null;
    category: string | null;
    subCategory: string | null;
    /**
     * 科目名より上の段の値（上から。空の段は null）。段の数は科目の欄に入れた列の数 − 1 で、0 なら科目名だけの表。
     * 無ければ category・subCategory から作る（rows.ts の groupsOf）
     */
    groups?: Array<string | null>;
    balanceFlag: string | null;
    /** 貸方フラグ：正常な残高が貸方の科目（収益・負債・純資産。評価勘定はその逆） */
    creditFlag: string | null;
    section: SectionRecord;
}

/** 比較順の型。数と日付が混ざったら知らせる */
export type OrderKind = "number" | "date";

export interface EventInfo {
    name: string;
    /**
     * 比較順。大きい（日付なら新しい）ほど確かで、最新見込みで先に採る。日付は時刻の数にする。
     * 空（null）のイベントは主に採らず、比較にだけ使う（前年実績など）
     */
    order: number | null;
    orderKind: OrderKind | null;
    /** 届いた順。横持ちはメジャーの欄の並び（比較順の推測と、同じ比較順の並びに使う）。縦持ちでは意味を持たせない */
    seen: number;
    /** 横持ちのメジャー：書式ペインのメジャーごとの設定の保存先（queryName）と、保存した比較順・金額の持ち方（生のまま） */
    measure?: { queryName: string; savedOrder: string | null };
    /**
     * 系列：イベントの列と金額のメジャー 2 本以上を一緒に入れたときの、イベントが組になるメジャー（名前と金額の欄の並び）。
     * 既定・比較の落とし先は同じ系列の中だけで見る
     */
    series?: { name: string; index: number };
}

/** 指標の欄のメジャー 1 本。行として表に足す（量・会社独自の行・うち） */
export interface IndicatorInfo {
    /** 行のコード（INDICATOR_KEY + queryName） */
    code: string;
    /** 書式ペインのメジャーごとの設定の保存先 */
    queryName: string;
    name: string;
    /** メジャーの書式（モデルの書式の文字。0.0% など）。率のメジャーかを見分ける */
    formatString: string;
    /** 書式ペインの保存値（生のまま。viewModel が読む） */
    saved: Record<string, string>;
}

/**
 * 金額・指標の値。データの符号のまま（符号を掛ける前）。
 * 日付の列を入れると 1 か月に何行も届くので、value は月の合計（フローの行）、last はその月の最後の日の値（期末の行）
 */
export interface Fact {
    code: string;
    month: MonthIndex;
    event: string;
    /** 組織（事業など）。段が 2 つ以上なら上の段からの道筋（ORG_SEPARATOR でつなぐ。空の段は ""）。入れていなければ "" */
    org: string;
    value: number;
    last: number;
    /** last の日（月だけの値なら 0） */
    lastDay: number;
}

export interface InputData {
    /** 科目の表。キーはコード。届いた順 */
    accounts: Map<string, AccountRecord>;
    facts: Fact[];
    events: EventInfo[];
    /** イベントの取り方。column = イベントの列、measures = 金額のメジャーの名前、single = メジャー 1 本だけ */
    eventSource: "column" | "measures" | "single";
    has: {
        code: boolean;
        name: boolean;
        category: boolean;
        period: boolean;
        amount: boolean;
        eventOrder: boolean;
        organization: boolean;
        creditFlag: boolean;
        balanceFlag: boolean;
        sectionMaster: boolean;
        /** 科目の欄に入れた列の数（区分 → 中分類 → 科目名） */
        accountLevels: number;
        /** 並びの欄を入れたか */
        order: boolean;
    };
    /** 取り込めなかった分。表は合計を出すので、落ちた行は合計が黙って小さくなる。indicatorNotNumber は指標のメジャーが数でない値を返した数 */
    dropped: { noAccount: number; unreadableMonth: number; unreadableMonthSamples: string[]; noEvent: number; notNumber: number; indicatorNotNumber: number };
    /** セグメント・シナリオの組ごとの、区分・中分類・科目の届いた順（組の道筋がキー。続きの回の同じ組はつなぐ） */
    orderGroups: Map<string, OrderGroup>;
    /** 並びの欄の値（コード → 並び）。読めた科目だけ */
    explicitOrders: Map<string, number>;
    /** 届いた順だけでは並びを決めきれなかった所（区分の名前。区分どうしの並びなら ORDER_OF_SECTIONS） */
    orderGuessed: string[];
    /** 同じコードで定義が食い違った科目 */
    conflicts: string[];
    /** 同じ名前で比較順が食い違ったイベント（最初の値を使う） */
    orderConflicts: string[];
    /** 比較順が数か日付として読めなかった値（イベント：値）。そのイベントは比較にだけ使う */
    unreadableOrders: string[];
    /** 金額に同じ名前で入れたメジャー（1 つのイベントに足される） */
    duplicateMeasures: string[];
    /** Power BI がまだ行を残している（metadata.segment） */
    truncated: boolean;
    /** 金額に入れたメジャーの数 */
    measureCount: number;
    /** 指標の欄のメジャー（欄の並び） */
    indicators: IndicatorInfo[];
    /** 指標の値（組織 × イベント × 月の小計）。横持ちはイベントが ""（書式ペインのメジャーごとのイベントで割り当てる。viewModel） */
    indicatorFacts: Fact[];
    /**
     * 指標の小計：科目の最初の段の欄（小計を入れる段。queryName）、その段の小計の節点が届いたか、小計を入れる設定が保存されているか
     * （保存が無ければ visual.ts が保存する）
     */
    indicatorLevel: { queryName: string | null; subtotals: boolean; enabled: boolean };
    /**
     * 指標の値が届いた所（指標の行のコード）：科目の葉（読まない）と、科目の最初の段の小計の節点。葉にだけ値があれば、メジャーが科目の段でだけ
     * 値を返している（ISINSCOPE の条件が逆など）ので知らせる
     */
    indicatorPresence: { leaf: Set<string>; subtotal: Set<string> };
    /** この回で新しく読んだ行の葉・小計の節点の道筋（続きの回の重なりを 2 度読まないため）。足し合わせると、それまでに読んだすべて */
    leafKeys: Set<string>;
    /**
     * 届いた DataView の中身の指紋（重なって読まなかった葉も含む、行の葉と小計の節点の道筋と値・行と列の段の欄・メジャーの queryName・
     * 続きの有無）。同じ回がもう一度届いた（大きさの変更・書式の変更・保存の応答）のを見分ける（loading.ts）
     */
    fingerprint: string;
    /** 金額・指標のメジャーの名前と書式の保存値の印（指紋に入れない。行と値が同じで、ここだけ変わったら当て直す。refreshMeasures） */
    measureKey: string;
    /** 金額のメジャーの名前（欄の並び。系列の名前の付け替えに使う） */
    measureNames: string[];
    /** 列（月の値）が上限（COLUMN_LIMIT）に届いた。列には続きの読み込みが無いので、それより後の期間が落ちているおそれ */
    columnsCut: boolean;
    /** 組織の段の名前（上の段から。組織を入れていなければ空） */
    orgLevelNames: string[];
    /** 組織の道筋（Fact.org）。届いた順（Power BI の並べ替えの順）。組織のブロックの並びに使う（orgs.ts） */
    orgs: string[];
    /** 選択 ID を作る節点（visual.ts が withMatrixNode で作る）。値を入れた科目・組織・月のもの */
    selection: SelectionNodes;
}

/**
 * 選択 ID の元になる matrix の節点。科目は科目コードの段（無ければ科目名の段）の節点、組織は道筋の頭（上の段から）ごとに、その段までの
 * 組織の段の節点、月は列の月の段の節点（日付の列なら 1 か月に何日も）。続きの読み込みの回のものも足し合わせる
 */
export interface SelectionNodes {
    rowLevels: DataViewHierarchyLevel[];
    columnLevels: DataViewHierarchyLevel[];
    accounts: Map<string, DataViewMatrixNode>;
    orgs: Map<string, DataViewMatrixNode[]>;
    months: Map<MonthIndex, DataViewMatrixNode[]>;
}

/** 1 つの組（セグメント・シナリオ）の、段ごとの届いた順。中分類は区分ごと、科目は区分・中分類ごと（キーは orderBucket） */
export interface OrderGroup {
    categories: string[];
    subCategories: Map<string, string[]>;
    accounts: Map<string, string[]>;
}

/** 区分どうしの並びを決めきれなかったときの印（orderGuessed） */
export const ORDER_OF_SECTIONS = "\u0000sections";

const orderBucket = (category: string | null, subCategory: string | null) => `${category ?? ""}\u0001${subCategory ?? ""}`;

const emptySelection = (): SelectionNodes => ({ rowLevels: [], columnLevels: [], accounts: new Map(), orgs: new Map(), months: new Map() });

/**
 * 列（月の値）の上限。capabilities.json の columns は top 2000 だが、行を window（続きを読む受け方）で受けると、
 * Power BI は列を 60 で切る（行の数によらない。列の小計も届かない）。行が多い表なので行の window を残し、60 に届いたら知らせる
 */
export const COLUMN_LIMIT = 60;

/**
 * 組織の段の値をつなぐ字。道筋を段に分け直すので、
 * 組織の名前に入らない制御文字にする（「東日本 › 北海道」のような名前で親の組織ができないように）
 */
export const ORG_SEPARATOR = "\u001f";

const EMPTY: InputData = {
    accounts: new Map(),
    facts: [],
    events: [],
    eventSource: "single",
    has: {
        code: false,
        name: false,
        category: false,
        period: false,
        amount: false,
        eventOrder: false,
        organization: false,
        creditFlag: false,
        balanceFlag: false,
        sectionMaster: false,
        accountLevels: 0,
        order: false,
    },
    dropped: { noAccount: 0, unreadableMonth: 0, unreadableMonthSamples: [], noEvent: 0, notNumber: 0, indicatorNotNumber: 0 },
    orderGroups: new Map(),
    explicitOrders: new Map(),
    orderGuessed: [],
    conflicts: [],
    orderConflicts: [],
    unreadableOrders: [],
    duplicateMeasures: [],
    truncated: false,
    measureCount: 0,
    indicators: [],
    indicatorFacts: [],
    indicatorLevel: { queryName: null, subtotals: false, enabled: false },
    indicatorPresence: { leaf: new Set(), subtotal: new Set() },
    leafKeys: new Set(),
    fingerprint: "",
    measureKey: "",
    measureNames: [],
    columnsCut: false,
    orgLevelNames: [],
    orgs: [],
    selection: emptySelection(),
};

/** 空・空白は null。数は String() の字にする（コード 4000 と "4000" を同じにする） */
export function text(value: PrimitiveValue | undefined): string | null {
    if (value === null || value === undefined) return null;
    if (value instanceof Date) return value.toISOString();
    const s = String(value).trim();
    return s === "" ? null : s;
}

const isBlank = (value: PrimitiveValue | undefined) => value === null || value === undefined || value === "";

function num(value: PrimitiveValue | undefined): number | null {
    if (value === null || value === undefined || value === "") return null;
    const n = typeof value === "number" ? value : Number(value);
    return Number.isFinite(n) ? n : null;
}

/**
 * 比較順の値：数・数の字はそのまま。日付・日付の字（ISO の日時・2025-10-01・2025/10/1・2025年10月1日）は時刻の数
 * （ファイルの更新日などで比べられる）。読めなければ null
 */
export function parseOrder(value: PrimitiveValue | undefined): { value: number; kind: OrderKind } | null {
    if (value === null || value === undefined || value === "" || typeof value === "boolean") return null;
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : { value: value.getTime(), kind: "date" };
    if (typeof value === "number") return Number.isFinite(value) ? { value, kind: "number" } : null;
    const s = String(value).trim();
    if (s === "") return null;
    const n = Number(s);
    if (Number.isFinite(n)) return { value: n, kind: "number" };
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s)) {
        const date = new Date(s);
        if (!Number.isNaN(date.getTime())) return { value: date.getTime(), kind: "date" };
    }
    const match = /^(\d{4})\s*[-/.年]\s*(\d{1,2})(?:\s*[-/.月]\s*(\d{1,2})日?)?$/.exec(s);
    if (!match) return null;
    const [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3] ?? 1)];
    const date = new Date(Date.UTC(year, month - 1, day));
    // 2025-02-30 のような無い日は、Date が次の月に送るので読まない
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
    return { value: date.getTime(), kind: "date" };
}

/** 欄（ロール）の名前。ROLES の値のどれか */
const ROLE_NAMES = new Set<string>(Object.values(ROLES));
const roleOf = (source: DataViewMetadataColumn | undefined): string | null =>
    source ? (Object.keys(source.roles ?? {}).find((role) => ROLE_NAMES.has(role)) ?? null) : null;
/** 節点の値（levelValues が新しい形。古い value も読む） */
const nodeValue = (node: DataViewMatrixNode): PrimitiveValue => node.levelValues?.[0]?.value ?? node.value ?? null;
/** 節点が段のどの欄の値か（複合の段は無い宣言なので最初の値の欄） */
const nodeSourceIndex = (node: DataViewMatrixNode): number => node.levelValues?.[0]?.levelSourceIndex ?? node.levelSourceIndex ?? 0;

/** 書式ペインの保存値（object の中の字の値だけ）。指標の欄のメジャーごとの設定 */
function savedTexts(object: powerbi.DataViewObject | undefined): Record<string, string> {
    const result: Record<string, string> = {};
    for (const [key, value] of Object.entries(object ?? {})) {
        const text = savedString(value);
        if (text !== null) result[key] = text;
    }
    return result;
}

/**
 * 書式ペインの保存値を文字で読む。ドロップダウンの値が数字の形（「3」・科目のコード「4110」など）だと、Desktop は数で保存し
 * （visual.json に 3D）、数で届く（2026-09-28 に Desktop で確かめた。文字しか受けないと読めずに既定に戻った）。数は文字に戻す
 * （先頭の 0 は戻らない）
 */
export function savedString(value: unknown): string | null {
    if (typeof value === "string") return value !== "" ? value : null;
    if (typeof value === "number" && Number.isFinite(value)) return String(value);
    return null;
}

/** 列の葉：月の値と、どのメジャーの列か（メジャーが 1 本なら 0）。小計の列の葉は値を読まない */
interface ColumnLeaf {
    period: PrimitiveValue | undefined;
    measure: number;
    subtotal: boolean;
}

/** 字の指紋（FNV-1a、32 ビット）を続けて混ぜる */
function mix(hash: number, text: string): number {
    let h = hash;
    for (let i = 0; i < text.length; i++) {
        h ^= text.charCodeAt(i);
        h = Math.imul(h, 16777619) >>> 0;
    }
    return h;
}

/** 道筋の値を字にする（行の葉・小計の節点のキー） */
const pathText = (values: PrimitiveValue[]) =>
    values.map((value) => (value instanceof Date ? `d${value.getTime()}` : value === null || value === undefined ? "" : `${typeof value}:${String(value)}`)).join("\u0001");

/**
 * 1 回分を読む。seen は前の回までに読んだ行の葉・小計の節点の道筋（続きの回で重なって届いたものは読まない）
 */
export function readInput(dataView: DataView | undefined, seen?: ReadonlySet<string>): InputData {
    const matrix = dataView?.matrix;
    if (!matrix) return { ...EMPTY, leafKeys: new Set(), indicatorPresence: { leaf: new Set(), subtotal: new Set() }, orgs: [], selection: emptySelection() };
    const rowLevels = matrix.rows?.levels ?? [];
    const columnLevels = matrix.columns?.levels ?? [];
    const valueSources = matrix.valueSources ?? [];

    // 行の段：どの段がどの欄か。組織は最大 3 段
    const levelRoles = rowLevels.map((level) => roleOf(level.sources?.[0]));
    const levelOf = (role: string) => levelRoles.indexOf(role);
    const orgLevels = levelRoles.map((role, i) => (role === ROLES.organization ? i : -1)).filter((i) => i >= 0);
    const codeLevel = levelOf(ROLES.accountCode);
    // 科目の欄は 1 つに階層を上から入れる：一番上が区分、一番下が科目名、あいだが中分類（それまでは区分・中分類・科目名・並びが別の欄）
    const accountLevels = levelRoles.map((role, i) => (role === ROLES.account ? i : -1)).filter((i) => i >= 0);
    const nameLevel = accountLevels.at(-1) ?? -1;
    const orderLevel = levelOf(ROLES.accountOrder);
    const eventLevel = levelOf(ROLES.event);
    const eventOrderLevel = levelOf(ROLES.eventOrder);
    const attributeLevels = {
        category: accountLevels.length >= 2 ? accountLevels[0] : -1,
        subCategory: accountLevels.length >= 3 ? accountLevels[1] : -1,
        balanceFlag: levelOf(ROLES.balanceFlag),
        creditFlag: levelOf(ROLES.creditFlag),
    };
    // 科目の最初の段（組織・イベント・比較順より下の最初の段）。この段の小計の節点が、組織 × イベント × 月の値
    const accountLevel = levelRoles.findIndex((role) => role !== null && !ABOVE_ACCOUNT.has(role));

    // 値：金額の欄と指標の欄のメジャー（valueSources の番号 → 欄の中の番号）
    const indexIn = (role: string) => {
        let count = 0;
        return valueSources.map((source) => (source.roles?.[role] ? count++ : -1));
    };
    const amountIndex = indexIn(ROLES.amount);
    const indicatorIndex = indexIn(ROLES.indicator);
    const amountSources = valueSources.filter((source) => source.roles?.[ROLES.amount]);
    const indicatorSources = valueSources.filter((source) => source.roles?.[ROLES.indicator]);
    const hasPeriod = columnLevels.some((level) => level.sources?.some((source) => source.roles?.[ROLES.period]));

    // 列の葉を並びの順に集める（行の葉の values のキーはこの並びの番号）。列の小計は宣言で切っているが、届けば番号には数え
    // （飛ばすと、小計より後ろの列の値が 1 つずつずれる）、値は読まない
    const columnLeaves: ColumnLeaf[] = [];
    const selection: SelectionNodes = { rowLevels, columnLevels, accounts: new Map(), orgs: new Map(), months: new Map() };
    const walkColumns = (node: DataViewMatrixNode, period: PrimitiveValue | undefined, measure: number, subtotal: boolean): void => {
        const children = node.children ?? [];
        if (children.length === 0) {
            columnLeaves.push({ period, measure, subtotal });
            return;
        }
        for (const child of children) {
            const inSubtotal = subtotal || !!child.isSubtotal;
            const source = columnLevels[child.level ?? 0]?.sources?.[nodeSourceIndex(child)];
            if (source?.roles?.[ROLES.period] && !child.isSubtotal) {
                // 月の列の節点を覚える（列の見出し・セルを押したときの選択 ID）
                const monthDay = toMonthDay(nodeValue(child));
                if (monthDay !== null && !inSubtotal) {
                    const list = selection.months.get(monthDay.month) ?? [];
                    if (!list.includes(child)) selection.months.set(monthDay.month, [...list, child]);
                }
                walkColumns(child, nodeValue(child), measure, inSubtotal);
            }
            else if (source?.isMeasure) {
                // メジャーの段：どのメジャーかは queryName で valueSources の番号に直す（無ければ段の中の番号）
                const k = valueSources.findIndex((v) => v.queryName === source.queryName);
                walkColumns(child, period, k >= 0 ? k : nodeSourceIndex(child), inSubtotal);
            } else walkColumns(child, period, measure, inSubtotal);
        }
    };
    if (matrix.columns?.root) walkColumns(matrix.columns.root, undefined, 0, false);
    const periods = new Set(columnLeaves.filter((leaf) => !leaf.subtotal).map((leaf) => text(leaf.period ?? null) ?? ""));
    /** 列の葉の月。月の欄が無ければ、すべて 1 つの月（全期間） */
    const monthOfLeaf = (leaf: ColumnLeaf | undefined) => (hasPeriod ? toMonthDay(leaf?.period ?? null) : { month: ALL_PERIODS_MONTH, day: 0 });

    const eventSource: InputData["eventSource"] = eventLevel >= 0 ? "column" : amountSources.length > 1 ? "measures" : "single";
    const measureName = (a: number) => amountSources[a].displayName ?? `値${a + 1}`;
    const indicators: IndicatorInfo[] = indicatorSources.map((source, i) => {
        const queryName = source.queryName ?? source.displayName ?? `指標${i + 1}`;
        return {
            code: INDICATOR_KEY + queryName,
            queryName,
            name: source.displayName ?? `指標${i + 1}`,
            formatString: source.format ?? "",
            saved: savedTexts(source.objects?.indicators),
        };
    });

    const accounts = new Map<string, AccountRecord>();
    const conflicts = new Set<string>();
    const events = new Map<string, EventInfo>();
    const orderConflicts = new Set<string>();
    const unreadableOrders = new Set<string>();
    const factMap = new Map<string, Fact>();
    const indicatorMap = new Map<string, Fact>();
    const dropped = { noAccount: 0, unreadableMonth: 0, unreadableMonthSamples: [] as string[], noEvent: 0, notNumber: 0, indicatorNotNumber: 0 };
    const presence = { leaf: new Set<string>(), subtotal: new Set<string>() };

    // 比較順は最初に読めた値を使う。イベントの表の列なので同じ名前なら同じ値のはず。食い違えば知らせる
    const noteEvent = (name: string, rawOrder: PrimitiveValue | undefined, series: EventInfo["series"]) => {
        const order = parseOrder(rawOrder);
        // 空白だけは空と同じ（読めない値として知らせない）
        if (order === null && text(rawOrder) !== null) unreadableOrders.add(`${name}：${text(rawOrder)}`);
        const known = events.get(name);
        if (!known) events.set(name, { name, order: order?.value ?? null, orderKind: order?.kind ?? null, seen: events.size, ...(series ? { series } : {}) });
        else if (order !== null) {
            if (known.order === null) {
                known.order = order.value;
                known.orderKind = order.kind;
            } else if (known.order !== order.value) orderConflicts.add(name);
        }
    };
    // 横持ち：メジャーがイベント。比較順は書式ペインのメジャーごとの値（無ければ欄に入れた順。events.ts）
    if (eventSource !== "column") {
        const savedText = savedString;
        amountSources.forEach((source, a) => {
            const saved = source.objects?.events;
            events.set(measureName(a), {
                name: measureName(a),
                order: null,
                orderKind: null,
                seen: a,
                measure: { queryName: source.queryName ?? measureName(a), savedOrder: savedText(saved?.order) },
            });
        });
    }

    // 行の木を歩く。道筋（段ごとの値）は使い回す配列に持つ（葉は一番下の段なので、葉までのどの段も書き直される）
    const path: PrimitiveValue[] = rowLevels.map((): PrimitiveValue => null);
    // 道筋の段ごとの節点（選択 ID を作るため。path と同じく使い回す）
    const pathNodes: Array<DataViewMatrixNode | null> = rowLevels.map((): DataViewMatrixNode | null => null);
    /** 値を入れた科目・組織の節点を覚える（選択 ID）。code が null なら組織だけ（指標の小計の節点） */
    const noteSelection = (code: string | null) => {
        // キーに使った段の節点（科目コードが空で科目名に落ちた科目は科目名の段。空のコードの節点だと「コードが空」で絞った）
        const accountNode = pathNodes[codeLevel >= 0 && at(codeLevel) !== null ? codeLevel : nameLevel];
        if (code !== null && accountNode && !selection.accounts.has(code)) selection.accounts.set(code, accountNode);
        for (let i = 0; i < orgLevels.length; i++) {
            const key = orgLevels.slice(0, i + 1).map((level) => text(path[level]) ?? "").join(ORG_SEPARATOR);
            if (selection.orgs.has(key)) continue;
            const nodes = orgLevels.slice(0, i + 1).map((level) => pathNodes[level]);
            if (nodes.every((n): n is DataViewMatrixNode => n !== null)) selection.orgs.set(key, nodes);
        }
    };
    const at = (level: number): string | null => (level >= 0 ? text(path[level]) : null);
    const leafKeys = new Set<string>();
    // 指紋：行の葉・小計の節点の道筋と値（重なりで読まないものも入れる。届いた DataView そのものの指紋）
    // 段の欄は、列（queryName）とどの欄（roles）に入れたか。同じ列を別の欄へ移しただけでも、読み直す
    const sourceKey = (source: DataViewMetadataColumn) => `${source.queryName ?? ""}#${Object.keys(source.roles ?? {}).sort().join("+")}`;
    const levelsKey = [...rowLevels, ...columnLevels].map((level) => (level.sources ?? []).map(sourceKey).join(",")).join("/");
    let hashA = mix(mix(0x811c9dc5, levelsKey), columnLeaves.map((leaf) => `${text(leaf.period ?? null)}|${leaf.measure}|${leaf.subtotal}`).join("\u0002"));
    let hashB = mix(0x01000193, valueSources.map(sourceKey).join("\u0002"));
    let leafCount = 0;
    let subtotalCount = 0;

    /** 月が読めない値を数える（見本の字を 3 つまで残す） */
    const unreadableMonth = (period: PrimitiveValue | undefined) => {
        dropped.unreadableMonth++;
        const sample = text(period ?? null) ?? "（空）";
        if (dropped.unreadableMonthSamples.length < 3 && !dropped.unreadableMonthSamples.includes(sample)) dropped.unreadableMonthSamples.push(sample);
    };
    // 組織の道筋。値を入れた組織を届いた順に覚える（組織のブロックの並び。行が全部落ちた組織は並べない）
    const orgSeen = new Set<string>();
    const orgOfPath = () => orgLevels.map((level) => text(path[level]) ?? "").join(ORG_SEPARATOR);
    // セグメント・シナリオの組ごとの、区分・中分類・科目の届いた順と、並びの欄の値
    const orderGroups = new Map<string, OrderGroup>();
    const explicitOrders = new Map<string, number>();
    /** 科目のキー：科目コード。無ければ区分・中分類・科目名（同じ科目名が別の区分にあっても 1 行に足さない） */
    const keyOfPath = (): string | null => {
        // 科目の欄も科目コードの欄も無い：値をすべて 1 つの科目に足す（表は「合計」の 1 行）
        if (codeLevel < 0 && nameLevel < 0) return TOTAL_ACCOUNT;
        const code = at(codeLevel);
        if (code !== null) return code;
        const name = at(nameLevel);
        if (name === null) return null;
        return [...accountLevels.slice(0, -1).map(at), name].filter((v): v is string => v !== null).join(" / ");
    };
    /** 届いた順を覚える（区分・中分類・科目の段ごと）。続きの回で重なって届いた葉も覚える（覚えないと、後の回の新しい科目に前後の足場が無い） */
    const noteOrder = (code: string) => {
        const groupKey = pathText(path.slice(0, accountLevel));
        let group = orderGroups.get(groupKey);
        if (!group) orderGroups.set(groupKey, (group = { categories: [], subCategories: new Map(), accounts: new Map() }));
        const category = at(attributeLevels.category);
        const subCategory = at(attributeLevels.subCategory);
        const push = (list: string[], value: string) => {
            if (list.at(-1) !== value && !list.includes(value)) list.push(value);
        };
        push(group.categories, category ?? "");
        const subs = group.subCategories.get(category ?? "") ?? [];
        group.subCategories.set(category ?? "", subs);
        push(subs, subCategory ?? "");
        const bucket = orderBucket(category, subCategory);
        const codes = group.accounts.get(bucket) ?? [];
        group.accounts.set(bucket, codes);
        push(codes, code);
    };

    const readLeaf = (node: DataViewMatrixNode): void => {
        // 行の葉は道筋で 1 つに決まる（同じ道筋の葉は、続きの回の重なり）
        const leafKey = pathText(path);
        const cells = node.values ?? {};
        const keys = Object.keys(cells);
        leafCount++;
        hashA = mix(hashA, leafKey);
        for (const key of keys) hashB = mix(hashB, `${key}=${String(cells[Number(key)].value)}`);
        if (seen?.has(leafKey) || leafKeys.has(leafKey)) {
            // 重なって届いた葉：金額は読まないが、並びの足場には使う
            const amount = keys.some((key) => {
                const cell = cells[Number(key)];
                if (columnLeaves[Number(key)]?.subtotal) return false;
                const k = cell.valueSourceIndex ?? columnLeaves[Number(key)]?.measure ?? 0;
                return (amountIndex[k] ?? -1) >= 0 && num(cell.value) !== null;
            });
            const code = amount ? keyOfPath() : null;
            if (code !== null) noteOrder(code);
            return;
        }
        leafKeys.add(leafKey);
        let hasAmount = false;
        for (const key of keys) {
            const cell = cells[Number(key)];
            if (columnLeaves[Number(key)]?.subtotal) continue;
            const k = cell.valueSourceIndex ?? columnLeaves[Number(key)]?.measure ?? 0;
            // 葉の指標の値は読まない（組織 × イベント × 月の小計で読む）。届いたことだけ覚える
            const i = indicatorIndex[k] ?? -1;
            if (i >= 0 && num(cell.value) !== null) presence.leaf.add(indicators[i].code);
            if ((amountIndex[k] ?? -1) < 0) continue;
            // 金額のメジャーが文字を返す（FORMAT で整えたなど）と数にならない。黙って捨てずに数える
            if (num(cell.value) !== null) hasAmount = true;
            else if (!isBlank(cell.value)) dropped.notNumber++;
        }
        // 組織・イベントより下の段が無い表は小計の節点が無いので、指標の値を葉で読む（葉が組織 × イベント × 月の値）
        if (accountLevel < 0 && indicators.length > 0) readSubtotal(node);
        // 金額が 1 つも無い葉（指標だけが科目ごとに届いた葉）の科目は受け取らない。量の科目が空の行として出ないように
        if (!hasAmount) return;
        const code = keyOfPath();
        if (code === null) {
            dropped.noAccount++;
            return;
        }
        const record: AccountRecord = {
            code,
            name: at(nameLevel) ?? (code === TOTAL_ACCOUNT ? TOTAL_ACCOUNT_NAME : code),
            order: null,
            category: at(attributeLevels.category),
            subCategory: at(attributeLevels.subCategory),
            groups: accountLevels.slice(0, -1).map(at),
            balanceFlag: at(attributeLevels.balanceFlag),
            creditFlag: at(attributeLevels.creditFlag),
            section: { order: null },
        };
        const known = accounts.get(code);
        if (!known) accounts.set(code, record);
        else if (!sameDefinition(known, record)) conflicts.add(code);
        // 並び：並びの欄の値（食い違えば定義の食い違いとして知らせる）と、組ごとの届いた順（Power BI は区分・中分類・科目名を「列で並べ替え」の順で届ける）
        const explicit = orderLevel >= 0 ? num(path[orderLevel]) : null;
        if (explicit !== null) {
            const knownOrder = explicitOrders.get(code);
            if (knownOrder === undefined) explicitOrders.set(code, explicit);
            else if (knownOrder !== explicit) conflicts.add(code);
        }
        noteOrder(code);

        const org = orgOfPath();
        for (const key of keys) {
            const cell = cells[Number(key)];
            const leaf = columnLeaves[Number(key)];
            if (leaf?.subtotal) continue;
            const k = cell.valueSourceIndex ?? leaf?.measure ?? 0;
            const a = amountIndex[k] ?? -1;
            const value = num(cell.value);
            if (a < 0 || value === null) continue;
            const monthDay = monthOfLeaf(leaf);
            if (monthDay === null) {
                unreadableMonth(leaf?.period);
                continue;
            }
            let event: string;
            if (eventSource === "column") {
                const name = at(eventLevel);
                if (name === null) {
                    dropped.noEvent++;
                    continue;
                }
                const series = amountSources.length > 1 ? { name: measureName(a), index: a } : undefined;
                event = series ? `${name}・${series.name}` : name;
                noteEvent(event, eventOrderLevel >= 0 ? path[eventOrderLevel] : undefined, series);
            } else {
                event = measureName(a);
            }
            addFact(factMap, { code, month: monthDay.month, event, org, value, last: value, lastDay: monthDay.day });
            if ((orgLevels.length > 0 && !selection.orgs.has(org)) || !selection.accounts.has(code)) noteSelection(code);
            orgSeen.add(org);
        }
    };

    /**
     * 科目の最初の段の小計の節点：組織 × イベント × 月の値。指標だけを読む（金額は科目の葉を足す）。
     * 縦持ちのイベントは節点の上の段。系列（イベントの列と金額のメジャー 2 本以上）があれば、最初の系列の組に入れる（主の既定の系列。
     * 系列ごとに入れると、金額の無い系列のイベントまで増える）。横持ちは "" にして、書式ペインのメジャーごとのイベントで割り当てる（viewModel）。
     * イベントは金額の葉で並べる（指標だけのイベントは並べない。並べると比較の既定・候補や年度の選択肢に混ざり、金額の比較が空欄になる。
     * ）
     */
    const readSubtotal = (node: DataViewMatrixNode): void => {
        const subtotalKey = `${pathText(accountLevel < 0 ? path : path.slice(0, accountLevel))}\u0001§sub`;
        const cells = node.values ?? {};
        const keys = Object.keys(cells);
        subtotalCount++;
        hashA = mix(hashA, subtotalKey);
        for (const key of keys) hashB = mix(hashB, `s${key}=${String(cells[Number(key)].value)}`);
        if (seen?.has(subtotalKey) || leafKeys.has(subtotalKey)) return;
        leafKeys.add(subtotalKey);
        const values: Array<{ indicator: IndicatorInfo; month: MonthIndex; day: number; value: number }> = [];
        for (const key of keys) {
            const cell = cells[Number(key)];
            const leaf = columnLeaves[Number(key)];
            if (leaf?.subtotal) continue;
            const i = indicatorIndex[cell.valueSourceIndex ?? leaf?.measure ?? 0] ?? -1;
            if (i < 0) continue;
            const value = num(cell.value);
            if (value === null) {
                if (!isBlank(cell.value)) dropped.indicatorNotNumber++;
                continue;
            }
            const monthDay = monthOfLeaf(leaf);
            if (monthDay === null) {
                unreadableMonth(leaf?.period);
                continue;
            }
            values.push({ indicator: indicators[i], month: monthDay.month, day: monthDay.day, value });
            presence.subtotal.add(indicators[i].code);
        }
        if (values.length === 0) return;
        let event: string;
        if (eventSource === "column") {
            const name = at(eventLevel);
            if (name === null) {
                dropped.noEvent++;
                return;
            }
            event = amountSources.length > 1 ? `${name}・${measureName(0)}` : name;
        } else event = eventSource === "single" && amountSources.length === 1 ? measureName(0) : "";
        const org = orgOfPath();
        for (const v of values) {
            addFact(indicatorMap, { code: v.indicator.code, month: v.month, event, org, value: v.value, last: v.value, lastDay: v.day });
            if (orgLevels.length > 0 && !selection.orgs.has(org)) noteSelection(null);
            orgSeen.add(org);
        }
    };

    let subtotalsSeen = false;
    const walkRows = (node: DataViewMatrixNode): void => {
        for (const child of node.children ?? []) {
            const level = child.level ?? 0;
            if (child.isSubtotal) {
                // 科目の最初の段の小計だけを読む（ほかの段の小計は、宣言で切るか、読まない）
                if (level === accountLevel && accountLevel >= 0) {
                    subtotalsSeen = true;
                    if (indicators.length > 0) readSubtotal(child);
                }
                continue;
            }
            if (level < path.length) {
                path[level] = nodeValue(child);
                pathNodes[level] = child;
            }
            if (child.children && child.children.length > 0) walkRows(child);
            else readLeaf(child);
        }
    };
    const root = matrix.rows?.root;
    // 行の段が 1 つも無い（値だけ・値と列だけ）ときは、根が葉になって値を持つ
    if (root?.children && root.children.length > 0) walkRows(root);
    else if (root?.values) readLeaf(root);

    // 小計を入れる設定：全体の行の小計・段ごとの小計と、科目の最初の段の小計（段の欄の selector で保存したもの）
    const orderGuessed = assignOrder(accounts, orderGroups, explicitOrders);
    const accountSource = accountLevel >= 0 ? rowLevels[accountLevel]?.sources?.[0] : undefined;
    const subTotals = dataView?.metadata?.objects?.subTotals;
    const levelEnabled = accountSource?.objects?.subTotals?.levelSubtotalEnabled === true;

    return {
        accounts,
        facts: Array.from(factMap.values()),
        events: Array.from(events.values()),
        eventSource,
        has: {
            code: codeLevel >= 0,
            name: nameLevel >= 0,
            category: attributeLevels.category >= 0,
            accountLevels: accountLevels.length,
            period: hasPeriod,
            amount: amountSources.length > 0,
            eventOrder: eventOrderLevel >= 0,
            organization: orgLevels.length > 0,
            creditFlag: attributeLevels.creditFlag >= 0,
            balanceFlag: attributeLevels.balanceFlag >= 0,
            // 区分の並びは、区分の最初の科目の並び（区分の並びの欄はやめた）
            sectionMaster: false,
            order: orderLevel >= 0,
        },
        dropped,
        orderGroups,
        explicitOrders,
        orderGuessed,
        conflicts: Array.from(conflicts),
        orderConflicts: Array.from(orderConflicts),
        unreadableOrders: Array.from(unreadableOrders),
        // 同じ名前のメジャーは、イベント（横持ち）や系列の名前が重なり、金額が 1 つに足される
        duplicateMeasures: Array.from(new Set(amountSources.map((_, a) => measureName(a)).filter((name, a, names) => names.indexOf(name) !== a))),
        truncated: !!dataView?.metadata?.segment,
        measureCount: amountSources.length,
        indicators,
        indicatorFacts: Array.from(indicatorMap.values()),
        indicatorLevel: {
            queryName: accountSource?.queryName ?? null,
            subtotals: subtotalsSeen || accountLevel < 0,
            enabled: subTotals?.rowSubtotals === true && subTotals?.perRowLevel === true && levelEnabled,
        },
        indicatorPresence: presence,
        leafKeys,
        fingerprint: `${leafCount}|${subtotalCount}|${hashA.toString(16)}|${hashB.toString(16)}|${dataView?.metadata?.segment ? "more" : "end"}`,
        measureKey: JSON.stringify([
            amountSources.map((source) => [source.queryName ?? "", source.displayName ?? "", source.objects?.events ?? null]),
            indicatorSources.map((source) => [source.queryName ?? "", source.displayName ?? "", source.objects?.indicators ?? null]),
        ]),
        measureNames: amountSources.map((_, a) => measureName(a)),
        columnsCut: periods.size >= COLUMN_LIMIT,
        orgLevelNames: orgLevels.map((level, i) => rowLevels[level].sources?.[0]?.displayName ?? `セグメント${i + 1}`),
        orgs: orgLevels.length > 0 ? Array.from(orgSeen) : [],
        selection,
    };
}

/** 続きの回の選択の節点を足す（先に届いたものを使う。段の欄は新しい回のもの） */
function mergeSelection(base: SelectionNodes, next: SelectionNodes): SelectionNodes {
    const accounts = new Map(base.accounts);
    for (const [k, v] of next.accounts) if (!accounts.has(k)) accounts.set(k, v);
    const orgs = new Map(base.orgs);
    for (const [k, v] of next.orgs) if (!orgs.has(k)) orgs.set(k, v);
    const months = new Map(base.months);
    for (const [k, v] of next.months) if (!months.has(k)) months.set(k, v);
    return { rowLevels: next.rowLevels, columnLevels: next.columnLevels, accounts, orgs, months };
}

const factKey = (fact: Fact) => `${fact.code}\u0000${fact.month}\u0000${fact.event}\u0000${fact.org}`;

/** 同じ組の金額を足す。月ごとに足し（フロー）、最後の日の値も残す（期末） */
function combine(fact: Fact, next: Fact): void {
    fact.value += next.value;
    if (next.lastDay > fact.lastDay) {
        fact.last = next.last;
        fact.lastDay = next.lastDay;
    } else if (next.lastDay === fact.lastDay) fact.last += next.last;
}

/**
 * 金額を足す。日付の列を入れたときは 1 か月に何日分も届くので、月ごとに足し（フロー）、最後の日の値も残す（期末）。
 * 定義の食い違う科目が別の葉になったときも同じ決まりで足す
 */
function addFact(factMap: Map<string, Fact>, next: Fact): void {
    const key = factKey(next);
    const fact = factMap.get(key);
    if (!fact) factMap.set(key, { ...next });
    else combine(fact, next);
}

/** 足し合わせた受け取りの金額・指標の値の索引（続きの読み込みで回を重ねても、足すのを新しい回の分だけにする） */
const factIndexes = new WeakMap<InputData, { facts: Map<string, Fact>; indicators: Map<string, Fact> }>();

/** 値の並び（base の配列）に、新しい回の値を足し合わせる。同じ組は足し、無い組は写して足す */
function mergeFacts(target: Fact[], index: Map<string, Fact>, next: Fact[]): void {
    for (const fact of next) {
        const key = factKey(fact);
        const known = index.get(key);
        if (known) combine(known, fact);
        else {
            const copy = { ...fact };
            index.set(key, copy);
            target.push(copy);
        }
    }
}

const indexOf = (facts: Fact[]) => new Map(facts.map((fact): [string, Fact] => [factKey(fact), fact]));

/**
 * 続きの読み込みの 1 回分（next）を、それまでの分（base）に足し合わせる。欄の有無・メジャーは同じ宣言なので base のまま、
 * 行の上限に届いたか（truncated）・指紋・列の上限は新しい回のもの。
 * base は足し合わせに使い回す（科目・金額・葉の道筋を書き足す）。呼んだあとは返した方を使い、base は使わない
 * （回ごとに全体を作り直すと、回の数 × 全体の手間になる）
 */
export function mergeInput(base: InputData, next: InputData): InputData {
    const accounts = base.accounts;
    const conflicts = new Set([...base.conflicts, ...next.conflicts]);
    for (const [code, record] of next.accounts) {
        const known = accounts.get(code);
        if (!known) accounts.set(code, record);
        else if (!sameDefinition(known, record)) conflicts.add(code);
    }
    // 並びは、組ごとの届いた順（同じ組は続きをつなぐ）と並びの欄の値を足し合わせて振り直す
    const orderGroups = base.orderGroups;
    for (const [key, group] of next.orderGroups) {
        const known = orderGroups.get(key);
        if (!known) {
            orderGroups.set(key, group);
            continue;
        }
        // 同じ組の続き：前の回の届いた順の後ろに足す（重なって届いた葉は、もうあるので足さない）
        const append = (list: string[], more: string[]) => [...list, ...more.filter((v) => !list.includes(v))];
        known.categories = append(known.categories, group.categories);
        for (const [k, list] of group.subCategories) known.subCategories.set(k, append(known.subCategories.get(k) ?? [], list));
        for (const [k, list] of group.accounts) known.accounts.set(k, append(known.accounts.get(k) ?? [], list));
    }
    const explicitOrders = base.explicitOrders;
    for (const [code, order] of next.explicitOrders) {
        const known = explicitOrders.get(code);
        if (known === undefined) explicitOrders.set(code, order);
        else if (known !== order) conflicts.add(code);
    }
    const orderGuessed = assignOrder(accounts, orderGroups, explicitOrders);
    const index = factIndexes.get(base) ?? { facts: indexOf(base.facts), indicators: indexOf(base.indicatorFacts) };
    mergeFacts(base.facts, index.facts, next.facts);
    mergeFacts(base.indicatorFacts, index.indicators, next.indicatorFacts);
    for (const key of next.leafKeys) base.leafKeys.add(key);
    const events = new Map<string, EventInfo>(base.events.map((e): [string, EventInfo] => [e.name, { ...e }]));
    const orderConflicts = new Set([...base.orderConflicts, ...next.orderConflicts]);
    for (const event of next.events) {
        const known = events.get(event.name);
        if (!known) events.set(event.name, { ...event, seen: events.size });
        else if (event.order !== null) {
            if (known.order === null) {
                known.order = event.order;
                known.orderKind = event.orderKind;
            } else if (known.order !== event.order) orderConflicts.add(event.name);
        }
    }
    const samples = [...base.dropped.unreadableMonthSamples];
    for (const sample of next.dropped.unreadableMonthSamples) if (samples.length < 3 && !samples.includes(sample)) samples.push(sample);
    const merged: InputData = {
        ...base,
        accounts,
        facts: base.facts,
        indicatorFacts: base.indicatorFacts,
        events: Array.from(events.values()),
        dropped: {
            noAccount: base.dropped.noAccount + next.dropped.noAccount,
            unreadableMonth: base.dropped.unreadableMonth + next.dropped.unreadableMonth,
            unreadableMonthSamples: samples,
            noEvent: base.dropped.noEvent + next.dropped.noEvent,
            notNumber: base.dropped.notNumber + next.dropped.notNumber,
            indicatorNotNumber: base.dropped.indicatorNotNumber + next.dropped.indicatorNotNumber,
        },
        orderGroups,
        explicitOrders,
        orderGuessed,
        conflicts: Array.from(conflicts),
        orderConflicts: Array.from(orderConflicts),
        unreadableOrders: Array.from(new Set([...base.unreadableOrders, ...next.unreadableOrders])),
        truncated: next.truncated,
        indicatorLevel: { ...next.indicatorLevel, subtotals: base.indicatorLevel.subtotals || next.indicatorLevel.subtotals },
        indicatorPresence: {
            leaf: new Set([...base.indicatorPresence.leaf, ...next.indicatorPresence.leaf]),
            subtotal: new Set([...base.indicatorPresence.subtotal, ...next.indicatorPresence.subtotal]),
        },
        leafKeys: base.leafKeys,
        fingerprint: next.fingerprint,
        measureKey: next.measureKey,
        measureNames: next.measureNames,
        columnsCut: base.columnsCut || next.columnsCut,
        orgLevelNames: next.orgLevelNames,
        orgs: Array.from(new Set([...base.orgs, ...next.orgs])),
        selection: mergeSelection(base.selection, next.selection),
    };
    factIndexes.set(merged, index);
    return merged;
}

/**
 * 行と値は同じで、金額・指標のメジャーの名前か書式の保存値（横持ちのメジャーごとの比較順・金額の持ち方、指標のメジャーごとの設定）
 * だけが変わった回（part）を、足し合わせた受け取り（input）に当て直す。1 回分ずつ受けた読み込みのあとは、ホストの手元に最後の回
 * しか無いので、読み直すと表が縮む。横持ちのイベントは queryName で、縦持ちの系列（イベント・メジャー）は
 * 欄の並びで結ぶ。指標の値は queryName のコードで持つので、名前と設定を差し替えるだけ。
 * input を書き換えて返す（金額の索引は捨てる。次に足し合わせるときに新しい名前で作り直す）
 */
export function refreshMeasures(input: InputData, part: InputData): InputData {
    const rename = new Map<string, string>();
    let events = input.events;
    if (input.eventSource !== "column") {
        events = input.events.map((event) => {
            const next = part.events.find((e) => e.measure && e.measure.queryName === event.measure?.queryName);
            if (!next) return event;
            if (next.name !== event.name) rename.set(event.name, next.name);
            return { ...event, name: next.name, measure: next.measure };
        });
    } else {
        events = input.events.map((event) => {
            if (!event.series) return event;
            const to = part.measureNames[event.series.index];
            if (to === undefined || to === event.series.name) return event;
            const base = event.name.slice(0, event.name.length - event.series.name.length - 1);
            const name = `${base}・${to}`;
            rename.set(event.name, name);
            return { ...event, name, series: { ...event.series, name: to } };
        });
    }
    if (rename.size > 0) {
        for (const fact of [...input.facts, ...input.indicatorFacts]) {
            const to = rename.get(fact.event);
            if (to !== undefined) fact.event = to;
        }
        factIndexes.delete(input);
    }
    input.events = events;
    input.indicators = part.indicators;
    input.measureKey = part.measureKey;
    input.measureNames = part.measureNames;
    input.duplicateMeasures = part.duplicateMeasures;
    return input;
}

function sameDefinition(a: AccountRecord, b: AccountRecord): boolean {
    return (
        a.name === b.name &&
        a.category === b.category &&
        a.subCategory === b.subCategory &&
        JSON.stringify(a.groups ?? []) === JSON.stringify(b.groups ?? []) &&
        a.balanceFlag === b.balanceFlag &&
        a.creditFlag === b.creditFlag
    );
}

/**
 * 組ごとの届いた順（sequences。どれも全体の並びの一部を順に持つ）から、全体の並びを決める。組の中で隣り合う 2 つを「前 → 後ろ」の関係にし、
 * 関係に沿って前から並べる（位相ソート）。次に置けるものが 2 つ以上あるのは、届いた順では前後が決まらない所：key（並びの欄の値）の小さい方を
 * 先にし、key でも決まらなければ（同じ値か、どちらかに値が無い）最初に届いた方を先にして guessed で知らせる。関係が輪になった（組によって前後が逆）ときも知らせる。
 * 決まる所は、組の届く順に左右されない（組の順につなぐ形では、本社にだけある区分の位置が届く順で変わり、計算行の
 * 範囲と値まで変わった）
 */
export function orderFromSequences(sequences: string[][], key: (value: string) => number = () => Number.POSITIVE_INFINITY): { order: string[]; guessed: boolean } {
    const seen = new Map<string, number>();
    const next = new Map<string, Set<string>>();
    const indegree = new Map<string, number>();
    for (const sequence of sequences) {
        sequence.forEach((value, i) => {
            if (!seen.has(value)) {
                seen.set(value, seen.size);
                next.set(value, new Set());
                indegree.set(value, 0);
            }
            const before = i > 0 ? sequence[i - 1] : undefined;
            if (before !== undefined && before !== value && !next.get(before)!.has(value)) {
                next.get(before)!.add(value);
                indegree.set(value, indegree.get(value)! + 1);
            }
        });
    }
    const rank = (value: string): [number, number] => [key(value), seen.get(value)!];
    const before = (a: string, b: string) => {
        const [ka, sa] = rank(a);
        const [kb, sb] = rank(b);
        return ka !== kb ? ka < kb : sa < sb;
    };
    const order: string[] = [];
    let guessed = false;
    const remaining = new Set(seen.keys());
    while (remaining.size > 0) {
        const ready = Array.from(remaining).filter((value) => indegree.get(value) === 0);
        // 輪になった：残りから選ぶ（知らせる）
        const pool = ready.length > 0 ? ready : Array.from(remaining);
        if (ready.length === 0) guessed = true;
        let pick = pool[0];
        for (const value of pool) if (before(value, pick)) pick = value;
        // 次に置けるものが 2 つ以上あり、並びの欄の値でも前後が決まらない（同じ値か、どちらかに値が無い）
        const undecided = (value: string) => key(value) === key(pick) || !Number.isFinite(key(value)) || !Number.isFinite(key(pick));
        if (ready.length > 1 && pool.some((value) => value !== pick && undecided(value))) guessed = true;
        order.push(pick);
        remaining.delete(pick);
        for (const after of next.get(pick)!) indegree.set(after, indegree.get(after)! - 1);
    }
    return { order, guessed };
}

/**
 * 科目の並び（record.order）を振る。区分 → 区分の中の中分類 → 中分類の中の科目の段ごとに、組ごとの届いた順から決める（orderFromSequences）。
 * どの段も届いた順で決まる所はそれ（区分に付けた「列で並べ替え」も効く）、決まらない所は並びの欄の値で決める（区分・中分類は、その中の
 * 科目の並びの一番小さい値）。並びの欄でも決めきれなかった所を返す（区分の名前か ORDER_OF_SECTIONS）
 */
export function assignOrder(accounts: Map<string, AccountRecord>, groups: Map<string, OrderGroup>, explicit: Map<string, number>): string[] {
    const INF = Number.POSITIVE_INFINITY;
    const guessed = new Set<string>();
    const list = Array.from(groups.values());
    const records = Array.from(accounts.values());
    // 区分・中分類の並びの欄の値：その中の科目の一番小さい値
    const minOf = (pick: (a: AccountRecord) => string) => {
        const map = new Map<string, number>();
        for (const a of records) map.set(pick(a), Math.min(map.get(pick(a)) ?? INF, explicit.get(a.code) ?? INF));
        return map;
    };
    const categoryKey = minOf((a) => a.category ?? "");
    const subKey = minOf((a) => orderBucket(a.category, a.subCategory));
    const rank = (sequences: string[][], key: (value: string) => number, label: string) => {
        const result = orderFromSequences(sequences, key);
        if (result.guessed) guessed.add(label);
        return new Map(result.order.map((v, i): [string, number] => [v, i]));
    };
    const categories = rank(
        list.map((g) => g.categories),
        (category) => categoryKey.get(category) ?? INF,
        ORDER_OF_SECTIONS
    );
    const subRanks = new Map<string, Map<string, number>>();
    const accountRanks = new Map<string, Map<string, number>>();
    const cached = (cache: Map<string, Map<string, number>>, key: string, make: () => Map<string, number>) => {
        let ranks = cache.get(key);
        if (!ranks) cache.set(key, (ranks = make()));
        return ranks;
    };
    const keyOf = (a: AccountRecord) => {
        const category = a.category ?? "";
        const bucket = orderBucket(a.category, a.subCategory);
        const subs = cached(subRanks, category, () =>
            rank(
                list.map((g) => g.subCategories.get(category) ?? []),
                (sub) => subKey.get(orderBucket(a.category, sub === "" ? null : sub)) ?? INF,
                category
            )
        );
        const codes = cached(accountRanks, bucket, () =>
            rank(
                list.map((g) => g.accounts.get(bucket) ?? []),
                (code) => explicit.get(code) ?? INF,
                category
            )
        );
        return [categories.get(category) ?? INF, subs.get(a.subCategory ?? "") ?? INF, codes.get(a.code) ?? INF];
    };
    const keys = new Map(records.map((a): [AccountRecord, number[]] => [a, keyOf(a)]));
    const sorted = Array.from(keys.entries()).sort(([, x], [, y]) => {
        for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return x[i] < y[i] ? -1 : 1;
        return 0;
    });
    sorted.forEach(([record], i) => {
        record.order = i;
    });
    return Array.from(guessed);
}
