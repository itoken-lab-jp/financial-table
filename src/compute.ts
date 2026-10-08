/**
 * 行 × 期間の主と比較の値と、その差。
 *
 * - 値を取る側（EventRef）はイベントの候補の並び。組織（事業など）× まとまり × 月ごとに、データのある最初の候補を使う。
 *   主の最新見込みなら比較順の大きい順、比較の個々のイベントはそのイベント → 比較順が下のイベント、1 つのイベントならそれだけ
 * - 「データがある」は組織 × まとまり × 月 × イベントで決める。まとまりはフローの科目全体と残高の科目全体と、指標のメジャー 1 本ずつ（別のファクトの
 *   人数が見通しの月に 0 にならない）。その組に 1 行でも金額があれば、金額の無い科目は 0（意図して省いたもの。科目ごとに別のイベントへ
 *   落とすと、同じ月に科目ごとに違うイベントが混ざる）。指標は、どのイベントにも無ければ空欄
 * - 月ごとの値を先に決めてから、期間の列は足すだけ（上期 + 下期 = 通期が崩れない）。
 *   フロー（合計）は月を足す、ストック（期末）は組織ごとに期間の最後の月の値、計算は集計した値で式を評価する（比率は割り直す）
 * - 差は、組織 × まとまりごとに主と比較の両方がある月だけで取る（予算・前年の無い月は主からも外す）。比較の値の列も同じ月の範囲
 * - 空欄と 0 を分ける。どの組織にもデータの無い期間は空欄
 * - 前年同期は、同じ候補の 12 か月前の月（月の名前ではなく月番号で引く）
 * - 小計・段階の行は、足す行（区分・中分類・科目）の借方プラスのデータの値を足し、行の向き（sign）で見せる。
 *   足す行ごとに自分のまとまりでイベントを採るので、フローと残高の区分をまたぐ段階の行も、表示した行の和と合う
 * - 比率の行は、分子と分母がどちらもフロー（合計）なら、組織ごとに、どちらにも値のある月だけで集計してから割る（指標はメジャー
 *   1 本ずつのまとまりなので、総労働時間の入っていない月・組織の営業利益を足さない）。残高（期末）は期間の
 *   最後の月を採るので、フロー ÷ 残高（売上 ÷ 総資産）の月はそろえない（そろえると通期の売上が残高のある月に切られた）
 */
import { Fact, INDICATOR_KEY } from "./data";
import { evaluateFormula } from "./formula";
import { MonthIndex } from "./periods";
import { RowDef, RowModel } from "./rows";

/** 値を取る側。events は候補（先が優先）、shift は月のずらし（前年同期は -12） */
export interface EventRef {
    events: string[];
    shift: number;
}

export interface Comparison {
    /** 比べる範囲（単位ごとに両側のある月）で足した主と比較。重ならなければ null */
    main: number | null;
    compare: number | null;
    diff: number | null;
}

type Series = Map<MonthIndex, number>;
/** 数字の有無を見るまとまり：フローの科目・残高の科目と、指標のメジャー 1 本ずつ（行のコード） */
type Group = string;
/** 組織 × まとまり（キー） */
type Slot = string;
/** 組織 × まとまり → 表の月 → 使うイベント */
type Resolution = Map<Slot, Map<MonthIndex, string>>;

interface Cell {
    sum: number;
    last: number;
}

export class Calculator {
    /** 科目・イベント → 組織 → 月 → 金額（データの符号のまま） */
    private readonly cells = new Map<string, Map<string, Map<MonthIndex, Cell>>>();
    /** イベント → 組織 × まとまり → データのある月 */
    private readonly presence = new Map<string, Map<Slot, Set<MonthIndex>>>();
    /** 組織 × まとまり → 組織とまとまり */
    private readonly slots = new Map<Slot, { org: string; group: Group }>();
    private readonly resolutions = new Map<string, Resolution>();
    private readonly seriesCache = new Map<string, Map<Slot, Series>>();
    /** 科目・イベント・組織 → データの月の値 */
    private readonly orgValueCache = new Map<string, Series>();

    constructor(private readonly model: RowModel, facts: Fact[]) {
        for (const fact of facts) {
            const slot = this.slotOf(fact.org, this.groupOf(fact.code));
            const byOrg = getOrSet(this.cells, key(fact.code, fact.event), () => new Map<string, Map<MonthIndex, Cell>>());
            const byMonth = getOrSet(byOrg, fact.org, () => new Map<MonthIndex, Cell>());
            const cell = byMonth.get(fact.month);
            if (cell) {
                cell.sum += fact.value;
                cell.last += fact.last;
            } else byMonth.set(fact.month, { sum: fact.value, last: fact.last });
            getOrSet(getOrSet(this.presence, fact.event, () => new Map<Slot, Set<MonthIndex>>()), slot, () => new Set<MonthIndex>()).add(fact.month);
        }
    }

    private row(code: string): RowDef | undefined {
        return this.model.rows.get(code);
    }

    /**
     * 値のまとまり：指標はメジャー 1 本ずつ（置く場所が無く行が無い指標も、科目のまとまりに混ぜない）。科目は期末の行（残高）かフローか
     */
    private groupOf(code: string): Group {
        if (code.startsWith(INDICATOR_KEY)) return code;
        return this.row(code)?.aggregation === "stock" ? "stock" : "flow";
    }

    /** 科目のまとまり（フロー・残高）の組か。累計の終わり・見出しの月ごとのイベントは科目の数字で決める（指標だけの月に伸ばさない） */
    private accountSlot(slot: Slot): boolean {
        return !this.slots.get(slot)!.group.startsWith(INDICATOR_KEY);
    }

    /** 科目のまとまりの組があるか（指標だけの表なら、指標の組も数える） */
    private hasAccountSlots(): boolean {
        for (const slot of this.slots.keys()) if (this.accountSlot(slot)) return true;
        return false;
    }

    private slotOf(org: string, group: Group): Slot {
        const slot = `${org}\u0000${group}`;
        if (!this.slots.has(slot)) this.slots.set(slot, { org, group });
        return slot;
    }

    /** 組織 × まとまりごとに、表の月それぞれで使うイベント（候補のうちデータのある最初のもの） */
    resolution(ref: EventRef): Resolution {
        const cacheKey = refKey(ref);
        const cached = this.resolutions.get(cacheKey);
        if (cached) return cached;
        const result: Resolution = new Map();
        for (const slot of this.slots.keys()) {
            const months = new Map<MonthIndex, string>();
            // 優先の低い候補から書き、高い候補で上書きする
            for (let i = ref.events.length - 1; i >= 0; i--) {
                for (const m of this.presence.get(ref.events[i])?.get(slot) ?? []) months.set(m - ref.shift, ref.events[i]);
            }
            if (months.size > 0) result.set(slot, months);
        }
        this.resolutions.set(cacheKey, result);
        return result;
    }

    /** 科目の、組織・イベントのデータの月の値（データの符号のまま）。期末の行は月の最後の日の値 */
    private orgValues(row: RowDef, event: string, org: string): Series {
        const cacheKey = `${key(row.code, event)}\u0000${org}`;
        const cached = this.orgValueCache.get(cacheKey);
        if (cached) return cached;
        const result: Series = new Map();
        for (const [m, cell] of this.cells.get(key(row.code, event))?.get(org) ?? []) result.set(m, row.aggregation === "stock" ? cell.last : cell.sum);
        this.orgValueCache.set(cacheKey, result);
        return result;
    }

    /**
     * 組織 × まとまりごとの月の値（データの符号のまま）。表の月で持つ（前年同期なら 12 か月前の値が今の月に入る）。
     * 科目は、自分のまとまりで採ったイベントの値（データのある組で金額の無い科目は 0）。
     * 小計・段階の行は、足す行それぞれの値に重みを掛けた和（どの足す行にも無い月は足さない）
     */
    private series(row: RowDef, ref: EventRef, stack = new Set<string>()): Map<Slot, Series> {
        const cacheKey = `${row.code}\u0000${refKey(ref)}`;
        const cached = this.seriesCache.get(cacheKey);
        if (cached) return cached;
        const result = new Map<Slot, Series>();
        // 足す行を持つ行：小計・段階・その他と、小計をうちにした行（picks.ts の UNDER_KEY）
        if (row.type === "subtotal" || row.type === "step" || row.type === "others" || (row.type === "breakdown" && row.summands.length > 0)) {
            if (stack.has(row.code)) return result;
            stack.add(row.code);
            for (const { code, weight } of row.summands) {
                const def = this.row(code);
                if (!def) continue;
                for (const [slot, series] of this.series(def, ref, stack)) {
                    const sum = getOrSet(result, slot, () => new Map<MonthIndex, number>());
                    for (const [m, v] of series) sum.set(m, (sum.get(m) ?? 0) + weight * v);
                }
            }
            stack.delete(row.code);
        } else if (row.type === "detail" || row.type === "breakdown" || row.type === "measure") {
            const group = this.groupOf(row.code);
            for (const [slot, months] of this.resolution(ref)) {
                const { org, group: slotGroup } = this.slots.get(slot)!;
                if (slotGroup !== group) continue;
                const series: Series = new Map();
                for (const [m, event] of months) series.set(m, this.orgValues(row, event, org).get(m + ref.shift) ?? 0);
                result.set(slot, series);
            }
        }
        this.seriesCache.set(cacheKey, result);
        return result;
    }

    /** 式の参照（コードか名前）を行のコードにする */
    private refCode(row: RowDef, name: string): string | undefined {
        return this.model.rows.has(name) ? name : row.refs.find((r) => this.model.rows.get(r)?.name === name);
    }

    /**
     * 表示の値（符号を掛けたあと）を、月の集まりで集計する。byOrg を渡すと、組織ごとにその月だけで集計する（比率の行の分子・分母）
     */
    aggregate(code: string, ref: EventRef, months: MonthIndex[], stack = new Set<string>(), byOrg?: Map<string, MonthIndex[]>): number | null {
        const row = this.row(code);
        if (!row || stack.has(code) || months.length === 0) return null;
        switch (row.type) {
            case "heading":
            case "blank":
                return null;
            case "calc": {
                if (!row.formula) return null;
                // フローの分子・分母が 2 つ以上なら、組織ごとに、どれにも値のある月だけで集計する
                const flows = this.flowRefs(row);
                const common = flows.length >= 2 ? this.commonByOrg(flows, ref, months) : undefined;
                if (common && common.size === 0) return null;
                stack.add(code);
                const value = evaluateFormula(row.formula, (name) => {
                    const target = this.refCode(row, name);
                    return target ? this.aggregate(target, ref, months, stack, common) : null;
                });
                stack.delete(code);
                return noNegativeZero(value);
            }
            default: {
                let total: number | null = null;
                for (const [slot, series] of this.series(row, ref)) {
                    const slotMonths = byOrg ? (byOrg.get(this.slots.get(slot)!.org) ?? []) : months;
                    if (slotMonths.length === 0) continue;
                    const v = aggregateSeries(series, row, slotMonths);
                    if (v !== null) total = (total ?? 0) + v;
                }
                return total === null ? null : signed(total, row.sign);
            }
        }
    }

    /** 主と比較。組織 × まとまりごとに両側のある月だけで集計してから足し、引く */
    compare(code: string, main: EventRef, compare: EventRef, months: MonthIndex[]): Comparison {
        const basis = this.basis(code, main, compare, this.basisMonths(main, compare, months), new Set());
        if (!basis) return { main: null, compare: null, diff: null };
        return { ...basis, diff: basis.main === null || basis.compare === null ? null : basis.main - basis.compare };
    }

    /** 組織 × まとまりごとの、比べる月（主と比較の両方にデータのある月） */
    private basisMonths(main: EventRef, compare: EventRef, months: MonthIndex[]): Map<Slot, MonthIndex[]> {
        const a = this.resolution(main);
        const b = this.resolution(compare);
        const result = new Map<Slot, MonthIndex[]>();
        for (const slot of this.slots.keys()) {
            const ma = a.get(slot);
            const mb = b.get(slot);
            if (!ma || !mb) continue;
            const overlap = months.filter((m) => ma.has(m) && mb.has(m));
            if (overlap.length > 0) result.set(slot, overlap);
        }
        return result;
    }

    private basis(
        code: string,
        main: EventRef,
        compare: EventRef,
        months: Map<Slot, MonthIndex[]>,
        stack: Set<string>
    ): { main: number | null; compare: number | null } | null {
        const row = this.row(code);
        if (!row || stack.has(code) || months.size === 0) return null;
        switch (row.type) {
            case "heading":
            case "blank":
                return null;
            case "calc": {
                if (!row.formula) return null;
                // フローの分子・分母が 2 つ以上なら、どれにも比べる月（主と比較の両方）のある月だけで集計する
                const flows = this.flowRefs(row);
                const aligned = flows.length >= 2 ? this.alignBasis(flows, main, months) : months;
                if (aligned.size === 0) return null;
                stack.add(code);
                const bases = new Map(row.refs.map((r) => [r, this.basis(r, main, compare, aligned, stack)]));
                stack.delete(code);
                if (Array.from(bases.values()).every((b) => b === null)) return null;
                const side = (pick: "main" | "compare") =>
                    noNegativeZero(
                        evaluateFormula(row.formula!, (name) => {
                            const target = this.refCode(row, name);
                            return target ? (bases.get(target)?.[pick] ?? null) : null;
                        })
                    );
                return { main: side("main"), compare: side("compare") };
            }
            default: {
                // 行のまとまりの組だけ（ほかのまとまりの組は series に無い）
                const side = (ref: EventRef) => {
                    const bySlot = this.series(row, ref);
                    let total: number | null = null;
                    for (const [slot, slotMonths] of months) {
                        const series = bySlot.get(slot);
                        if (!series) continue;
                        const v = aggregateSeries(series, row, slotMonths);
                        if (v !== null) total = (total ?? 0) + v;
                    }
                    return total === null ? null : signed(total, row.sign);
                };
                return { main: side(main), compare: side(compare) };
            }
        }
    }

    /** 行に値のある月（組織ごと。months のうち） */
    private presentByOrg(code: string, ref: EventRef, months: MonthIndex[]): Map<string, Set<MonthIndex>> {
        const row = this.row(code);
        const wanted = new Set(months);
        const result = new Map<string, Set<MonthIndex>>();
        if (!row) return result;
        for (const [slot, series] of this.series(row, ref)) {
            const present = getOrSet(result, this.slots.get(slot)!.org, () => new Set<MonthIndex>());
            for (const m of series.keys()) if (wanted.has(m)) present.add(m);
        }
        return result;
    }

    /** 比率の行の分子・分母のうち、フロー（合計）の行。月をそろえるのはこの行どうし（残高は期間の最後の月を採る） */
    private flowRefs(row: RowDef): string[] {
        return row.refs.filter((r) => this.row(r)?.aggregation === "flow");
    }

    /** 行のどれにも値のある月（組織ごと、months の並びのまま）。どれかに無い組織は入れない */
    private commonByOrg(refs: string[], ref: EventRef, months: MonthIndex[]): Map<string, MonthIndex[]> {
        const present = refs.map((r) => this.presentByOrg(r, ref, months));
        const result = new Map<string, MonthIndex[]>();
        for (const org of present[0]?.keys() ?? []) {
            const kept = months.filter((m) => present.every((byOrg) => byOrg.get(org)?.has(m) ?? false));
            if (kept.length > 0) result.set(org, kept);
        }
        return result;
    }

    /**
     * 比べる月（組織 × まとまりごと）を、比率の行の分子・分母のどれにもある月に絞る（組織ごと）。分子・分母それぞれの月は、その行の
     * まとまりの組の比べる月（主と比較の両方にある月）
     */
    private alignBasis(refs: string[], main: EventRef, months: Map<Slot, MonthIndex[]>): Map<Slot, MonthIndex[]> {
        const present = refs.map((r) => {
            const def = this.row(r);
            const byOrg = new Map<string, Set<MonthIndex>>();
            if (def) {
                for (const slot of this.series(def, main).keys()) {
                    const set = getOrSet(byOrg, this.slots.get(slot)!.org, () => new Set<MonthIndex>());
                    for (const m of months.get(slot) ?? []) set.add(m);
                }
            }
            return byOrg;
        });
        const result = new Map<Slot, MonthIndex[]>();
        for (const [slot, slotMonths] of months) {
            const org = this.slots.get(slot)!.org;
            const kept = slotMonths.filter((m) => present.every((byOrg) => byOrg.get(org)?.has(m) ?? false));
            if (kept.length > 0) result.set(slot, kept);
        }
        return result;
    }

    /**
     * 科目・指標の行に、表の月で採ったイベントの数字があるか（refs のどれか）。データのある組で数字の無い科目は 0 で出るが、それは数えない。
     * skipZero なら 0 の数字も無いとみなす
     */
    hasValue(code: string, refs: EventRef[], months: Iterable<MonthIndex>, skipZero: boolean): boolean {
        const group = this.groupOf(code);
        const shown = Array.from(months);
        for (const ref of refs) {
            for (const [slot, byMonth] of this.resolution(ref)) {
                const { org, group: slotGroup } = this.slots.get(slot)!;
                if (slotGroup !== group) continue;
                for (const m of shown) {
                    const event = byMonth.get(m);
                    if (event === undefined) continue;
                    const cell = this.cells.get(key(code, event))?.get(org)?.get(m + ref.shift);
                    if (cell && (!skipZero || cell.sum !== 0 || cell.last !== 0)) return true;
                }
            }
        }
        return false;
    }

    /**
     * 行の値に使ったイベント（月ごと、候補の並びの順）。行のまとまりの組だけで見る（指標はメジャー 1 本ずつなので、列の見出しの
     * イベントと違うことがある。セルのツールチップで名乗る）
     */
    rowEventsByMonth(code: string, ref: EventRef, months: MonthIndex[]): Map<MonthIndex, string[]> {
        const row = this.row(code);
        const result = new Map<MonthIndex, string[]>();
        if (!row) return result;
        const resolution = this.resolution(ref);
        const slots = Array.from(this.series(row, ref).keys());
        for (const m of months) {
            const used = new Set<string>();
            for (const slot of slots) {
                const event = resolution.get(slot)?.get(m);
                if (event !== undefined) used.add(event);
            }
            if (used.size > 0) result.set(m, ref.events.filter((e) => used.has(e)));
        }
        return result;
    }

    /** 期間の月ごとに使ったイベント（どの組織・まとまりかを問わず、候補の並びの順。科目のまとまりがあれば科目だけで見る） */
    eventsByMonth(ref: EventRef, months: MonthIndex[]): Map<MonthIndex, string[]> {
        const resolution = this.resolution(ref);
        const accountsOnly = this.hasAccountSlots();
        const result = new Map<MonthIndex, string[]>();
        for (const m of months) {
            const used = new Set<string>();
            for (const [slot, byMonth] of resolution) {
                if (accountsOnly && !this.accountSlot(slot)) continue;
                const event = byMonth.get(m);
                if (event !== undefined) used.add(event);
            }
            if (used.size > 0) result.set(m, ref.events.filter((e) => used.has(e)));
        }
        return result;
    }

    /** 比べる範囲（主と比較の両方がある月）で、比較の側が使ったイベント（科目のまとまりがあれば科目だけで見る） */
    compareEventsByMonth(main: EventRef, compare: EventRef, months: MonthIndex[]): Map<MonthIndex, string[]> {
        const basis = this.basisMonths(main, compare, months);
        const resolution = this.resolution(compare);
        const accountsOnly = this.hasAccountSlots();
        const result = new Map<MonthIndex, string[]>();
        for (const m of months) {
            const used = new Set<string>();
            for (const [slot, slotMonths] of basis) {
                if (!slotMonths.includes(m) || (accountsOnly && !this.accountSlot(slot))) continue;
                const event = resolution.get(slot)?.get(m);
                if (event !== undefined) used.add(event);
            }
            if (used.size > 0) result.set(m, compare.events.filter((e) => used.has(e)));
        }
        return result;
    }

    /**
     * 表の月のうち値のある最後の月。accept を渡すと、そのイベントを使った月だけを見る。無ければ null。
     * 科目のまとまりがあれば科目だけで見る（人数だけが先の月まである、で累計の終わり・年度の既定を伸ばさない）
     */
    lastMonth(ref: EventRef, accept: (event: string) => boolean = () => true): MonthIndex | null {
        let last: MonthIndex | null = null;
        const accountsOnly = this.hasAccountSlots();
        for (const [slot, byMonth] of this.resolution(ref)) {
            if (accountsOnly && !this.accountSlot(slot)) continue;
            for (const [m, event] of byMonth) if (accept(event) && (last === null || m > last)) last = m;
        }
        return last;
    }

    /** データにある月（すべてのイベント）。科目のまとまりがあれば科目だけで見る（指標だけの月で年度の選択肢を増やさない） */
    months(): MonthIndex[] {
        const accountsOnly = this.hasAccountSlots();
        const all = new Set<MonthIndex>();
        for (const bySlot of this.presence.values()) {
            for (const [slot, months] of bySlot) if (!accountsOnly || this.accountSlot(slot)) for (const m of months) all.add(m);
        }
        return Array.from(all).sort((a, b) => a - b);
    }

    /** そのイベントに、月のどれかでデータがあるか（どの組織でも。科目のまとまりがあれば科目だけで見る：比較の候補に指標だけのイベントを並べない） */
    hasData(event: string, months: MonthIndex[]): boolean {
        const bySlot = this.presence.get(event);
        if (!bySlot) return false;
        const accountsOnly = this.hasAccountSlots();
        for (const [slot, present] of bySlot) if ((!accountsOnly || this.accountSlot(slot)) && months.some((m) => present.has(m))) return true;
        return false;
    }
}

/** 1 つの組織 × まとまりの月の値を期間で集計する（符号はまだ掛けない）。期末は期間のうち値のある最後の月 */
function aggregateSeries(series: Series, row: RowDef, months: MonthIndex[]): number | null {
    if (row.aggregation === "stock") {
        for (let i = months.length - 1; i >= 0; i--) {
            const v = series.get(months[i]);
            if (v !== undefined) return v;
        }
        return null;
    }
    let sum: number | null = null;
    for (const m of months) {
        const v = series.get(m);
        if (v !== undefined) sum = (sum ?? 0) + v;
    }
    return sum;
}

/** -0 を 0 にする（-0 は ▲0 と書かれてしまう） */
function noNegativeZero(value: number | null): number | null {
    return value === null ? null : value || 0;
}

/** 符号を掛ける。0 に -1 を掛けた -0 は 0 にする（-0 は ▲0 と書かれてしまう） */
function signed(value: number, sign: number): number {
    return value * sign || 0;
}

function getOrSet<K, V>(map: Map<K, V>, k: K, make: () => V): V {
    let v = map.get(k);
    if (v === undefined) map.set(k, (v = make()));
    return v;
}

function key(code: string, event: string): string {
    return `${code}\u0000${event}`;
}

function refKey(ref: EventRef): string {
    return `${ref.shift}\u0000${ref.events.join("\u0001")}`;
}
