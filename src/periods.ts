/**
 * 期間。データの月を「年 × 12 + 月」の通し番号（月番号）で持ち、年度・四半期・半期・累計はここで組む。
 *
 * 期首の月と年度の呼び方は決め打ちしない（書式で選ぶ）。月の名前だけで引くと、別の年度の同じ月が混ざるので、
 * 前年同期も「月番号 − 12」で引く（年度つきのキー）。
 */

/** 月番号。2025 年 4 月 = 2025 × 12 + 3 */
export type MonthIndex = number;

export const monthIndex = (year: number, month1to12: number): MonthIndex => year * 12 + (month1to12 - 1);
export const yearOf = (m: MonthIndex): number => Math.floor(m / 12);
export const monthOf = (m: MonthIndex): number => (m % 12) + 1;

/** 月と、その月の中の日（月だけの値なら 0）。日付の列で 1 か月に何行も届くとき、期末の行は最後の日の値を取る */
export interface MonthDay {
    month: MonthIndex;
    day: number;
}

/**
 * 月の値を月番号にする。読めなければ null。
 * 日付（Date）・日付の文字（2025-04-01T00:00:00 など）・2025-04・2025/4・2025年4月・202504・20250401（文字でも数でも）を読む
 */
export function toMonthIndex(value: unknown): MonthIndex | null {
    return toMonthDay(value)?.month ?? null;
}

export function toMonthDay(value: unknown): MonthDay | null {
    if (value === null || value === undefined || value === "") return null;
    if (value instanceof Date) {
        if (Number.isNaN(value.getTime())) return null;
        // ふつうはその場所の 0 時で届く。その場所の 0 時でなく UTC の 0 時なら UTC で読む
        // （UTC の 0 時の日付を西の時間帯で読むと、前の日＝前の月にずれる）
        const localMidnight = value.getHours() === 0 && value.getMinutes() === 0 && value.getSeconds() === 0;
        const utcMidnight = value.getUTCHours() === 0 && value.getUTCMinutes() === 0 && value.getUTCSeconds() === 0;
        if (!localMidnight && utcMidnight) {
            return { month: monthIndex(value.getUTCFullYear(), value.getUTCMonth() + 1), day: value.getUTCDate() };
        }
        return { month: monthIndex(value.getFullYear(), value.getMonth() + 1), day: value.getDate() };
    }
    if (typeof value === "number") {
        if (!Number.isInteger(value)) return null;
        // 20250401（カレンダーの表の日付のキー）と 202504。年だけ（2025）は月が分からないので読まない
        if (value >= 10000101 && value <= 99991231) return fromParts(Math.floor(value / 10000), Math.floor(value / 100) % 100, value % 100);
        if (value >= 100001 && value <= 999912) return fromParts(Math.floor(value / 100), value % 100, 0);
        return null;
    }
    const text = String(value).trim();
    // 時刻つきの ISO の日時（2024-03-31T15:00:00.000Z）は日時として読む。Power BI Desktop は日付の列を、その場所の 0 時を
    // UTC に直した文字で届ける（日本なら前の日の 15 時）。先頭の年月だけを読むと、前の月になる（2026-09-26 に Desktop で確かめた）
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(text)) {
        const date = new Date(text);
        if (!Number.isNaN(date.getTime())) return toMonthDay(date);
    }
    const match = /^(\d{4})\s*(?:[-/.年]\s*(\d{1,2})(?:\s*[-/.月]\s*(\d{1,2}))?|(\d{2})(\d{2})?)(?:\D|$)/.exec(text);
    if (!match) return null;
    return fromParts(Number(match[1]), Number(match[2] ?? match[4]), Number(match[3] ?? match[5] ?? 0));
}

function fromParts(year: number, month: number, day: number): MonthDay | null {
    if (month < 1 || month > 12 || day < 0 || day > 31) return null;
    return { month: monthIndex(year, month), day };
}

/** 年度の呼び方。start = 始まりの年で「2025年度」、end = 終わりの年で「FY2026」 */
export type YearLabel = "start" | "end";

export interface FiscalCalendar {
    /** 期首の月（1〜12） */
    startMonth: number;
    yearLabel: YearLabel;
}

/** 月が属する年度の、期首の月の月番号 */
export function fiscalYearStart(m: MonthIndex, cal: FiscalCalendar): MonthIndex {
    const offset = (monthOf(m) - cal.startMonth + 12) % 12;
    return m - offset;
}

/** 年度の見出し。start の年で 2025年度、end の年で FY2026（期首が 1 月なら同じ年） */
export function fiscalYearName(start: MonthIndex, cal: FiscalCalendar): string {
    if (cal.yearLabel === "end") return `FY${yearOf(start + 11)}`;
    return `${yearOf(start)}年度`;
}

/** sum は見る人が選んだ、出している月を足した列 */
export type PeriodKind = "month" | "quarter" | "half" | "year" | "ytd" | "sum";

export interface PeriodColumn {
    key: string;
    kind: PeriodKind;
    label: string;
    /** この列が覆う月（昇順） */
    months: MonthIndex[];
}

export interface PeriodOptions {
    showMonths: boolean;
    showQuarters: boolean;
    showHalves: boolean;
    showYear: boolean;
    showYtd: boolean;
}

/**
 * 年度の列を並べる。月のあとに集計の列をまとめる：4月 … 3月 1Q 2Q 3Q 4Q 上期 下期 通期 累計。
 * 累計は期首から lastMonth（主にデータのある最後の月）まで。lastMonth がこの年度に無ければ累計は出さない
 */
export function buildPeriodColumns(fyStart: MonthIndex, options: PeriodOptions, lastMonth: MonthIndex | null, calendar?: FiscalCalendar): PeriodColumn[] {
    const columns = buildYearColumns(fyStart, options, lastMonth);
    // 年度をまたぐ表（見る人が選んだ）では、集計の列のキーに年度を付け（1Q が 2 つ並んでもぶつからない）、見出しに年・年度を添える
    //
    if (!calendar) return columns;
    const prefix = `${yearOf(fyStart)}-${String(monthOf(fyStart)).padStart(2, "0")}`;
    const year = fiscalYearName(fyStart, calendar);
    return columns.map((c) =>
        c.kind === "month"
            ? { ...c, label: `${String(yearOf(c.months[0])).slice(-2)}/${String(monthOf(c.months[0])).padStart(2, "0")}` }
            : { ...c, key: `${prefix}:${c.key}`, label: `${year} ${c.label}` }
    );
}

function buildYearColumns(fyStart: MonthIndex, options: PeriodOptions, lastMonth: MonthIndex | null): PeriodColumn[] {
    const columns: PeriodColumn[] = [];
    const range = (from: number, count: number) => Array.from({ length: count }, (_, i) => fyStart + from + i);
    if (options.showMonths) {
        for (const m of range(0, 12)) columns.push({ key: `m${m}`, kind: "month", label: `${monthOf(m)}月`, months: [m] });
    }
    if (options.showQuarters) {
        for (let q = 0; q < 4; q++) columns.push({ key: `q${q + 1}`, kind: "quarter", label: `${q + 1}Q`, months: range(q * 3, 3) });
    }
    if (options.showHalves) {
        columns.push({ key: "h1", kind: "half", label: "上期", months: range(0, 6) }, { key: "h2", kind: "half", label: "下期", months: range(6, 6) });
    }
    if (options.showYear) columns.push({ key: "fy", kind: "year", label: "通期", months: range(0, 12) });
    if (options.showYtd && lastMonth !== null && lastMonth >= fyStart && lastMonth < fyStart + 12) {
        const count = lastMonth - fyStart + 1;
        const label = count === 12 ? "累計" : `累計（${monthOf(fyStart)}〜${monthOf(lastMonth)}月）`;
        columns.push({ key: "ytd", kind: "ytd", label, months: range(0, count) });
    }
    return columns;
}

/** すべての列（見る人が選ぶときの候補） */
export const ALL_PERIODS: PeriodOptions = { showMonths: true, showQuarters: true, showHalves: true, showYear: true, showYtd: true };

const ym = (m: MonthIndex) => `${yearOf(m)}-${String(monthOf(m)).padStart(2, "0")}`;

/**
 * 期間の名前。
 * 月は「2025-06」、四半期・半期・通期は種類と最初の月（quarter:2025-07）、累計は期首の月（ytd:2025-04。月が進んでも同じ）、
 * 合計の列は月の範囲と数（sum:2025-04:2025-06:3）
 */
export function periodIdOf(period: { kind: PeriodKind; months: MonthIndex[] }): string {
    if (period.kind === "month") return ym(period.months[0]);
    // 合計の列は足した月の範囲と数で（月を選び直すと別の期間。この期間だけの比較対象が別の月の組に効かない）
    if (period.kind === "sum") return `sum:${ym(period.months[0])}:${ym(period.months[period.months.length - 1])}:${period.months.length}`;
    return `${period.kind}:${ym(period.months[0])}`;
}

/** 出している月を足した列。月がつながっていれば「4〜6月計」（12 か月を超えれば年を付けて「2024年4月〜2025年4月計」）、飛んでいれば「5か月計」 */
export function sumColumn(months: MonthIndex[]): PeriodColumn | null {
    const sorted = Array.from(new Set(months)).sort((a, b) => a - b);
    if (sorted.length === 0) return null;
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    const joined = last - first + 1 === sorted.length;
    const label = !joined
        ? `${sorted.length}か月計`
        : first === last
          ? `${monthOf(first)}月計`
          : last - first >= 12
            ? `${yearOf(first)}年${monthOf(first)}月〜${yearOf(last)}年${monthOf(last)}月計`
            : `${monthOf(first)}〜${monthOf(last)}月計`;
    return { key: "sum", kind: "sum", label, months: sorted };
}

/** 見る人が選んだ期間（保存値）。base は選んだときの書式ペインの期間の設定。作り手が変えたら捨てる */
export interface PeriodPick {
    base: string;
    /** 出す列の名前（periodIdOf） */
    columns: string[];
    /** 出している月を足した列を出すか */
    total: boolean;
}

/** 見る人のダイアログに並べる年度と、その年度の列の候補（年度は新しい順） */
export interface PeriodChoiceYear {
    key: string;
    name: string;
    columns: Array<{ id: string; kind: PeriodKind; label: string; months: MonthIndex[] }>;
}

export interface TablePeriods {
    columns: PeriodColumn[];
    /** 表に出す年度（期首の月、古い順） */
    starts: MonthIndex[];
    title: string;
    /** 見る人の選択で組んだか（選択が無い・選んだ列が 1 つも無く書式ペインのままに戻したときは false） */
    fromPick: boolean;
}

/**
 * 表の期間の列。見る人の選択が無ければ、書式ペインの年度と出す列のまま。選択があれば、データのある年度の列の候補から選んだものを
 * 時間の順に並べ（年度の中は今の並び）、合計の列を最後に足す。見る人がすべて外したら列の無い表（書式ペインに戻すと、次にチェックを
 * 入れたとき書式ペインの列から始まり逆に効いた。Desktop で確かめたとき）。選んだ列がデータから無くなった（年度が無くなったなど）ときだけ、
 * 書式ペインのまま
 */
export function tablePeriods(
    starts: MonthIndex[],
    selectedStart: MonthIndex,
    author: PeriodOptions,
    lastMonth: MonthIndex | null,
    calendar: FiscalCalendar,
    pick: PeriodPick | null
): TablePeriods {
    const authorTable = (): TablePeriods => ({
        columns: buildPeriodColumns(selectedStart, author, lastMonth),
        starts: [selectedStart],
        title: fiscalYearName(selectedStart, calendar),
        fromPick: false,
    });
    if (!pick) return authorTable();
    if (pick.columns.length === 0) return { columns: [], starts: [selectedStart], title: fiscalYearName(selectedStart, calendar), fromPick: true };
    const chosen = new Set(pick.columns);
    const ascending = [...starts].sort((a, b) => a - b);
    const byYear = ascending.map((start) => ({ start, columns: buildPeriodColumns(start, ALL_PERIODS, lastMonth).filter((c) => chosen.has(periodIdOf(c))) }));
    const used = byYear.filter((y) => y.columns.length > 0);
    if (used.length === 0) return authorTable();
    const multi = used.length > 1;
    const byOrder = used.flatMap((y) =>
        multi ? buildPeriodColumns(y.start, ALL_PERIODS, lastMonth, calendar).filter((c) => chosen.has(periodIdOf(c))) : y.columns
    );
    // 年度をまたいでも、月をすべて先に並べ、集計の列はそのあとに年度の順でまとめる
    const columns = [...byOrder.filter((c) => c.kind === "month"), ...byOrder.filter((c) => c.kind !== "month")];
    const total = pick.total ? sumColumn(columns.filter((c) => c.kind === "month").flatMap((c) => c.months)) : null;
    if (total) columns.push(total);
    const months = columns.flatMap((c) => c.months);
    const first = Math.min(...months);
    const last = Math.max(...months);
    const title = multi
        ? `${fiscalYearName(fiscalYearStart(first, calendar), calendar)} ${monthOf(first)}月〜${fiscalYearName(fiscalYearStart(last, calendar), calendar)} ${monthOf(last)}月`
        : fiscalYearName(used[0].start, calendar);
    return { columns, starts: used.map((y) => y.start), title, fromPick: true };
}

/** 見る人のダイアログの候補：データのある年度（新しい順）ごとの、すべての列 */
export function periodChoices(starts: MonthIndex[], lastMonth: MonthIndex | null, calendar: FiscalCalendar): PeriodChoiceYear[] {
    return [...starts]
        .sort((a, b) => b - a)
        .map((start) => ({
            key: ym(start),
            name: fiscalYearName(start, calendar),
            columns: buildPeriodColumns(start, ALL_PERIODS, lastMonth).map((c) => ({ id: periodIdOf(c), kind: c.kind, label: c.label, months: c.months })),
        }));
}
