/**
 * 行別の数値書式の書式と、セルの文字。
 *
 * 科目の表の「書式」の列：
 * - 空（または「金額」）… 金額。表示単位（万・億など）で割る。桁は書式ペインの小数点以下の桁数
 * - % を含む（0.0%）… 比率。100 倍して % を付ける。差はポイント（pt）
 * - ほかの字を付けた数（#,0.0h・0.00倍・#,0人）… 単位で割らない数。桁は書式の小数の 0 の数
 *
 * 単位の違う行を表示単位で割らない（時間や率を千円で割ると別の数字になる）。
 */
import { NEGATIVE_STYLES, SignStyle, ZERO_STYLES, formatSigned, shownSignOf } from "./shared/numberFormat";

export type ValueKind = "amount" | "percent" | "number";

export interface RowFormat {
    kind: ValueKind;
    /** 小数の桁。金額で null なら書式ペインの桁 */
    decimals: number | null;
    prefix: string;
    suffix: string;
    grouping: boolean;
}

export const AMOUNT_FORMAT: RowFormat = { kind: "amount", decimals: null, prefix: "", suffix: "", grouping: true };

/**
 * 読めない書式は金額として扱う（落とさない）。読めたかどうかは ok で返す。
 * Excel の書式をそのまま貼れるように、; のあと（マイナスの書き方）は見ず、"人" の引用符と \ は外す。
 * 字の付かない数の形（#,##0.0）は、桁だけを決めた金額（表示単位で割る）
 */
export function parseRowFormat(text: string | null): { format: RowFormat; ok: boolean } {
    const src = (text ?? "").split(";")[0].trim();
    if (src === "" || src === "値" || src === "金額") return { format: AMOUNT_FORMAT, ok: true };
    const match = /[#0,.]*[#0][#0,.]*/.exec(src);
    if (!match) return { format: AMOUNT_FORMAT, ok: false };
    const pattern = match[0];
    const dot = pattern.indexOf(".");
    const decimals = dot < 0 ? 0 : pattern.slice(dot + 1).replace(/[^#0]/g, "").length;
    const literal = (s: string) => s.replace(/["\\]/g, "");
    const prefix = literal(src.slice(0, match.index));
    const suffix = literal(src.slice(match.index + pattern.length));
    const grouping = pattern.includes(",");
    if (suffix.includes("%") || prefix.includes("%")) {
        return { format: { kind: "percent", decimals, prefix: prefix.replace("%", ""), suffix: suffix.replace("%", ""), grouping }, ok: true };
    }
    if (prefix.trim() === "" && suffix.trim() === "") return { format: { ...AMOUNT_FORMAT, decimals }, ok: true };
    return { format: { kind: "number", decimals, prefix, suffix, grouping }, ok: true };
}

/** 0 から遠い方へ丸める（toLocaleString と同じ。+0.5 と -0.5 で扱いを変えない） */
export function roundHalfAway(value: number, decimals: number): number {
    const factor = Math.pow(10, decimals);
    const rounded = (Math.sign(value) * Math.round(Math.abs(value) * factor)) / factor;
    return rounded === 0 ? 0 : rounded;
}

/** 金額の表示単位。divisor で割り、decimals の桁で出す */
export interface AmountUnit {
    divisor: number;
    decimals: number;
}

function plain(value: number, decimals: number, grouping: boolean): string {
    return value.toLocaleString("ja-JP", { minimumFractionDigits: decimals, maximumFractionDigits: decimals, useGrouping: grouping });
}

/**
 * 表に出る桁で丸めた値（金額は表示単位で割ったあと、比率は % の値）。
 * 丸めて 0 になる差を「+0.0」「-0.0」と書かず、色も付けないために使う
 */
export function shownValue(value: number, format: RowFormat, unit: AmountUnit): number {
    const [scaled, decimals] =
        format.kind === "amount"
            ? [value / unit.divisor, format.decimals ?? unit.decimals]
            : format.kind === "percent"
              ? [value * 100, format.decimals ?? 1]
              : [value, format.decimals ?? 0];
    return roundHalfAway(scaled, decimals);
}

/**
 * マイナスと 0 の書き方（書式ペインの「数値」）。値と差で 0 の書き方を分ける（値は 0、差は ±0 が帳票の習い）。
 * negativeZero は丸めて 0 になるマイナスに符号を残す（▲0）
 */
export interface SignOptions {
    negative: string;
    zero: string;
    diffZero: string;
    negativeZero: boolean;
}

/** 記号を付けない書き方（-1,234、0、差も 0。丸めて 0 のマイナスは 0） */
export const PLAIN_SIGN: SignOptions = { negative: NEGATIVE_STYLES.minus, zero: ZERO_STYLES.zero, diffZero: ZERO_STYLES.zero, negativeZero: false };

/** 行の書式に合わせて、単位で割り、マイナスと 0 の書き方を当てる。suffix は比率の行の %・pt */
function signed(value: number, format: RowFormat, unit: AmountUnit, style: Partial<SignStyle>, percentSuffix: string): string {
    switch (format.kind) {
        case "amount":
            // 字は単位を数字のあとに置くとき（1,234百万円）だけ。ふつうの金額は字を持たない
            return formatSigned(value, unit.divisor, String(format.decimals ?? unit.decimals), style, format.suffix);
        case "percent":
            return `${format.prefix}${formatSigned(value * 100, 1, String(format.decimals ?? 1), style, percentSuffix + format.suffix)}`;
        case "number":
            return `${format.prefix}${formatSigned(value, 1, String(format.decimals ?? 0), style, format.suffix)}`;
    }
}

/**
 * 比率の行（% の書式の指標・計算行の比率）の値と差を、比・率と同じ上限で止める（「≧999%」「≦▲999pt」）。
 * 金額・数の行と、上限を渡さないとき（ツールチップの実の値）は止めない
 */
function capped(value: number, format: RowFormat, unit: AmountUnit, style: Partial<SignStyle>, suffix: string, cap: number | undefined): string {
    if (cap === undefined || format.kind !== "percent" || Math.abs(value) <= cap) return signed(value, format, unit, style, suffix);
    return `${value > 0 ? "≧" : "≦"}${signed(value > 0 ? cap : -cap, { ...format, decimals: 0 }, unit, style, suffix)}`;
}

/** 行の値の文字。空欄は空文字。cap を渡すと、比率の行の値をその上限（倍）で止める */
export function formatRowValue(value: number | null, format: RowFormat, unit: AmountUnit, sign: SignOptions = PLAIN_SIGN, cap?: number): string {
    if (value === null) return "";
    return capped(value, format, unit, { negative: sign.negative, zero: sign.zero, negativeZero: sign.negativeZero, plus: false }, "%", cap);
}

/** 差の文字。正に + を付ける。比率の行の差はポイント。cap を渡すと、比率の行の差をその上限（ポイント）で止める */
export function formatDiff(diff: number | null, format: RowFormat, unit: AmountUnit, sign: SignOptions = PLAIN_SIGN, cap?: number): string {
    if (diff === null) return "";
    // 丸めて 0 になるプラスは「+0」と書かない（0 の書き方にそろえる）
    if (diff > 0 && shownValue(diff, format, unit) === 0) diff = 0;
    return capped(diff, format, unit, { negative: sign.negative, zero: sign.diffZero, negativeZero: sign.negativeZero, plus: true }, "pt", cap);
}

/**
 * 差の見える符号（-1・0・1）。formatDiff が書く字と同じ丸めで決める：丸めて 0 のプラスは 0、
 * 丸めて 0 のマイナスは、▲0 を残す設定なら -1（見た目の符号と色をそろえる）
 */
export function diffSign(diff: number | null, format: RowFormat, unit: AmountUnit, sign: SignOptions = PLAIN_SIGN): -1 | 0 | 1 {
    if (diff === null) return 0;
    const style = { negativeZero: sign.negativeZero };
    switch (format.kind) {
        case "amount":
            return shownSignOf(diff, unit.divisor, String(format.decimals ?? unit.decimals), style);
        case "percent":
            return shownSignOf(diff * 100, 1, String(format.decimals ?? 1), style);
        case "number":
            return shownSignOf(diff, 1, String(format.decimals ?? 0), style);
    }
}

/** 比・率の上限の既定（倍）。普段の「101.2%」と同じ 6 文字に収まる 3 けたで止める。書式ペインで変えられる */
export const RATIO_CAP = 9.99;

/** 上限の % の値（9.99 → 999） */
const capPercent = (cap: number) => Math.round(cap * 100);

/** 出せない比・率（符号がまたがる、比べる値が 0） */
export const NOT_AVAILABLE = "―";

/**
 * 比（101.2%）。実績 ÷ 比較。符号がまたがるときと比較が 0 のときは「―」、大きすぎれば「≧999%」。
 * 「出せない」と「大きすぎて丸めた」は別の字にする
 */
export function formatRatio(main: number | null, compare: number | null, decimals = 1, cap = RATIO_CAP): string {
    if (main === null || compare === null) return "";
    if (compare === 0 || main * compare < 0) return NOT_AVAILABLE;
    const ratio = main / compare;
    if (ratio > cap) return `≧${plain(capPercent(cap), 0, true)}%`;
    return `${plain(ratio * 100, decimals, true)}%`;
}

/** 率（+1.2%）。（実績 − 比較）÷ |比較|。比較が 0 なら「―」。比較が負でも向きが保たれる。マイナスと 0 は差と同じ書き方 */
export function formatRate(main: number | null, compare: number | null, sign: SignOptions = PLAIN_SIGN, decimals = 1, cap = RATIO_CAP): string {
    if (main === null || compare === null) return "";
    if (compare === 0) return NOT_AVAILABLE;
    const rate = (main - compare) / Math.abs(compare);
    const style = { negative: sign.negative, zero: sign.diffZero, negativeZero: sign.negativeZero, plus: true };
    const limit = capPercent(cap);
    if (rate > cap) return `≧${formatSigned(limit, 1, "0", style, "%")}`;
    if (rate < -cap) return `≦${formatSigned(-limit, 1, "0", style, "%")}`;
    return formatSigned(rate > 0 && roundHalfAway(rate * 100, decimals) === 0 ? 0 : rate * 100, 1, String(decimals), style, "%");
}

/** ツールチップの比・率・比率の行に足す桁（表は小数 1 桁、ツールチップは 3 桁） */
export const EXACT_EXTRA_DECIMALS = 2;

/**
 * セルのツールチップの正確な値。表は表示単位で丸めるので、金額は単位で割らずに通貨の字を付ける。
 * 金額と数（人・時間）は端数があれば小数 2 桁、比率の行は表より 2 桁多く出す。マイナスと 0 は表と同じ書き方（丸めて 0 のマイナスは符号を残す）
 */
export function formatExact(value: number | null, format: RowFormat, currency: string, sign: SignOptions = PLAIN_SIGN, diff = false): string {
    if (value === null) return "";
    const style = { negative: sign.negative, zero: diff ? sign.diffZero : sign.zero, negativeZero: true, plus: diff };
    const decimals = (min: number) => String(Math.max(min, Number.isInteger(roundHalfAway(value, 2)) ? 0 : 2));
    switch (format.kind) {
        case "amount":
            return formatSigned(value, 1, decimals(0), style, currency);
        case "percent":
            return `${format.prefix}${formatSigned(value * 100, 1, String((format.decimals ?? 1) + EXACT_EXTRA_DECIMALS), style, (diff ? "pt" : "%") + format.suffix)}`;
        case "number":
            return `${format.prefix}${formatSigned(value, 1, decimals(format.decimals ?? 0), style, format.suffix)}`;
    }
}
