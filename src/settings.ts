"use strict";

/**
 * 書式ペイン。カードの name とスライスの name は capabilities.json の objects と完全に一致させる
 * 。
 *
 * 主と比較・年度の選択肢と、横持ちのメジャーごとの比較順、計算行の挿入位置・分子・分母の選択肢、指標のメジャーごとの設定は
 * データ次第なので、update() のたびに viewModel の結果から流し込む（applyData）。
 */

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import SimpleCard = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;
import { DIRECTION_ITEMS, HierarchyFormat, HierarchyOverrides, ROOT_ITEMS, STYLE_ITEMS, TOTAL_ITEMS, dropdown } from "./hierarchySettings";

import { UNIT_TYPES, UNIT_NOTATIONS, PRECISIONS } from "./shared/units";
import {
    DEFAULT_BAD_COLOR,
    DEFAULT_GOOD_COLOR,
    DIFF_TONE_MODE_ITEMS,
    NEGATIVE_STYLES,
    NEGATIVE_STYLE_ITEMS,
    TONE_MODES,
    ZERO_STYLES,
    ZERO_STYLE_ITEMS,
} from "./shared/numberFormat";

import { SCROLL_START_ITEMS } from "./shared/scrollStart";
import type { Theme } from "./theme";
import { MeasureSetting, ORDER_NONE } from "./events";
import { CALC_KEY, CalcChoices } from "./rows";

export { UNIT_TYPES, UNIT_NOTATIONS, PRECISIONS };

export const DEFAULT_FONT_FAMILY = '"Segoe UI", wf_segoe-ui_normal, helvetica, arial, sans-serif';
export const DEFAULT_FONT_SIZE = 10;
export { DEFAULT_GOOD_COLOR, DEFAULT_BAD_COLOR, TONE_MODES };
export const DEFAULT_CURRENCY = "円";
/**
 * 囲みの段ごとの色（外側から）。薄い青の濃淡で、外側ほど濃い。標準テーマの 1 番目のデータの色 #118DFF を地に寄せた色（theme.ts の TINTS）。
 * 塗るのはまとまりを締める行の名前の所だけで、明細の行は地の色
 */
export const DEFAULT_BAND_COLORS = ["#D4EAFF", "#E7F4FF", "#F3F9FF", "#F8FCFF"];
/** 見出しの背景 */
export const DEFAULT_HEAD_BACKGROUND = "#F3F2F1";
/** 罫線の既定の色 */
export const DEFAULT_LINE_COLORS = {
    row: "#C8C6C4",
    subtotal: "#605E5C",
    total: "#252423",
    head: "#252423",
    // セグメントの箱（段ごとの縦線と切れ目の横線）。細い実線
    block: "#8A8886",
    // 囲み（区分・中分類の箱）の線。行の間の点線と同じ濃さ
    band: "#C8C6C4",
    period: "#E1DFDD",
    // 表の外枠
    outer: "#8A8886",
    name: "#C8C6C4",
};
/**
 * 要素ごとの文字サイズの既定（pt）。主の数字は大きく、比較と 2 段目は小さく
 */
// 主の数字は比較の数字の 1.2 倍まで
export const DEFAULT_TEXT_SIZES = { org: 10, name: 10, period: 10, column: 10, compareHead: 10, main: 12, compare: 10, sub: 9 };
/** 比・率の上限の既定（%） */
export const DEFAULT_RATIO_CAP = 999;
export const UNIT_PLACES = { right: "right", corner: "corner", name: "name", cell: "cell" } as const;
const UNIT_PLACE_ITEMS: powerbi.IEnumMember[] = [
    { value: UNIT_PLACES.right, displayName: "表の右上" },
    { value: UNIT_PLACES.corner, displayName: "左上の角（行の名前の上）" },
    { value: UNIT_PLACES.name, displayName: "行の名前の横（売上高（百万円）・台数（台））" },
    { value: UNIT_PLACES.cell, displayName: "数字のあと（1,234百万円）" },
];

/** セグメントの箱の段ごとの塗りの色（外側から）。区分の青と分けて、薄いオレンジの濃淡。外側ほど濃い */
export const DEFAULT_SEGMENT_COLORS = ["#F4B183", "#F8CBAD", "#FBE5D6", "#FDF2EA"];

export const PARENT_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: "below", displayName: "下（明細のあとに小計）" },
    { value: "above", displayName: "上（小計のあとに明細）" },
];

/** 組織の計（子の組織を足したブロック）の位置。区分の小計の位置とは別に選ぶ */
export const ORG_TOTAL_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: "after", displayName: "下（子のセグメントのあとに小計）" },
    { value: "before", displayName: "上（小計のあとに子のセグメント）" },
];

/** 金額の持ち方（メジャーが返す値の符号）。自動はイベントごとに見分ける（signs.ts） */
export const AMOUNT_SIGN_ITEMS: powerbi.IEnumMember[] = [
    { value: "auto", displayName: "自動（シナリオごとに見分ける）" },
    { value: "debit", displayName: "借方プラス（貸方の行がマイナス）" },
    { value: "credit", displayName: "貸方プラス（借方の行がマイナス。足すと利益）" },
    { value: "positive", displayName: "すべてプラス" },
];


export const YEAR_LABEL_ITEMS: powerbi.IEnumMember[] = [
    { value: "start", displayName: "開始年（2025年度）" },
    { value: "end", displayName: "終了年（FY2026）" },
];

/** 年度（表の題）を置く所：表の上（単位と同じ行の左）か、左上の角（期間の見出しと同じ段） */
export const TITLE_PLACES = { top: "top", corner: "corner" } as const;
const TITLE_PLACE_ITEMS: powerbi.IEnumMember[] = [
    { value: TITLE_PLACES.top, displayName: "表の上" },
    { value: TITLE_PLACES.corner, displayName: "左上の角" },
];

/** 行の高さ：セルの上下の余白・字の行の高さ・2 段の行の高さの倍率（visual.less の --ft-row-pad・--ft-row-line・--ft-row-two） */
export const ROW_HEIGHTS = {
    normal: { pad: "0.15em", line: "normal", two: 1.35 },
    compact: { pad: "0.05em", line: "1.25", two: 1.2 },
    tight: { pad: "0em", line: "1.1", two: 1.1 },
} as const;
export type RowHeight = keyof typeof ROW_HEIGHTS;
/** 列の余白：セルの左右の余白（visual.less の --ft-col-pad） */
export const COLUMN_PADDINGS = { normal: "0.5em", compact: "0.3em", tight: "0.15em" } as const;
export type ColumnPadding = keyof typeof COLUMN_PADDINGS;
const ROW_HEIGHT_ITEMS: powerbi.IEnumMember[] = [
    { value: "normal", displayName: "標準" },
    { value: "compact", displayName: "詰める" },
    { value: "tight", displayName: "もっと詰める" },
];

/** 文字の要素ごとの太さ：既定（行の種類で変える）・太字・太字にしない */
export const BOLD_KEYS = ["org", "name", "period", "column", "compareHead", "main", "compare", "sub"] as const;
export type BoldKey = (typeof BOLD_KEYS)[number];
export type BoldMode = "auto" | "on" | "off";
const BOLD_ITEMS: powerbi.IEnumMember[] = [
    { value: "auto", displayName: "既定" },
    { value: "on", displayName: "太字" },
    { value: "off", displayName: "太字にしない" },
];

export const FISCAL_START_ITEMS: powerbi.IEnumMember[] = Array.from({ length: 12 }, (_, i) => ({
    value: String(i + 1),
    displayName: `${i + 1}月`,
}));

/** 前年同期（主と同じイベントの 12 か月前）を比較に選ぶときの値 */
export const PRIOR_YEAR = "__priorYear__";
export const PRIOR_YEAR_LABEL = "前年同期";
/** 最新見込み：組織 × 月ごとに、数字のあるイベントのうち比較順の一番大きいものを採るときの値 */
export const LATEST = "__latest__";
export const LATEST_LABEL = "最新見込み";
/** 年度の「最新」（主に値のある最後の年度） */
export const LATEST_YEAR = "latest";

const NO_ITEM: powerbi.IEnumMember = { value: "", displayName: "（なし）" };

/** ItemDropdown の値の文字列。保存値が items に無いと value は undefined のままになる */
export function dropdownValue(slice: formattingSettings.ItemDropdown, fallback: string): string {
    const raw = slice.value as powerbi.IEnumMember | string | undefined;
    if (raw === undefined || raw === null) return fallback;
    if (typeof raw === "object") return raw.value === undefined || raw.value === null ? fallback : String(raw.value);
    return String(raw);
}

const itemOf = (items: powerbi.IEnumMember[], value: string): powerbi.IEnumMember => items.find((i) => i.value === value) ?? items[0];

/**
 * 比較の列の見せ方（列ごと。見る人が表の上のバーと期間の見出しの ▾ で選ぶ）。1 列の中に何を積むかだけで、1 段は値・差・比・率、2 段は
 * 差と比・差と率。比は 主 ÷ 比較（101.2%）、率は（主 − 比較）÷ |比較|（+1.2%）で同じことの書き方違いなので
 * 重ねない。差と比の両方を見せるときは列を2つ使う
 */
export const COMPARE_VIEWS = {
    value: "value",
    diff: "diff",
    ratio: "ratio",
    rate: "rate",
    diffRatio: "diffRatio",
    diffRate: "diffRate",
} as const;
export type CompareView = (typeof COMPARE_VIEWS)[keyof typeof COMPARE_VIEWS];

export const COMPARE_VIEW_ITEMS: Array<{ value: CompareView; displayName: string }> = [
    { value: COMPARE_VIEWS.value, displayName: "値" },
    { value: COMPARE_VIEWS.diff, displayName: "差" },
    { value: COMPARE_VIEWS.ratio, displayName: "比（101.2%）" },
    { value: COMPARE_VIEWS.rate, displayName: "率（+1.2%）" },
    { value: COMPARE_VIEWS.diffRatio, displayName: "差・比（2 段）" },
    { value: COMPARE_VIEWS.diffRate, displayName: "差・率（2 段）" },
];

/** 比較の列の見せ方の既定 */
export const DEFAULT_COMPARE_VIEW: CompareView = COMPARE_VIEWS.diffRatio;

/**
 * シナリオの列：主、2 段の見せ方、差の色、列の見出しと数字の文字。比較の列は見る人が選ぶので書式ペインに持たない
 */
export class ComparisonCardSettings extends formattingSettings.CompositeCard {
    name = "comparison";
    displayName = "シナリオの列";

    main = new formattingSettings.ItemDropdown({
        name: "main",
        displayName: "基準",
        description:
            "表の値に出すもの。「最新見込み」はセグメント × 月ごとに、数字のあるシナリオのうちシナリオの順序の一番後ろのもの（順序が 2 つ以上のシナリオにあるとき）。順序の無いシナリオ（前年実績など）は基準にしない。見る人は表の上のメニューで変えられる",
        items: [NO_ITEM],
        value: NO_ITEM,
    });

    diffSwap = new formattingSettings.ToggleSwitch({
        name: "diffSwap",
        displayName: "2 段の上下入れ替え",
        description: "差と比・差と率の 2 段で、差を下、比・率を上に出す",
        value: false,
    });

    toneMode = new formattingSettings.ItemDropdown({
        name: "toneMode",
        displayName: "差の色",
        description: "差の列を、良い向き（収益は増えて良い、費用は増えて悪い）で塗る",
        items: DIFF_TONE_MODE_ITEMS,
        value: DIFF_TONE_MODE_ITEMS.find((i) => i.value === TONE_MODES.both)!,
    });
    good = new formattingSettings.ColorPicker({ name: "good", displayName: "良い差", value: { value: DEFAULT_GOOD_COLOR } });
    bad = new formattingSettings.ColorPicker({ name: "bad", displayName: "悪い差", value: { value: DEFAULT_BAD_COLOR } });

    // 主の列と比較の列の見出しは別々に変える。columnSize は主の列の見出し
    columnSize = new formattingSettings.NumUpDown({
        name: "columnSize",
        displayName: "基準列の見出し",
        description: "実績・最新見込みなどの列の見出しと、左上の角に置いた単位",
        value: DEFAULT_TEXT_SIZES.column,
    });
    compareHeadSize = new formattingSettings.NumUpDown({
        name: "compareHeadSize",
        displayName: "比較列の見出し",
        description: "期初予算差・見通し比などの列の見出し（2 行なら 2 行とも）",
        value: DEFAULT_TEXT_SIZES.compareHead,
    });
    columnBold = new formattingSettings.ItemDropdown({
        name: "columnBold",
        displayName: "基準列の見出しの太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    compareHeadBold = new formattingSettings.ItemDropdown({
        name: "compareHeadBold",
        displayName: "比較列の見出しの太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    mainBold = new formattingSettings.ItemDropdown({
        name: "mainBold",
        displayName: "基準列の数字の太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    compareBold = new formattingSettings.ItemDropdown({
        name: "compareBold",
        displayName: "比較列の数字の太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    subBold = new formattingSettings.ItemDropdown({
        name: "subBold",
        displayName: "2 段目の数字の太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    mainSize = new formattingSettings.NumUpDown({ name: "mainSize", displayName: "基準列の数字", value: DEFAULT_TEXT_SIZES.main });
    compareSize = new formattingSettings.NumUpDown({
        name: "compareSize",
        displayName: "比較列の数字",
        description: "比較の列（比較の値・差・比・率）。2 段なら上の段",
        value: DEFAULT_TEXT_SIZES.compare,
    });
    subSize = new formattingSettings.NumUpDown({ name: "subSize", displayName: "2 段目の数字", value: DEFAULT_TEXT_SIZES.sub });

    general = new formattingSettings.Group({ name: "comparisonMain", displayName: "基準", slices: [this.main] });
    /** 2 段の見せ方の上下（比較の列は見る人が選ぶ） */
    twoRows = new formattingSettings.Group({ name: "comparisonTwoRows", displayName: "2 段の表示", slices: [this.diffSwap] });
    tones = new formattingSettings.Group({ name: "comparisonTones", displayName: "差の色", slices: [this.toneMode, this.good, this.bad] });
    texts = new formattingSettings.Group({
        name: "comparisonText",
        displayName: "文字",
        slices: [this.columnSize, this.columnBold, this.compareHeadSize, this.compareHeadBold, this.mainSize, this.mainBold, this.compareSize, this.compareBold, this.subSize, this.subBold],
    });

    groups = [this.general, this.twoRows, this.tones, this.texts];

    /** 主の選択肢を流し込む。main は viewModel が保存値とデータから決めたもの */
    applyEvents(mainItems: powerbi.IEnumMember[], main: string): void {
        this.main.items = mainItems.length > 0 ? mainItems : [NO_ITEM];
        this.main.value = itemOf(this.main.items, main);
    }
}

/**
 * 横持ち（金額にメジャーを 2 本以上）のときの、メジャーごとの比較順と金額の持ち方。保存はメジャーの queryName の selector
 * メジャーの数と名前はデータ次第なので、スライスは applyMeasures で作る。
 * 縦持ちでは出さない（比較順はイベントの表の列。イベントごとの金額の持ち方は「イベントの金額の持ち方」カードの枠）
 */
export class EventsCardSettings extends formattingSettings.CompositeCard {
    name = "events";
    displayName = "シナリオ（メジャーごと）";
    description = "メジャーを積んだときの、メジャーごとのシナリオの順序";
    visible = false;

    orderGroup = new formattingSettings.Group({
        name: "eventsOrder",
        displayName: "シナリオの順序",
        description: "後ろほど確かで、最新見込みで先に採る。「比較のみ」は基準にせず比較にだけ使う（前年実績など）。既定は欄に入れた順（後ろほど確か）",
        slices: [],
    });

    groups = [this.orderGroup];

    applyMeasures(measures: MeasureSetting[]): void {
        this.visible = measures.length > 0;
        this.orderGroup.slices = measures.map((measure) => {
            const items: powerbi.IEnumMember[] = [{ value: ORDER_NONE, displayName: "比較のみ" }, ...measure.orderChoices.map((value) => ({ value, displayName: value }))];
            return new formattingSettings.ItemDropdown({
                name: "order",
                displayName: measure.name,
                items,
                value: itemOf(items, measure.order),
                selector: { metadata: measure.queryName },
            });
        });
    }
}

/** 期間の列：年度、出す列、区切りの線、見出しの文字。月の欄が無い表では年度と出す列を出さない */
export class PeriodsCardSettings extends formattingSettings.CompositeCard {
    name = "periods";
    displayName = "期間の列";

    fiscalYear = new formattingSettings.ItemDropdown({
        name: "fiscalYear",
        displayName: "年度",
        items: [{ value: LATEST_YEAR, displayName: "最新" }],
        value: { value: LATEST_YEAR, displayName: "最新" },
    });

    fiscalStartMonth = new formattingSettings.ItemDropdown({
        name: "fiscalStartMonth",
        displayName: "期首月",
        items: FISCAL_START_ITEMS,
        value: FISCAL_START_ITEMS[3],
    });

    yearLabel = new formattingSettings.ItemDropdown({
        name: "yearLabel",
        displayName: "年度の表記",
        items: YEAR_LABEL_ITEMS,
        value: YEAR_LABEL_ITEMS[0],
    });

    titlePlace = new formattingSettings.ItemDropdown({
        name: "titlePlace",
        displayName: "年度の位置",
        description: "左上の角にすると、行の名前の列の上（期間の見出しと同じ段）に置く",
        items: TITLE_PLACE_ITEMS,
        value: TITLE_PLACE_ITEMS[0],
    });

    showMonths = new formattingSettings.ToggleSwitch({ name: "showMonths", displayName: "月", value: true });
    showQuarters = new formattingSettings.ToggleSwitch({ name: "showQuarters", displayName: "四半期", value: true });
    showHalves = new formattingSettings.ToggleSwitch({ name: "showHalves", displayName: "半期", value: false });
    showYear = new formattingSettings.ToggleSwitch({ name: "showYear", displayName: "通期", value: true });
    showYtd = new formattingSettings.ToggleSwitch({
        name: "showYtd",
        displayName: "累計",
        description: "期首から、シナリオの順序の一番後ろのシナリオ（実績など）に値のある最後の月まで",
        value: false,
    });

    scrollStart = new formattingSettings.ItemDropdown({
        name: "scrollStart",
        displayName: "初期スクロール位置",
        description: "列がはみ出すとき、開いたときに左端（先頭）と右端（末尾）のどちらから見せるか",
        items: SCROLL_START_ITEMS,
        value: SCROLL_START_ITEMS[0],
    });

    periodLines = new formattingSettings.ToggleSwitch({ name: "periodLines", displayName: "期間の区切り", description: "期間のあいだに縦の線を引く", value: true });
    aggregateDouble = new formattingSettings.ToggleSwitch({
        name: "aggregateDouble",
        displayName: "集計の列の区切りを二重線",
        description: "四半期・半期・通期・累計・合計の列の前の区切りを二重線にする（月の列との境目と、集計の列どうしの境目）。「期間の区切り」を入れているときに効く",
        value: false,
    });
    periodLine = new formattingSettings.ColorPicker({ name: "periodLine", displayName: "期間の区切りの色", value: { value: DEFAULT_LINE_COLORS.period } });
    periodBold = new formattingSettings.ItemDropdown({
        name: "periodBold",
        displayName: "期間の見出しの太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    periodSize = new formattingSettings.NumUpDown({ name: "periodSize", displayName: "期間の見出しの文字", value: DEFAULT_TEXT_SIZES.period });

    year = new formattingSettings.Group({ name: "periodsYear", displayName: "年度", slices: [this.fiscalYear, this.fiscalStartMonth, this.yearLabel, this.titlePlace] });
    columns = new formattingSettings.Group({
        name: "periodsColumns",
        displayName: "表示列",
        slices: [this.showMonths, this.showQuarters, this.showHalves, this.showYear, this.showYtd, this.scrollStart],
    });
    look = new formattingSettings.Group({ name: "periodsLook", displayName: "区切りと文字", slices: [this.periodLines, this.aggregateDouble, this.periodLine, this.periodSize, this.periodBold] });

    groups = [this.year, this.columns, this.look];

    /** 年度の選択肢（新しい順）。値は期首の月の「年-月」、先頭は「最新」。月の欄が無ければ年度と出す列を出さない */
    applyYears(years: Array<{ value: string; displayName: string }>, selected: string, hasPeriods = true): void {
        this.fiscalYear.items = [{ value: LATEST_YEAR, displayName: "最新" }, ...years];
        this.fiscalYear.value = itemOf(this.fiscalYear.items, selected);
        this.year.visible = hasPeriods;
        this.columns.visible = hasPeriods;
    }
}

/** 根の合計行（すべての区分を足した行）の出し方 */
export const TOTAL_ROW = { auto: "auto", on: "on", off: "off" } as const;
const TOTAL_ROW_ITEMS: powerbi.IEnumMember[] = [
    { value: TOTAL_ROW.auto, displayName: "自動（行の段が 1 段以下なら出す）" },
    { value: TOTAL_ROW.on, displayName: "表示" },
    { value: TOTAL_ROW.off, displayName: "非表示" },
];

/** 科目の行：並べ方（小計の位置・合計行）、囲み、行の線、行の名前の文字 */
export class RowsCardSettings extends formattingSettings.CompositeCard {
    name = "rows";
    displayName = "行";

    accountStyle = dropdown("accountStyle", "表示", STYLE_ITEMS, "box", "区分・中分類と、その下の行の見せ方。囲みなし（字下げだけ）・囲み（入れ子の箱）・分割（子ごとに表を分ける）・横積み（親の名前を左の列に）");
    accountTotal = dropdown("accountTotal", "展開時の合計", TOTAL_ITEMS, "bottom", "区分・中分類の小計の行を、中身の下・上に置くか、置かないか（なしでも畳むと小計を出す）。うちはいつも親の下");
    accountDirection = dropdown("accountDirection", "分割の向き", DIRECTION_ITEMS, "vertical", "表示が「分割」のとき、子の表を縦に並べるか左右に並べるか");
    stepParents = new formattingSettings.ToggleSwitch({
        name: "stepParents",
        displayName: "計算行を親にする",
        description: "計算行（売上総利益・営業利益…）を、足す範囲の区分の親にする。損益計算書が段階利益の入れ子になり、横積みで親の名前を左に並べられる。切ると計算行は区分と並ぶ 1 行",
        value: false,
    });
    accountRoot = dropdown("accountRoot", "最上位の配置", ROOT_ITEMS, "joined", "一番上の段の行（区分・合計）を 1 つの表につなげるか、表を分けて縦・左右に並べるか（貸借対照表の資産と負債・純資産を左右に）");

    totalName = new formattingSettings.TextInput({
        name: "totalName",
        displayName: "合計行の名前",
        description: "すべての区分を足した合計行と、行の欄を入れない表の 1 行の名前",
        value: "",
        placeholder: "合計",
    });

    totalRow = new formattingSettings.ItemDropdown({
        name: "totalRow",
        displayName: "合計行",
        description:
            "すべての区分を足した合計行。区分が 2 つ以上のときに出せる。見る人はこの行を親にして、区分を「その他」「うち」にまとめられる。行の欄を入れない表で「非表示」にすると、行の名前の列を出さない",
        items: TOTAL_ROW_ITEMS,
        value: TOTAL_ROW_ITEMS[0],
    });

    bandFill = new formattingSettings.ToggleSwitch({
        name: "bandFill",
        displayName: "囲みの塗り",
        description: "区分・中分類の箱の行の名前の所を、段ごとの色で塗る（切ると線だけ）",
        value: true,
    });

    bandColor = new formattingSettings.ColorPicker({
        name: "bandColor",
        displayName: "囲みの色（1 段目）",
        description: "一番外の段（区分）の箱の塗り",
        value: { value: DEFAULT_BAND_COLORS[0] },
    });
    bandColor2 = new formattingSettings.ColorPicker({ name: "bandColor2", displayName: "囲みの色（2 段目）", value: { value: DEFAULT_BAND_COLORS[1] } });
    bandColor3 = new formattingSettings.ColorPicker({ name: "bandColor3", displayName: "囲みの色（3 段目）", value: { value: DEFAULT_BAND_COLORS[2] } });
    bandColor4 = new formattingSettings.ColorPicker({ name: "bandColor4", displayName: "囲みの色（4 段目から）", value: { value: DEFAULT_BAND_COLORS[3] } });
    bandLine = new formattingSettings.ColorPicker({
        name: "bandLine",
        displayName: "囲み線",
        description: "区分・中分類の囲み（箱）の段ごとの縦線。箱の横の辺は行の間の線",
        value: { value: DEFAULT_LINE_COLORS.band },
    });

    rowLine = new formattingSettings.ColorPicker({ name: "rowLine", displayName: "行間の線", value: { value: DEFAULT_LINE_COLORS.row } });
    subtotalLine = new formattingSettings.ColorPicker({ name: "subtotalLine", displayName: "小計の上線", value: { value: DEFAULT_LINE_COLORS.subtotal } });
    totalLine = new formattingSettings.ColorPicker({
        name: "totalLine",
        displayName: "段階利益の二重線",
        description: "売上総利益・営業利益・資産合計など、計算行の小計の上",
        value: { value: DEFAULT_LINE_COLORS.total },
    });

    hideBlankRows = new formattingSettings.ToggleSwitch({
        name: "hideBlankRows",
        displayName: "値の無い行を隠す",
        description: "表に出している列（基準・比較のすべての期間）がどれも空の行を出さない。値が 0 の行は出す",
        value: false,
    });
    breakdownBrackets = new formattingSettings.ToggleSwitch({
        name: "breakdownBrackets",
        displayName: "数字をかっこで囲む",
        description: "うちの行の基準の数字を [ ] で囲み、親の合計に足さない行だと分かるようにする",
        value: true,
    });
    breakdownMuted = new formattingSettings.ToggleSwitch({ name: "breakdownMuted", displayName: "灰色にする", description: "うちの行の名前と数字を薄い字にする", value: true });
    breakdownTag = new formattingSettings.ToggleSwitch({ name: "breakdownTag", displayName: "「うち」の字を付ける", description: "うちの行の名前の前に小さく「うち」と付ける", value: true });
    othersCount = new formattingSettings.ToggleSwitch({
        name: "othersCount",
        displayName: "その他の件数を出す",
        description: "「その他（3件）」のように、まとめた行の数を名前に添える",
        value: true,
    });
    boldAggregates = new formattingSettings.ToggleSwitch({
        name: "boldAggregates",
        displayName: "集計の数字を太字",
        description: "区分・中分類の行の数字も太字にする。切ると、数字の太字は計算行（利益・合計）だけ",
        value: false,
    });
    nameBold = new formattingSettings.ItemDropdown({
        name: "nameBold",
        displayName: "行名の太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    nameSize = new formattingSettings.NumUpDown({ name: "nameSize", displayName: "行名の文字", value: DEFAULT_TEXT_SIZES.name });

    layout = new formattingSettings.Group({
        name: "rowsLayout",
        displayName: "表示",
        description: "段・行ごとに変えるときは「段・行ごとの配置」カード",
        slices: [this.accountStyle, this.accountTotal, this.accountDirection, this.stepParents, this.accountRoot, this.hideBlankRows],
    });
    order = new formattingSettings.Group({ name: "rowsOrder", displayName: "合計行", slices: [this.totalRow, this.totalName] });
    box = new formattingSettings.Group({
        name: "rowsBands",
        displayName: "囲みの塗りと線",
        slices: [this.bandFill, this.bandColor, this.bandColor2, this.bandColor3, this.bandColor4, this.bandLine],
    });
    breakdown = new formattingSettings.Group({
        name: "rowsBreakdown",
        displayName: "うち・その他の行",
        slices: [this.breakdownBrackets, this.breakdownMuted, this.breakdownTag, this.othersCount],
    });
    look = new formattingSettings.Group({ name: "rowsLook", displayName: "行の線と文字", slices: [this.rowLine, this.subtotalLine, this.totalLine, this.boldAggregates, this.nameSize, this.nameBold] });

    groups = [this.layout, this.order, this.box, this.breakdown, this.look];

    /** 段ごとの囲みの色（外側から） */
    bandColors(): string[] {
        return [this.bandColor, this.bandColor2, this.bandColor3, this.bandColor4].map((c, i) => c.value?.value || DEFAULT_BAND_COLORS[i]);
    }
}

/** セグメントの行：小計の位置・合計の名前・データの無い科目、箱の塗りと線、名前の文字。セグメントの欄を入れたときだけ出す */
export class SegmentsCardSettings extends formattingSettings.CompositeCard {
    name = "segments";
    displayName = "セグメント";
    visible = false;

    orgStyle = dropdown("orgStyle", "表示", STYLE_ITEMS, "split", "セグメントと、その下のセグメントの見せ方。分割は子のセグメントごとに表を分ける");
    orgTotal = dropdown("orgTotal", "展開時の合計", TOTAL_ITEMS, "bottom", "子のセグメントを足した合計の表を、子の下・上に置くか、置かないか（なしでも畳むと合計を出す）");
    orgDirection = dropdown("orgDirection", "分割の向き", DIRECTION_ITEMS, "vertical", "表示が「分割」のとき、子の表を縦に並べるか左右に並べるか");

    rootName = new formattingSettings.TextInput({
        name: "rootName",
        displayName: "合計の名前",
        description: "セグメントの一番上の段が 2 つ以上あるとき、それを足した合計のブロックの名前",
        value: "",
        placeholder: "合計",
        visible: false,
    });

    hideEmptyAccounts = new formattingSettings.ToggleSwitch({
        name: "hideEmptyAccounts",
        displayName: "データなしの行を非表示",
        description: "セグメントのブロックで、そのセグメントにデータ（ファクト）の無い行と、下の行が全部隠れた区分・中分類の行を出さない。データがあって値が 0 の行は出す",
        value: true,
    });

    segmentFill = new formattingSettings.ToggleSwitch({
        name: "segmentFill",
        displayName: "セグメントの塗り",
        description: "セグメントの箱の名前の所を、段ごとの色で塗る（切ると線だけ）",
        value: true,
    });
    segmentColor = new formattingSettings.ColorPicker({
        name: "segmentColor",
        displayName: "セグメントの色（1 段目）",
        description: "一番外の段のセグメントの箱の塗り",
        value: { value: DEFAULT_SEGMENT_COLORS[0] },
    });
    segmentColor2 = new formattingSettings.ColorPicker({ name: "segmentColor2", displayName: "セグメントの色（2 段目）", value: { value: DEFAULT_SEGMENT_COLORS[1] } });
    segmentColor3 = new formattingSettings.ColorPicker({ name: "segmentColor3", displayName: "セグメントの色（3 段目）", value: { value: DEFAULT_SEGMENT_COLORS[2] } });
    segmentColor4 = new formattingSettings.ColorPicker({ name: "segmentColor4", displayName: "セグメントの色（4 段目から）", value: { value: DEFAULT_SEGMENT_COLORS[3] } });
    blockLine = new formattingSettings.ColorPicker({
        name: "blockLine",
        displayName: "セグメントの囲み線",
        description: "セグメントの段ごとの縦線と、セグメントの切れ目の横線",
        value: { value: DEFAULT_LINE_COLORS.block },
    });
    orgBold = new formattingSettings.ItemDropdown({
        name: "orgBold",
        displayName: "セグメント名の太字",
        description: "既定は行の種類で太さを変える（計算行・区分の名前は太字など）。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    orgSize = new formattingSettings.NumUpDown({ name: "orgSize", displayName: "セグメント名の文字", value: DEFAULT_TEXT_SIZES.org });

    order = new formattingSettings.Group({
        name: "segmentsOrder",
        displayName: "表示",
        slices: [this.orgStyle, this.orgTotal, this.orgDirection, this.rootName, this.hideEmptyAccounts],
    });
    box = new formattingSettings.Group({
        name: "segmentsBoxes",
        displayName: "箱と文字",
        slices: [this.segmentFill, this.segmentColor, this.segmentColor2, this.segmentColor3, this.segmentColor4, this.blockLine, this.orgSize, this.orgBold],
    });

    groups = [this.order, this.box];

    /** セグメントの欄を入れたときだけ出す。合計の名前は、一番上の段が 2 つ以上で合計を置くときだけ */
    applyOrgs(hasOrgs: boolean, hasRoot: boolean): void {
        this.visible = hasOrgs;
        this.rootName.visible = hasOrgs && hasRoot;
    }

    /** 段ごとのセグメントの箱の色（外側から） */
    segmentColors(): string[] {
        return [this.segmentColor, this.segmentColor2, this.segmentColor3, this.segmentColor4].map((c, i) => c.value?.value || DEFAULT_SEGMENT_COLORS[i]);
    }
}

/** 書式ペインで書ける計算行の数（書式ペインの部品では足し消しできないので決め打ち）。損益と貸借対照表を 1 つの表に入れても足りる数 */
export const CALC_ROW_SLOTS = 20;
export const CALC_KINDS = { none: "none", subtotal: "subtotal", ratio: "ratio" } as const;
const CALC_KIND_ITEMS: powerbi.IEnumMember[] = [
    { value: CALC_KINDS.none, displayName: "使わない" },
    { value: CALC_KINDS.subtotal, displayName: "小計" },
    { value: CALC_KINDS.ratio, displayName: "比率" },
];

export const CALC_GOOD_ITEMS: powerbi.IEnumMember[] = [
    { value: "up", displayName: "上がると良い" },
    { value: "down", displayName: "上がると悪い" },
    { value: "neutral", displayName: "色を付けない" },
];
/** どこからの「表の最初」と、挿入位置の「表の最後」（区分の名前とぶつからない値） */
export const CALC_TABLE_START = "__start__";
export const CALC_TABLE_END = "__end__";
/** 表に無い保存値の表示名：書式ペインの小計は「計算行 N」、区分・指標は名前か queryName（キーの頭を外す）、科目はコード */
const missingLabel = (value: string) => (value.startsWith(CALC_KEY) ? `計算行 ${value.slice(CALC_KEY.length)}` : value.replace(/^§[a-z]+:/, ""));
export const CALC_PROPS = ["kind", "name", "from", "after", "numerator", "denominator", "format", "good"] as const;
export type CalcProp = (typeof CALC_PROPS)[number];
/** 書式ペインの計算行の保存値（生の文字。空は ""） */
export type CalcSaved = Record<CalcProp, string>;
/** 計算行 i（1 から）の設定の名前。1 つの object（calcRows）に kind1・name1… と並べる */
export const calcProp = (prop: CalcProp, slot: number): string => `${prop}${slot}`;
export const EMPTY_CALC: CalcSaved = { kind: "", name: "", from: "", after: "", numerator: "", denominator: "", format: "", good: "" };

/** 計算行の 1 本（書式ペインのコンテナーの項目。公式のサンプルどおり SimpleCard を継ぐ。name は capabilities のカード名でなく識別子） */
class CalcRowItem extends SimpleCard {
    kind: formattingSettings.ItemDropdown;
    rowName: formattingSettings.TextInput;
    from: formattingSettings.ItemDropdown;
    after: formattingSettings.ItemDropdown;
    numerator: formattingSettings.ItemDropdown;
    denominator: formattingSettings.ItemDropdown;
    format: formattingSettings.TextInput;
    good: formattingSettings.ItemDropdown;

    constructor(readonly slot: number) {
        super();
        this.name = `calcRow${slot}`;
        // 表示名は固定。行の名前を出すと、名前や種類を変えるたびに Desktop で編集する行の選択が 1 本目に戻った
        // （2026-09-27、Desktop で確かめた。uid を displayNameKey で固定しても戻った。表示名の変わらない欄の変更では戻らない）
        this.displayName = `計算行 ${slot}`;
        this.kind = new formattingSettings.ItemDropdown({
            name: calcProp("kind", slot),
            displayName: "種類",
            description:
                "小計：開始位置の区分から終了位置の区分までを足す（営業利益・負債合計・フリー CF など）。比率：分子の行 ÷ 分母の行（利益率・原価率・1 人当たり・時間当たり）",
            items: CALC_KIND_ITEMS,
            value: CALC_KIND_ITEMS[0],
        });
        this.rowName = new formattingSettings.TextInput({ name: calcProp("name", slot), displayName: "名前", value: "", placeholder: "営業利益・売上原価率など" });
        this.from = new formattingSettings.ItemDropdown({
            name: calcProp("from", slot),
            displayName: "開始位置",
            description: "小計を足し始める区分（表の並び）。負債合計は流動負債から",
            items: [NO_ITEM],
            value: NO_ITEM,
        });
        this.after = new formattingSettings.ItemDropdown({
            name: calcProp("after", slot),
            displayName: "挿入位置",
            description: "この区分の後に出す。小計はここまで足す",
            items: [NO_ITEM],
            value: NO_ITEM,
        });
        this.numerator = new formattingSettings.ItemDropdown({ name: calcProp("numerator", slot), displayName: "分子", items: [NO_ITEM], value: NO_ITEM });
        this.denominator = new formattingSettings.ItemDropdown({ name: calcProp("denominator", slot), displayName: "分母", items: [NO_ITEM], value: NO_ITEM });
        this.format = new formattingSettings.TextInput({
            name: calcProp("format", slot),
            displayName: "書式",
            description: "空なら 0.0%。#,0 や #,0円 のように書くと数で出す（表示単位で割らない。字を付けるとその字も出す）",
            value: "",
            placeholder: "0.0%・#,0円",
        });
        this.good = new formattingSettings.ItemDropdown({ name: calcProp("good", slot), displayName: "良し悪し", items: CALC_GOOD_ITEMS, value: CALC_GOOD_ITEMS[0] });
        this.slices = [this.kind, this.rowName, this.from, this.after, this.numerator, this.denominator, this.format, this.good];
    }
}

/**
 * 計算行。段階利益・合計・比率を、書式ペインで足す（小計と比率）。区分の名前から型を当てて最初から入れる行は持たない
 *
 */
export class CalcRowsCardSettings extends formattingSettings.CompositeCard {
    name = "calcRows";
    displayName = "計算行";
    description = "段階利益・合計（小計）と比率を足す";

    items = Array.from({ length: CALC_ROW_SLOTS }, (_, i) => new CalcRowItem(i + 1));
    custom = new formattingSettings.Group({
        name: "calcCustom",
        displayName: "追加行",
        description: "小計（開始位置 〜 終了位置の区分を足す。売上総利益・営業利益・資産合計など）と比率（分子 ÷ 分母。利益率・自己資本比率など）。種類を選ぶと要る欄が出る",
        slices: [],
        container: new formattingSettings.Container({ displayName: "編集対象", containerItems: this.items }),
    });

    groups = [this.custom];

    /**
     * 挿入位置・分子・分母の選択肢を流し込む。保存値は viewModel が生で読んだもの
     * （データ次第の選択肢は populate の時点では items に無いので value に入らない）。選択肢に無い保存値は「（表に無い）」で残す
     */
    applyCalc(choices: CalcChoices, saved: CalcSaved[]): void {
        const keep = (items: powerbi.IEnumMember[], value: string) =>
            value === "" || items.some((i) => i.value === value) ? items : [...items, { value, displayName: `${missingLabel(value)}（表に無い）` }];
        const sectionItems = choices.sections.map((c) => ({ value: c.value, displayName: c.displayName }));
        const refItems = choices.refs.map((c) => ({ value: c.value, displayName: c.displayName }));
        this.items.forEach((item, i) => {
            const s = saved[i] ?? EMPTY_CALC;
            // 読めない種類（試作の形など）は使わないとして扱う（行も出さない。viewModel の calcSpecs）
            const kind = s.kind === CALC_KINDS.subtotal || s.kind === CALC_KINDS.ratio ? s.kind : CALC_KINDS.none;
            if (kind === CALC_KINDS.none) item.kind.value = CALC_KIND_ITEMS[0];
            const used = kind !== CALC_KINDS.none;
            const ratio = kind === CALC_KINDS.ratio;
            // 種類で要る欄だけを出す（欄ごとのグレーアウトは部品に無い）：小計は名前・どこから・挿入位置、比率は名前・挿入位置・分子・分母・書式・良し悪し
            item.rowName.visible = used;
            item.from.visible = kind === CALC_KINDS.subtotal;
            item.after.visible = used;
            // 小計の「挿入位置」は足す範囲の最後でもある（その区分まで足して、その後ろに置く）。どこからと組で読めるように名前を変える
            item.after.displayName = ratio ? "挿入位置" : "終了位置";
            item.numerator.visible = ratio;
            item.denominator.visible = ratio;
            item.format.visible = ratio;
            item.good.visible = ratio;
            item.from.items = keep([{ value: CALC_TABLE_START, displayName: "表の最初" }, ...sectionItems], s.from);
            item.from.value = itemOf(item.from.items, s.from || CALC_TABLE_START);
            item.after.items = keep([...sectionItems, { value: CALC_TABLE_END, displayName: "表の最後" }], s.after);
            item.after.value = itemOf(item.after.items, s.after || CALC_TABLE_END);
            // 分子・分母の選択肢は長い（表の行すべて）ので、比率の行にだけ入れる
            item.numerator.items = ratio ? keep([NO_ITEM, ...refItems], s.numerator) : [NO_ITEM];
            item.numerator.value = itemOf(item.numerator.items, s.numerator);
            item.denominator.items = ratio ? keep([NO_ITEM, ...refItems], s.denominator) : [NO_ITEM];
            item.denominator.value = itemOf(item.denominator.items, s.denominator);
        });
    }
}

/** 指標の置き方：挿入位置の行の後ろか、行の「うち」か */
export const INDICATOR_MODES = { after: "after", under: "under", hidden: "hidden" } as const;
const INDICATOR_MODE_ITEMS: powerbi.IEnumMember[] = [
    { value: INDICATOR_MODES.after, displayName: "行の後ろ" },
    { value: INDICATOR_MODES.under, displayName: "行のうち（親の行と同じ見せ方）" },
    { value: INDICATOR_MODES.hidden, displayName: "表に出さない（計算行で使う）" },
];
const INDICATOR_AGGREGATION_ITEMS: powerbi.IEnumMember[] = [
    { value: "flow", displayName: "合計（四半期・通期で足す）" },
    { value: "stock", displayName: "期末（期間の最後の月の値）" },
];
/** 横持ちの指標のイベントの「すべてのイベント」（イベントに依らない値。営業日数など） */
export const INDICATOR_ALL_EVENTS = "__all__";

/** 書式ペインに出す、指標のメジャーごとの設定（viewModel が保存値とデータから決めたもの） */
export interface IndicatorSetting {
    name: string;
    queryName: string;
    /** 行のコード（挿入位置の選択肢から自分を外す） */
    code: string;
    /** 挿入位置（CalcChoices.places の value か CALC_TABLE_END） */
    after: string;
    placement: string;
    aggregation: string;
    format: string;
    good: string;
    /** 横持ちのときのイベント（金額のメジャーの queryName か INDICATOR_ALL_EVENTS）。縦持ちでは null（イベントの列で分かれて届く） */
    event: string | null;
    eventChoices: Array<{ value: string; displayName: string }>;
}

/** 指標の欄の説明（メジャーごとのグループの欄で同じものを使う） */
const INDICATOR_HELP = {
    after: "この行の後ろか、この行のうちに出す。区分は区分の合計行（区分の中の行の後ろ）。既定は表の最後",
    placement:
        "表に出さない：行は出さず、計算行の分子・分母に選べる（率の分子と分母のメジャーを入れて、計算行の比率で率を作る。どの期間でも分子の合計 ÷ 分母の合計になる。人数のような分母も「期間の集計」は合計（延べ人数）にする。期末にすると、四半期の時間を最後の月の人数で割ってしまう）。うち：挿入位置の行の下に 1 段下げて出し、親の行と同じ向き・集計・書式・良し悪しで見せる（値の行のうちは、値の欄と同じ持ち方のメジャー。CALCULATE([値], 製品[区分] = \"新製品\") の形）。親の合計には足さない",
    aggregation: "四半期・通期の値。人数のような残高は期末",
    format: "空なら値の欄と同じ（表示単位で割る。EBITDA など）。#,0人・#,0.0h のように書くと数で出す（表示単位で割らない。字も出す）。0.0% は比率",
    good: "差の色の向き。既定は色を付けない",
    event: "メジャーを積んだ（横持ち）ときの、指標の値のシナリオ。既定はシナリオの順序の一番後ろのメジャー。「すべてのシナリオ」はシナリオに依らない値（営業日数など）",
};

/**
 * 指標。指標の欄のメジャーを、どの行の後ろか・どの行のうちに置くかと、期間の集計・書式・良し悪しをメジャーごとに決める。
 * 保存はメジャーの queryName の selector。メジャーごとのグループに、そのメジャーの欄を並べる（計算行・行別の数値書式と同じく、対象ごとにまとめる）。
 * メジャーごとの項目を入れ物（container）に並べる形は、Desktop で書式ペインの値を変えると、別のメジャー（最初の項目）の selector に
 * 書かれ、ほかのメジャーの保存値が消えた（1.0.0.38）。グループは入れ物でないので、欄ごとの selector のまま書かれる。
 * メジャーの数と名前はデータ次第なので、グループは applyIndicators で作る
 */
export class IndicatorsCardSettings extends formattingSettings.CompositeCard {
    name = "indicators";
    displayName = "指標";
    description = "指標の欄のメジャー（人数・時間・EBITDA・うち）の挿入位置と見せ方（メジャーごと）";
    visible = false;

    groups: formattingSettings.Group[] = [];

    applyIndicators(settings: IndicatorSetting[], places: Array<{ value: string; displayName: string }>): void {
        this.visible = settings.length > 0;
        this.groups = settings.map((setting, i) => {
            const selector = { metadata: setting.queryName };
            const dropdown = (name: string, displayName: string, items: powerbi.IEnumMember[], value: string, description: string) =>
                new formattingSettings.ItemDropdown({ name, displayName, description, items, value: itemOf(items, value), selector });
            // 挿入位置の選択肢は表の行と表の最後（自分の行は入れない）。表に無い保存値（フィルターで行が消えたなど）は選択肢に残す
            const places_: powerbi.IEnumMember[] = [...places.filter((p) => p.value !== setting.code), { value: CALC_TABLE_END, displayName: "表の最後" }];
            if (!places_.some((p) => p.value === setting.after)) places_.push({ value: setting.after, displayName: `${missingLabel(setting.after)}（表に無い）` });
            const hidden = setting.placement === INDICATOR_MODES.hidden;
            const slices: formattingSettings.Slice[] = [
                // 表に出さない指標は置く場所が要らない
                ...(hidden ? [] : [dropdown("after", "挿入位置", places_, setting.after, INDICATOR_HELP.after)]),
                dropdown("placement", "配置", INDICATOR_MODE_ITEMS, setting.placement, INDICATOR_HELP.placement),
            ];
            // うちは親の行に合わせるので、集計・書式・良し悪しの欄を出さない。表に出さない指標は集計だけ（計算行の分子・分母の期間の足し方）
            if (hidden) slices.push(dropdown("aggregation", "期間の集計", INDICATOR_AGGREGATION_ITEMS, setting.aggregation, INDICATOR_HELP.aggregation));
            else if (setting.placement !== INDICATOR_MODES.under) {
                slices.push(
                    dropdown("aggregation", "期間の集計", INDICATOR_AGGREGATION_ITEMS, setting.aggregation, INDICATOR_HELP.aggregation),
                    new formattingSettings.TextInput({ name: "format", displayName: "書式", description: INDICATOR_HELP.format, value: setting.format, placeholder: "#,0人・#,0.0h", selector }),
                    dropdown("good", "良し悪し", CALC_GOOD_ITEMS, setting.good, INDICATOR_HELP.good)
                );
            }
            // 横持ちだけ：指標の値のシナリオ
            if (setting.event !== null) {
                const events: powerbi.IEnumMember[] = [...setting.eventChoices, { value: INDICATOR_ALL_EVENTS, displayName: "すべてのシナリオ" }];
                slices.push(dropdown("event", "シナリオ", events, setting.event, INDICATOR_HELP.event));
            }
            return new formattingSettings.Group({ name: `indicator${i + 1}`, displayName: setting.name, slices });
        });
    }
}

/** 行別の数値書式の枠の数（決め打ち。計算行・イベントの金額の持ち方と同じ形） */
export const ROW_NUMBER_SLOTS = 8;
/** 行別の数値書式の枠で上書きする設定（名前は「数値」カードと同じ。書式は指標の書式と同じ書き方） */
export const ROW_NUMBER_PROPS = ["target", "unitType", "precision", "negativeStyle", "zeroStyle", "diffZeroStyle", "negativeZero", "format", "currency"] as const;
export type RowNumberProp = (typeof ROW_NUMBER_PROPS)[number];
/** 枠 i（1 から）の設定の名前。1 つの object（rowNumbers）に target1・unitType1… と並べる */
export const rowNumberProp = (prop: RowNumberProp, slot: number): string => `${prop}${slot}`;
/** 行別の数値書式の枠の保存値（生の文字。空は ""。表全体に合わせるは ROW_NUMBER_TABLE） */
export type RowNumberSaved = Record<RowNumberProp, string>;
/** 「表全体と同じ」の値 */
export const ROW_NUMBER_TABLE = "table";
const TABLE_ITEM: powerbi.IEnumMember = { value: ROW_NUMBER_TABLE, displayName: "表全体と同じ" };
/** 行で選べる表示単位は固定の単位だけ（自動は表全体の 1 つ。行ごとに自動で変わると縦に比べられない） */
const ROW_UNIT_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...UNIT_TYPES.filter((u) => u.value !== "auto")];
const ROW_PRECISION_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...PRECISIONS.filter((p) => p.value !== "auto")];
const ROW_NEGATIVE_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...NEGATIVE_STYLE_ITEMS];
const ROW_ZERO_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...ZERO_STYLE_ITEMS];
const ROW_NEGATIVE_ZERO_ITEMS: powerbi.IEnumMember[] = [
    TABLE_ITEM,
    { value: "on", displayName: "残す（▲0）" },
    { value: "off", displayName: "残さない（0）" },
];

/** 行別の数値書式の枠の、設定ごとの選択肢の値（「表全体と同じ」を除く）。保存値が選択肢にあるかを見る。字を書く欄は null（どの字でもよい） */
export function rowNumberAllowed(prop: RowNumberProp): string[] | null {
    const items: Partial<Record<RowNumberProp, powerbi.IEnumMember[]>> = {
        unitType: ROW_UNIT_ITEMS,
        precision: ROW_PRECISION_ITEMS,
        negativeStyle: ROW_NEGATIVE_ITEMS,
        zeroStyle: ROW_ZERO_ITEMS,
        diffZeroStyle: ROW_ZERO_ITEMS,
        negativeZero: ROW_NEGATIVE_ZERO_ITEMS,
    };
    const list = items[prop];
    return list ? list.map((i) => String(i.value)).filter((v) => v !== ROW_NUMBER_TABLE) : null;
}

/** 行別の数値書式の枠 1 本（書式ペインのコンテナーの項目） */
class RowNumberItem extends SimpleCard {
    target: formattingSettings.ItemDropdown;
    unitType: formattingSettings.ItemDropdown;
    precision: formattingSettings.ItemDropdown;
    negativeStyle: formattingSettings.ItemDropdown;
    zeroStyle: formattingSettings.ItemDropdown;
    diffZeroStyle: formattingSettings.ItemDropdown;
    negativeZero: formattingSettings.ItemDropdown;
    format: formattingSettings.TextInput;
    currency: formattingSettings.TextInput;

    constructor(readonly slot: number) {
        super();
        this.name = `rowNumber${slot}`;
        // 表示名は固定（計算行と同じ。名前を変えると Desktop で編集する枠の選択が 1 本目に戻った）
        this.displayName = `設定 ${slot}`;
        const dropdown = (prop: RowNumberProp, displayName: string, items: powerbi.IEnumMember[], description?: string) =>
            new formattingSettings.ItemDropdown({ name: rowNumberProp(prop, slot), displayName, description, items, value: items[0] });
        this.target = dropdown("target", "対象", [NO_ITEM], "段（その深さの行すべて）か、行（その行と下の行）。行は段より後に効き、下の行ほど後に効く");
        this.unitType = dropdown("unitType", "表示単位", ROW_UNIT_ITEMS, "表全体と違う単位にすると、親と単位の違う行の名前に単位を添える（「販管費（千円）」）");
        this.precision = dropdown("precision", "小数点以下の桁数", ROW_PRECISION_ITEMS);
        this.negativeStyle = dropdown("negativeStyle", "マイナスの表記", ROW_NEGATIVE_ITEMS);
        this.zeroStyle = dropdown("zeroStyle", "0 の表記", ROW_ZERO_ITEMS);
        this.diffZeroStyle = dropdown("diffZeroStyle", "差 0 の表記", ROW_ZERO_ITEMS);
        this.negativeZero = dropdown("negativeZero", "丸めた 0 の符号", ROW_NEGATIVE_ZERO_ITEMS);
        this.format = new formattingSettings.TextInput({
            name: rowNumberProp("format", slot),
            displayName: "書式",
            description: "#,0h・#,0台 のように書くと、表示単位で割らずに字を付けて出す（時間・台数の行）。0.0% は比率。空は金額",
            value: "",
            placeholder: "#,0h・#,0台",
        });
        this.currency = new formattingSettings.TextInput({
            name: rowNumberProp("currency", slot),
            displayName: "単位文字",
            description: "金額の行の単位に添える字（「千円」の「円」）。空は表全体の単位文字",
            value: "",
            placeholder: "円・ドル",
        });
        this.slices = [this.target, this.unitType, this.precision, this.format, this.currency, this.negativeStyle, this.zeroStyle, this.diffZeroStyle, this.negativeZero];
    }
}

/**
 * 行別の数値書式（「数値」カードの設定を、段か行で上書きする）。対象は行のコード（区分・中分類・科目）か段で覚える（枠の番号や並びでは覚えない）。
 * 効くのは科目の木の行。計算行と後ろに置いた指標の行は表全体のまま（区分をまたぐ行なので）
 */
export class RowNumbersCardSettings extends formattingSettings.CompositeCard {
    name = "rowNumbers";
    displayName = "行別の数値書式";
    description = "「数値」カードの設定を、段か行ごとに上書きする。上書きしない設定は上の段・行か表全体に合わせる";
    visible = false;

    items = Array.from({ length: ROW_NUMBER_SLOTS }, (_, i) => new RowNumberItem(i + 1));
    group = new formattingSettings.Group({
        name: "rowNumbersGroup",
        displayName: "行ごと",
        description: "設定 1〜8 のそれぞれで、対象を 1 つ選び、変えるものを選ぶ。計算行（段階利益・合計・比率）と後ろに置いた指標の行は表全体のまま",
        slices: [],
        container: new formattingSettings.Container({ displayName: "編集対象", containerItems: this.items }),
    });

    groups = [this.group];

    /** 対象の選択肢と保存値を流し込む。対象が無ければ出さない。選択肢に無い保存値は「（表に無い）」で残す */
    applyRowNumbers(targets: Array<{ value: string; displayName: string }>, saved: RowNumberSaved[]): void {
        this.visible = targets.length > 0;
        const base: powerbi.IEnumMember[] = [NO_ITEM, ...targets];
        this.items.forEach((item, i) => {
            const slot = saved[i];
            const value = slot?.target ?? "";
            item.target.items = value === "" || targets.some((t) => t.value === value) ? base : [...base, { value, displayName: `${value}（表に無い）` }];
            item.target.value = itemOf(item.target.items, value);
            const set = (dropdown: formattingSettings.ItemDropdown, v: string | undefined) => (dropdown.value = itemOf(dropdown.items, v || ROW_NUMBER_TABLE));
            set(item.unitType, slot?.unitType);
            set(item.precision, slot?.precision);
            set(item.negativeStyle, slot?.negativeStyle);
            set(item.zeroStyle, slot?.zeroStyle);
            set(item.diffZeroStyle, slot?.diffZeroStyle);
            set(item.negativeZero, slot?.negativeZero);
            item.format.value = slot?.format ?? "";
            item.currency.value = slot?.currency ?? "";
        });
    }
}

export class NumbersCardSettings extends SimpleCard {
    name = "numbers";
    displayName = "数値";

    amountSign = new formattingSettings.ItemDropdown({
        name: "amountSign",
        displayName: "符号の持ち方",
        description:
            "メジャーが返す符号の持ち方。自動は、シナリオごとに貸方の行と借方の行の合計の符号で見分ける。表は区分の向き（区分の中の行の貸方フラグの値の多数）でプラスに見せる。貸方フラグが無ければ届いた値のまま",
        items: AMOUNT_SIGN_ITEMS,
        value: AMOUNT_SIGN_ITEMS[0],
    });

    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        description: "値の欄の行だけを割る（比率・時間・人数の行は割らない）。自動は、一番大きい値が 4 けた以上残る単位を表全体で 1 つ選ぶ",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0],
    });

    unitNotation = new formattingSettings.ItemDropdown({
        name: "unitNotation",
        displayName: "単位の表記",
        items: UNIT_NOTATIONS,
        value: UNIT_NOTATIONS[0],
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        description: "値の欄の行の桁。自動は、単位で割った値が 100 未満なら 1 桁、ほかは 0 桁（表全体でそろえる）",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    currency = new formattingSettings.TextInput({
        name: "currency",
        displayName: "単位文字",
        description: "右上の「単位：百万円」の「円」",
        value: DEFAULT_CURRENCY,
        placeholder: DEFAULT_CURRENCY,
    });

    negativeStyle = new formattingSettings.ItemDropdown({
        name: "negativeStyle",
        displayName: "マイナスの表記",
        items: NEGATIVE_STYLE_ITEMS,
        value: itemOf(NEGATIVE_STYLE_ITEMS, NEGATIVE_STYLES.triangle),
    });

    zeroStyle = new formattingSettings.ItemDropdown({
        name: "zeroStyle",
        displayName: "0 の表記",
        description: "値が 0（丸めて 0 を含む）のときの書き方",
        items: ZERO_STYLE_ITEMS,
        value: itemOf(ZERO_STYLE_ITEMS, ZERO_STYLES.zero),
    });

    diffZeroStyle = new formattingSettings.ItemDropdown({
        name: "diffZeroStyle",
        displayName: "差 0 の表記",
        description: "差・率が 0（変わらない）のときの書き方",
        items: ZERO_STYLE_ITEMS,
        value: itemOf(ZERO_STYLE_ITEMS, ZERO_STYLES.plusMinus),
    });

    negativeZero = new formattingSettings.ToggleSwitch({
        name: "negativeZero",
        displayName: "丸めた 0 の符号",
        description: "▲0 のように、丸めると 0 になるマイナスにも符号を付ける",
        value: true,
    });

    unitPlace = new formattingSettings.ItemDropdown({
        name: "unitPlace",
        displayName: "単位の位置",
        description: "「単位：百万円」を置く所。行の名前の横にすると、値の行は名前に（百万円）、台数・時間のような指標は数字から単位の字を外して名前に（台）を添える（比率の % は数字に残す）",
        items: UNIT_PLACE_ITEMS,
        value: UNIT_PLACE_ITEMS[0],
    });
    alignTails = new formattingSettings.ToggleSwitch({
        name: "alignTails",
        displayName: "単位の字の幅をそろえる",
        description: "数字のあとの単位の字（円・h・% など）の幅を列の中でそろえ、数字の右の端をそろえる。切ると単位の字の無い行の右に空きが出ないが、数字の右の端がずれる",
        value: true,
    });

    // options は付けない（素の numeric に付けると書式ペインが空になった記録がある）。範囲は viewModel でクランプする
    ratioCap = new formattingSettings.NumUpDown({
        name: "ratioCap",
        displayName: "比・率の上限（%）",
        description: "これを超える比・率は「≧999%」のように上限で止める（100〜99999）",
        value: DEFAULT_RATIO_CAP,
    });

    slices = [
        this.amountSign,
        this.unitType,
        this.unitNotation,
        this.precision,
        this.currency,
        this.unitPlace,
        this.alignTails,
        this.negativeStyle,
        this.zeroStyle,
        this.diffZeroStyle,
        this.negativeZero,
        this.ratioCap,
    ];
}

/** 表全体：フォント、表の上の文字、見出しの背景と線、外枠、左の列の区切り、画像のコピー */
export class TableCardSettings extends formattingSettings.CompositeCard {
    name = "table";
    displayName = "表全体";

    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "フォント", value: DEFAULT_FONT_FAMILY });
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "表の上の文字サイズ",
        description: "表の上のバー（題名・メニュー・単位）とダイアログ。表の中の文字は、行・列のカードで要素ごとに変える",
        value: DEFAULT_FONT_SIZE,
    });
    headBackground = new formattingSettings.ColorPicker({
        name: "headBackground",
        displayName: "見出しの背景",
        description: "期間と列の見出し、左上の角、セグメントの列の見出し",
        value: { value: DEFAULT_HEAD_BACKGROUND },
    });
    headLine = new formattingSettings.ColorPicker({ name: "headLine", displayName: "見出しの下線", value: { value: DEFAULT_LINE_COLORS.head } });
    nameLine = new formattingSettings.ColorPicker({
        name: "nameLine",
        displayName: "左列の区切り線",
        description: "セグメントの列と行の名前の列の右の縦線（見出しの左上の角も）",
        value: { value: DEFAULT_LINE_COLORS.name },
    });
    outerLine = new formattingSettings.ColorPicker({ name: "outerLine", displayName: "表の外枠", value: { value: DEFAULT_LINE_COLORS.outer } });
    copyButton = new formattingSettings.ToggleSwitch({
        name: "copyButton",
        displayName: "画像のコピー",
        description: "マウスを乗せたとき、右下に「画像としてコピー」のボタンを出す。押すと、見えている表を画像としてクリップボードに入れ、PowerPoint などに貼れる",
        value: true,
    });

    rowHeight = new formattingSettings.ItemDropdown({
        name: "rowHeight",
        displayName: "行の高さ",
        description: "セルの上下の余白と字の行の高さを詰める",
        items: ROW_HEIGHT_ITEMS,
        value: ROW_HEIGHT_ITEMS[0],
    });
    columnPadding = new formattingSettings.ItemDropdown({
        name: "columnPadding",
        displayName: "列の余白",
        description: "セルの左右の余白を詰める（基準と比較の列のあいだ・表の端）。列の幅は余白を含めて測り直す",
        items: ROW_HEIGHT_ITEMS,
        value: ROW_HEIGHT_ITEMS[0],
    });
    text = new formattingSettings.Group({ name: "tableText", displayName: "文字", slices: [this.fontFamily, this.fontSize, this.rowHeight, this.columnPadding] });
    head = new formattingSettings.Group({ name: "tableHead", displayName: "見出しと枠", slices: [this.headBackground, this.headLine, this.nameLine, this.outerLine] });
    copy = new formattingSettings.Group({ name: "tableCopy", displayName: "画像のコピー", slices: [this.copyButton] });

    groups = [this.text, this.head, this.copy];
}

/** update() のたびにデータから流し込むもの */
export interface DataDrivenFormat {
    hierarchy?: HierarchyFormat;
    mainItems: Array<{ value: string; displayName: string }>;
    main: string;
    /** 横持ちのメジャーごとの比較順と金額の持ち方（縦持ちなら空） */
    measureSettings: MeasureSetting[];
    years: Array<{ value: string; displayName: string }>;
    fiscalYear: string;
    /** 計算行：挿入位置・分子・分母の選択肢、書式ペインの保存値（生） */
    calcChoices: CalcChoices;
    calcSaved: CalcSaved[];
    /** 指標のメジャーごとの設定（挿入位置の選択肢は calcChoices.places） */
    indicatorSettings: IndicatorSetting[];
    /** 行別の数値書式：対象の選択肢（段と行）と、枠の保存値（生） */
    rowNumbers?: { targets: Array<{ value: string; displayName: string }>; saved: RowNumberSaved[] };
    /** 組織の欄を入れたか（組織の設定を出す）と、一番上の段が 2 つ以上で全社を置いたか（全社の名前を出す） */
    hasOrgs?: boolean;
    hasRoot?: boolean;
    /** 月の欄を入れたか（無ければ期間の列のカードの年度と出す列を出さない） */
    hasPeriods?: boolean;
}

export class VisualFormattingSettingsModel extends Model {
    hierarchyOverrides = new HierarchyOverrides();
    comparison = new ComparisonCardSettings();
    events = new EventsCardSettings();
    periods = new PeriodsCardSettings();
    rows = new RowsCardSettings();
    calcRows = new CalcRowsCardSettings();
    indicators = new IndicatorsCardSettings();
    segments = new SegmentsCardSettings();
    numbers = new NumbersCardSettings();
    rowNumbers = new RowNumbersCardSettings();
    table = new TableCardSettings();

    // 対象ごとのカード（2.0）：列（シナリオ・期間）→ 行（科目・セグメント・階層・計算・指標）→ 数値 → 表全体
    cards = [
        this.comparison,
        this.events,
        this.periods,
        this.rows,
        this.segments,
        this.hierarchyOverrides,
        this.calcRows,
        this.indicators,
        this.numbers,
        this.rowNumbers,
        this.table,
    ];

    /**
     * 書式ペインの色のうち、保存値の無いもの（saved が false）の既定をテーマの色にする。書式ペインにもテーマの色が出る。
     * populateFormattingSettingsModel のあと、表を組む前に呼ぶ
     */
    applyTheme(theme: Theme, saved: (object: string, property: string) => boolean): void {
        const set = (object: string, picker: formattingSettings.ColorPicker, value: string) => {
            if (!saved(object, picker.name)) picker.value = { value };
        };
        const { comparison, rows, segments, periods, table } = this;
        set(comparison.name, comparison.good, theme.good);
        set(comparison.name, comparison.bad, theme.bad);
        set(table.name, table.headBackground, theme.headBackground);
        set(rows.name, rows.rowLine, theme.lines.row);
        set(rows.name, rows.subtotalLine, theme.lines.subtotal);
        set(rows.name, rows.totalLine, theme.lines.total);
        set(table.name, table.headLine, theme.lines.head);
        set(segments.name, segments.blockLine, theme.lines.block);
        set(rows.name, rows.bandLine, theme.lines.band);
        set(table.name, table.outerLine, theme.lines.outer);
        set(periods.name, periods.periodLine, theme.lines.period);
        set(table.name, table.nameLine, theme.lines.name);
        [rows.bandColor, rows.bandColor2, rows.bandColor3, rows.bandColor4].forEach((picker, i) => set(rows.name, picker, theme.bandColors[i]));
        [segments.segmentColor, segments.segmentColor2, segments.segmentColor3, segments.segmentColor4].forEach((picker, i) => set(segments.name, picker, theme.segmentColors[i]));
    }

    applyData(data: DataDrivenFormat): void {
        this.hierarchyOverrides.apply(data.hierarchy);
        this.comparison.applyEvents(data.mainItems, data.main);
        this.events.applyMeasures(data.measureSettings ?? []);
        this.periods.applyYears(data.years, data.fiscalYear, data.hasPeriods ?? true);
        this.calcRows.applyCalc(data.calcChoices ?? { sections: [], refs: [], places: [] }, data.calcSaved ?? []);
        this.indicators.applyIndicators(data.indicatorSettings ?? [], data.calcChoices?.places ?? []);
        this.segments.applyOrgs(data.hasOrgs ?? false, data.hasRoot ?? false);
        this.rowNumbers.applyRowNumbers(data.rowNumbers?.targets ?? [], data.rowNumbers?.saved ?? []);
    }
}
