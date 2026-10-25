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
import { DIRECTION_ITEMS, HIERARCHY_SLOTS, HierarchyFormat, HierarchyKind, HierarchyOverrides, ROOT_ITEMS, STYLE_ITEMS, TOTAL_ITEMS, dropdown } from "./hierarchySettings";
import { SegmentRowsCard, SegmentRowsFormat } from "./segmentRows";
import { Reservation, SCOPE_ALL, ScopeChoice, resolveScope } from "./scope";

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
/** 表のすぐ上の年度と単位の文字サイズの既定（pt） */
export const DEFAULT_HEADING_SIZE = 9;
export { DEFAULT_GOOD_COLOR, DEFAULT_BAD_COLOR, TONE_MODES };
export const DEFAULT_CURRENCY = "円";
/**
 * 囲みの段ごとの色（外側から）。灰の濃淡で、外側ほど濃い。見出しの地 #F5F5F5 から濃い端 #C9C8C8（どちらも標準の字 #252423 と白から作る段。theme.ts）へ寄せた色（BAND_STEPS）。
 * 塗るのはまとまりを締める行の名前の所だけで、明細の行は地の色
 */
export const DEFAULT_BAND_COLORS = ["#E1E1E1", "#E8E8E8", "#EEEEEE", "#F5F5F5"];
/** 見出しの背景（白から標準の字へ 4.5% 寄せた、ごく薄い灰。標準の表とマトリックスは見出しを塗らない） */
export const DEFAULT_HEAD_BACKGROUND = "#F5F5F5";
/** 罫線の既定の色 */
export const DEFAULT_LINE_COLORS = {
    // 2026-10-08 ユーザーが Desktop で決めた色：行の中の線は薄い灰（テーマの白を 10% 暗く）、セグメント・期間の区切り・外枠は文字の色
    row: "#E6E6E6",
    subtotal: "#E6E6E6",
    total: "#E6E6E6",
    head: "#252423",
    // セグメントの箱（段ごとの縦線と切れ目の横線）。細い実線
    block: "#252423",
    // 囲み（区分・中分類の箱）の線。行の間の線と同じ濃さ
    band: "#E6E6E6",
    period: "#252423",
    // 表の外枠
    outer: "#252423",
    name: "#E6E6E6",
};
/**
 * 要素ごとの文字サイズの既定（pt）。主の数字は大きく、比較と 2 段目は小さく
 */
// 2026-10-08 ユーザーが旧 managementFigures に似せて Desktop で決めた大きさ（主の数字は大きく、名前・見出しは小さく）
export const DEFAULT_TEXT_SIZES = { org: 9, name: 9, period: 10, column: 9, compareHead: 8, main: 15, compare: 9, sub: 9 };
/** 比・率の上限の既定（%） */
export const DEFAULT_RATIO_CAP = 999;
export const UNIT_PLACES = { right: "right", corner: "corner", name: "name", cell: "cell" } as const;
const UNIT_PLACE_ITEMS: powerbi.IEnumMember[] = [
    { value: UNIT_PLACES.right, displayName: "表の右上" },
    { value: UNIT_PLACES.corner, displayName: "左上の角（行見出しの列の上）" },
    { value: UNIT_PLACES.name, displayName: "行見出しの横（売上高（百万円）・台数（台））" },
    { value: UNIT_PLACES.cell, displayName: "数字のあと（1,234百万円）" },
];

/** セグメントの箱の段ごとの塗りの色（外側から）。囲みよりひと段濃い灰の濃淡で、外側ほど濃い（theme.ts の SEGMENT_STEPS） */
export const DEFAULT_SEGMENT_COLORS = ["#D4D3D3", "#DDDCDC", "#E6E5E5", "#EEEEEE"];

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

/** 文字の要素ごとの太さ：自動（行の種類で変える）・太字・太字にしない */
export const BOLD_KEYS = ["org", "name", "period", "column", "compareHead", "main", "compare", "sub"] as const;
export type BoldKey = (typeof BOLD_KEYS)[number];
export type BoldMode = "auto" | "on" | "off";
const BOLD_ITEMS: powerbi.IEnumMember[] = [
    { value: "auto", displayName: "自動" },
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
export const LATEST_LABEL = "最新値";
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
/**
 * シナリオの列の見出しの書き換え。既定は、列で使ったシナリオのうち一番弱い（順序が前の）名前。書き換えると、指定したシナリオだけの列と
 * ほかが入る列を、それぞれ決めた見出しにする（2026-10-16 ユーザー決定。保存値 label〔「最新見込み」と名乗る〕は書き換えなしと同じ）
 */
export const LATEST_HEADERS = { label: "label", scenario: "scenario", custom: "custom" } as const;
const LATEST_HEADER_ITEMS: powerbi.IEnumMember[] = [
    { value: LATEST_HEADERS.scenario, displayName: "書き換えない" },
    { value: LATEST_HEADERS.custom, displayName: "指定したシナリオだけ" },
];

export class ComparisonCardSettings extends formattingSettings.CompositeCard {
    name = "comparison";
    displayName = "シナリオの列";
    description = "シナリオは「シナリオ」フィールド（実績・予算・見通しなど）の値。ここでは基準・見出し・差の色・文字を決める。比較の列は見る人が表の上のバーで選ぶので、書式ペインには無い";

    main = new formattingSettings.ItemDropdown({
        name: "main",
        displayName: "基準",
        description:
            "表の値に出すもの。「最新値」はセグメント × 月ごとに、数字のあるシナリオのうちシナリオの順序の一番後ろのもの（順序が 2 つ以上のシナリオにあるとき）。順序の無いシナリオ（前年の数字を入れたシナリオなど）は基準にしない。見る人は表の上のメニューで変えられる",
        items: [NO_ITEM],
        value: NO_ITEM,
    });

    latestHeader = new formattingSettings.ItemDropdown({
        name: "latestHeader",
        displayName: "見出しの書き換え",
        description:
            "最新値の列の見出し。書き換えないときは、列で使ったシナリオのうち一番弱い（シナリオの順序が前の）名前（4〜8 月は実績で 9 月からは見通しの通期は「見通し」）。書き換えると、指定したシナリオだけの列は 1 つ目の見出し、ほかのシナリオが入る列は 2 つ目の見出しにする",
        items: LATEST_HEADER_ITEMS,
        value: LATEST_HEADER_ITEMS[0],
    });
    confirmedScenario = new formattingSettings.TextInput({
        name: "confirmedScenario",
        displayName: "指定するシナリオ",
        description: "見出しを書き換えるときに見るシナリオの名前。いくつもあれば「,」で区切る",
        value: "",
        placeholder: "例: 実績",
    });
    confirmedLabel = new formattingSettings.TextInput({
        name: "confirmedLabel",
        displayName: "指定したシナリオだけのときの見出し",
        description: "列の数字が指定したシナリオだけのときの見出し。空ならそのシナリオの名前",
        value: "",
        placeholder: "例: 実績",
    });
    forecastLabel = new formattingSettings.TextInput({
        name: "forecastLabel",
        displayName: "ほかのシナリオが入るときの見出し",
        description: "列の数字に、指定したシナリオのほかが入るときの見出し。空なら一番弱いシナリオの名前",
        value: "",
        placeholder: "例: 見通し",
    });
    relabelMonths = new formattingSettings.ToggleSwitch({
        name: "relabelMonths",
        displayName: "月の列も書き換える",
        description: "切ると、書き換えるのは四半期・半期・通期・累計の列だけ。月の列は一番弱いシナリオの名前のまま",
        value: true,
    });

    diffSwap = new formattingSettings.ToggleSwitch({
        name: "diffSwap",
        displayName: "2 段の上下入れ替え",
        description: "差と比・差と率の 2 段で、差を下、比・率を上に出す",
        value: false,
    });

    toneMode = new formattingSettings.ItemDropdown({
        name: "toneMode",
        displayName: "塗り方",
        description: "差の列を良し悪し（収益は増えて良い、費用は増えて悪い）で塗る。行ごとの良し悪しは「数値」の適用先で行を選んで変える",
        items: DIFF_TONE_MODE_ITEMS,
        value: DIFF_TONE_MODE_ITEMS.find((i) => i.value === TONE_MODES.both)!,
    });
    good = new formattingSettings.ColorPicker({ name: "good", displayName: "良い差の色", description: "良し悪しで良い差の数字の色", value: { value: DEFAULT_GOOD_COLOR } });
    bad = new formattingSettings.ColorPicker({ name: "bad", displayName: "悪い差の色", description: "良し悪しで悪い差の数字の色", value: { value: DEFAULT_BAD_COLOR } });

    // 主の列と比較の列の見出しは別々に変える。columnSize は主の列の見出し
    columnSize = new formattingSettings.NumUpDown({
        name: "columnSize",
        displayName: "基準列の見出しの文字サイズ",
        description: "実績・最新値などの列の見出しと、左上の角に置いた単位",
        value: DEFAULT_TEXT_SIZES.column,
    });
    compareHeadSize = new formattingSettings.NumUpDown({
        name: "compareHeadSize",
        displayName: "比較列の見出しの文字サイズ",
        description: "期初予算差・見通し比などの列の見出し（2 段の見出しなら 2 段とも）",
        value: DEFAULT_TEXT_SIZES.compareHead,
    });
    columnBold = new formattingSettings.ItemDropdown({
        name: "columnBold",
        displayName: "基準列の見出しの太字",
        description: "自動はビジュアルの標準の太さ。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    compareHeadBold = new formattingSettings.ItemDropdown({
        name: "compareHeadBold",
        displayName: "比較列の見出しの太字",
        description: "自動はビジュアルの標準の太さ。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    mainSize = new formattingSettings.NumUpDown({ name: "mainSize", displayName: "基準列の数字の文字サイズ", description: "基準の列（実績・最新値など）の数字", value: DEFAULT_TEXT_SIZES.main });
    compareSize = new formattingSettings.NumUpDown({
        name: "compareSize",
        displayName: "比較列の数字の文字サイズ",
        description: "比較の列（比較の値・差・比・率）。2 段なら上の段",
        value: DEFAULT_TEXT_SIZES.compare,
    });
    subSize = new formattingSettings.NumUpDown({ name: "subSize", displayName: "2 段目の数字の文字サイズ", description: "2 段の見せ方の 2 段目の数字", value: DEFAULT_TEXT_SIZES.sub });

    general = new formattingSettings.Group({ name: "comparisonMain", displayName: "基準", slices: [this.main, this.latestHeader, this.confirmedScenario, this.confirmedLabel, this.forecastLabel, this.relabelMonths] });
    /** 2 段の見せ方の上下（比較の列は見る人が選ぶ） */
    twoRows = new formattingSettings.Group({ name: "comparisonTwoRows", displayName: "2 段の表示", slices: [this.diffSwap] });
    tones = new formattingSettings.Group({ name: "comparisonTones", displayName: "差の色", slices: [this.toneMode, this.good, this.bad] });
    texts = new formattingSettings.Group({
        name: "comparisonText",
        displayName: "文字",
        slices: [this.columnSize, this.columnBold, this.compareHeadSize, this.compareHeadBold, this.mainSize, this.compareSize, this.subSize],
    });

    groups = [this.general, this.twoRows, this.tones, this.texts];

    /** 主の選択肢を流し込む。main は viewModel が保存値とデータから決めたもの */
    applyEvents(mainItems: powerbi.IEnumMember[], main: string): void {
        this.main.items = mainItems.length > 0 ? mainItems : [NO_ITEM];
        this.main.value = itemOf(this.main.items, main);
        // 保存値 label（「最新見込み」と名乗る。2.1.7.14 でやめた）は書き換えないと同じ
        if (String(this.latestHeader.value?.value) !== LATEST_HEADERS.custom) this.latestHeader.value = LATEST_HEADER_ITEMS[0];
        // 文字の欄とトグルは書き換えるときだけ
        const custom = String(this.latestHeader.value.value) === LATEST_HEADERS.custom;
        this.confirmedScenario.visible = custom;
        this.confirmedLabel.visible = custom;
        this.forecastLabel.visible = custom;
        this.relabelMonths.visible = custom;
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
        description: "後ろほど確かで、最新値で先に採る。「比較のみ」は基準にせず比較にだけ使う（前年の数字を入れたシナリオなど）。既定は「値」フィールドに入れた順（後ろほど確か）",
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

/** 期間の列：年度、出す列、スクロール、区切りの線、見出しの文字。「月」フィールドが無い表では年度と出す列とスクロールを出さない */
export class PeriodsCardSettings extends formattingSettings.CompositeCard {
    name = "periods";
    displayName = "期間の列";

    fiscalYear = new formattingSettings.ItemDropdown({
        name: "fiscalYear",
        displayName: "表示する年度",
        description: "表に出す年度。「最新」は基準に値のある最後の年度",
        items: [{ value: LATEST_YEAR, displayName: "最新" }],
        value: { value: LATEST_YEAR, displayName: "最新" },
    });

    fiscalStartMonth = new formattingSettings.ItemDropdown({
        name: "fiscalStartMonth",
        displayName: "期首月",
        description: "年度の始まりの月。四半期・半期・通期・累計はこの月から数える",
        items: FISCAL_START_ITEMS,
        value: FISCAL_START_ITEMS[3],
    });


    showMonths = new formattingSettings.ToggleSwitch({ name: "showMonths", displayName: "月", description: "入れた種類の期間の列を表に出す", value: true });
    showQuarters = new formattingSettings.ToggleSwitch({ name: "showQuarters", displayName: "四半期", description: "入れた種類の期間の列を表に出す", value: true });
    showHalves = new formattingSettings.ToggleSwitch({ name: "showHalves", displayName: "半期", description: "入れた種類の期間の列を表に出す", value: false });
    showYear = new formattingSettings.ToggleSwitch({ name: "showYear", displayName: "通期", description: "入れた種類の期間の列を表に出す", value: true });
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
        displayName: "期間の種類の境を二重線",
        description: "期間の種類が変わる所（月と四半期、四半期と半期、半期と通期、累計・合計の前）の区切りを二重線にする。同じ種類の列どうし（1Q と 2Q など）は 1 本の線。「期間の区切り」を入れているときに効く",
        value: false,
    });
    periodLine = new formattingSettings.ColorPicker({ name: "periodLine", displayName: "期間の区切りの色", description: "期間のあいだの縦線と二重線の色", value: { value: DEFAULT_LINE_COLORS.period } });
    periodBold = new formattingSettings.ItemDropdown({
        name: "periodBold",
        displayName: "期間の見出しの太字",
        description: "自動はビジュアルの標準の太さ。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    periodSize = new formattingSettings.NumUpDown({ name: "periodSize", displayName: "期間の見出しの文字サイズ", description: "月・四半期などの見出しの字（左上の角に置いた年度も）", value: DEFAULT_TEXT_SIZES.period });

    year = new formattingSettings.Group({
        name: "periodsYear",
        displayName: "年度",
        description: "年度の数え方と年度の位置は「表全体 › 表の上」、年度の書き方は「用語」",
        slices: [this.fiscalYear, this.fiscalStartMonth],
    });
    columns = new formattingSettings.Group({
        name: "periodsColumns",
        displayName: "表示列",
        slices: [this.showMonths, this.showQuarters, this.showHalves, this.showYear, this.showYtd],
    });
    scroll = new formattingSettings.Group({ name: "periodsScroll", displayName: "スクロール", slices: [this.scrollStart] });
    look = new formattingSettings.Group({ name: "periodsLook", displayName: "区切りと文字", slices: [this.periodLines, this.aggregateDouble, this.periodLine, this.periodSize, this.periodBold] });

    groups = [this.year, this.columns, this.scroll, this.look];

    /** 年度の選択肢（新しい順）。値は期首の月の「年-月」、先頭は「最新」。「月」フィールドが無ければ年度と出す列とスクロールを出さない */
    applyYears(years: Array<{ value: string; displayName: string }>, selected: string, hasPeriods = true): void {
        this.fiscalYear.items = [{ value: LATEST_YEAR, displayName: "最新" }, ...years];
        this.fiscalYear.value = itemOf(this.fiscalYear.items, selected);
        this.year.visible = hasPeriods;
        this.columns.visible = hasPeriods;
        this.scroll.visible = hasPeriods;
    }
}

/** セグメントの合計のブロック（レベル 1 のセグメントが 2 つ以上のとき）の出し方 */
export const ROOT_TOTAL = { on: "on", off: "off" } as const;
const ROOT_TOTAL_ITEMS: powerbi.IEnumMember[] = [
    { value: ROOT_TOTAL.on, displayName: "表示" },
    { value: ROOT_TOTAL.off, displayName: "非表示" },
];

/** 根の合計行（すべての区分を足した行）の出し方 */
export const TOTAL_ROW = { auto: "auto", on: "on", off: "off" } as const;
const TOTAL_ROW_ITEMS: powerbi.IEnumMember[] = [
    { value: TOTAL_ROW.auto, displayName: "自動（行のレベルが 1 つ以下なら出す）" },
    { value: TOTAL_ROW.on, displayName: "表示" },
    { value: TOTAL_ROW.off, displayName: "非表示" },
];

/** 科目の行：並べ方（レイアウト・小計の位置）、合計行、太字、囲み、うち・その他、行見出しと線 */
/** 適用先の「すべて」の名前（「数値」「行」は すべての行、「セグメント」は すべてのセグメント） */
const ALL_ROWS_NAME = "すべての行";
const ALL_SEGMENTS_NAME = "すべてのセグメント";

export class RowsCardSettings extends formattingSettings.CompositeCard {
    name = "rows";
    displayName = "行";

    accountStyle = dropdown("accountStyle", "レイアウト", STYLE_ITEMS, "box", "小計の行（区分・中分類など）と、その下の行の見せ方。囲みなし（字下げだけ）・囲み（入れ子の囲み）・分割（子ごとに表を分ける）・横積み（小計の行の行見出しを左の列に）");
    accountTotal = dropdown("accountTotal", "小計の位置", TOTAL_ITEMS, "bottom", "小計の行（区分・中分類など）を、子の下・上に置くか、置かないか（なしでも畳むと小計を出す）。うちはいつも親の下");
    accountDirection = dropdown("accountDirection", "分割の向き", DIRECTION_ITEMS, "vertical", "レイアウトが「分割」のとき、子の表を縦に並べるか左右に並べるか");
    /** 書式ペインでは「計算行」カードの「並べ方」に出す（arrangeCards）。保存先は rows のまま */
    stepParents = new formattingSettings.ToggleSwitch({
        name: "stepParents",
        displayName: "計算行を親にする",
        description: "計算行（売上総利益・営業利益…）を、足す範囲のレベル 1 の行の親にする。損益計算書が段階利益の入れ子になり、横積みで親の行見出しを左に並べられる。切ると計算行はレベル 1 の行と並ぶ 1 行",
        value: false,
    });
    accountRoot = dropdown(
        "accountRoot",
        "レベル 1 の配置",
        ROOT_ITEMS,
        "joined",
        "レベル 1 の行（区分・合計など）どうしを 1 つの表につなげるか、表を分けて縦・左右に並べるか（貸借対照表の資産と負債・純資産を左右に）。レベル 1 の中の子の分け方はレイアウトで決める",
    );

    rowHeader = new formattingSettings.TextInput({
        name: "rowHeader",
        displayName: "行見出しの列の見出し",
        description: "行見出しの列の、左上の見出しの文字。空なら「項目」",
        value: "",
        placeholder: "項目",
    });
    rowHeaderBold = new formattingSettings.ToggleSwitch({ name: "rowHeaderBold", displayName: "行見出しの列の見出しを太字", description: "左上の角の「項目」の字を太字にする", value: false });

    totalName = new formattingSettings.TextInput({
        name: "totalName",
        displayName: "名前",
        description: "レベル 1 の行をすべて足した合計行と、「行」フィールドを入れない表の 1 行の名前",
        value: "",
        placeholder: "合計",
    });

    totalRow = new formattingSettings.ItemDropdown({
        name: "totalRow",
        displayName: "表示",
        description:
            "レベル 1 の行をすべて足した合計行。レベル 1 の行が 2 つ以上のときに出せる。見る人はこの行を親にして、レベル 1 の行を「その他」「うち」にまとめられる。「行」フィールドを入れない表で「非表示」にすると、行見出しの列を出さない",
        items: TOTAL_ROW_ITEMS,
        value: TOTAL_ROW_ITEMS[0],
    });

    bandFill = new formattingSettings.ToggleSwitch({
        name: "bandFill",
        displayName: "囲みの塗り",
        description: "小計の行（区分・中分類など）の囲みの行見出しを、レベルごとの色で塗る（切ると線だけ）。既定は塗らない",
        value: false,
    });

    bandColor = new formattingSettings.ColorPicker({
        name: "bandColor",
        displayName: "囲みの色（レベル 1）",
        description: "レベル 1 の小計の行の行見出しの塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_BAND_COLORS[0] },
    });
    bandColor2 = new formattingSettings.ColorPicker({
        name: "bandColor2",
        displayName: "囲みの色（レベル 2）",
        description: "そのレベルの小計の行の行見出しの塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_BAND_COLORS[1] },
    });
    bandColor3 = new formattingSettings.ColorPicker({
        name: "bandColor3",
        displayName: "囲みの色（レベル 3）",
        description: "そのレベルの小計の行の行見出しの塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_BAND_COLORS[2] },
    });
    bandColor4 = new formattingSettings.ColorPicker({
        name: "bandColor4",
        displayName: "囲みの色（レベル 4 から）",
        description: "そのレベルの小計の行の行見出しの塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_BAND_COLORS[3] },
    });
    bandLine = new formattingSettings.ColorPicker({
        name: "bandLine",
        displayName: "囲み線の色",
        description: "小計の行の囲みの、レベルごとの縦線。囲みの横の辺は「行間の線の色」",
        value: { value: DEFAULT_LINE_COLORS.band },
    });

    rowLine = new formattingSettings.ColorPicker({ name: "rowLine", displayName: "行間の線の色", description: "行と行のあいだの横線と、囲みの横の辺", value: { value: DEFAULT_LINE_COLORS.row } });
    subtotalLine = new formattingSettings.ColorPicker({ name: "subtotalLine", displayName: "小計の線の色", description: "小計の行の上の線", value: { value: DEFAULT_LINE_COLORS.subtotal } });
    totalLine = new formattingSettings.ColorPicker({
        name: "totalLine",
        displayName: "計算行の二重線の色",
        description: "売上総利益・営業利益・資産合計など、計算行の小計の上",
        value: { value: DEFAULT_LINE_COLORS.total },
    });

    hideBlankRows = new formattingSettings.ToggleSwitch({
        name: "hideBlankRows",
        displayName: "値の無い行を隠す",
        description: "表に出している列（基準・比較のすべての期間）がどれも空の行を出さない（科目と指標の行）。値が 0 の行は出す",
        value: false,
    });
    hideZeroRows = new formattingSettings.ToggleSwitch({
        name: "hideZeroRows",
        displayName: "0 だけの行を隠す",
        description: "表に出している列（基準・比較のすべての期間）の月の値がどれも 0 か空の行を出さない（科目と指標の行）。空の行も隠すので、入れると「値の無い行を隠す」は要らない",
        value: false,
    });
    breakdownBrackets = new formattingSettings.ToggleSwitch({
        name: "breakdownBrackets",
        displayName: "数字を [ ] で囲む",
        description: "うちの行の基準の数字を [ ] で囲み、親の合計に足さない行だと分かるようにする",
        value: true,
    });
    breakdownMuted = new formattingSettings.ToggleSwitch({ name: "breakdownMuted", displayName: "薄い文字の色にする", description: "うちの行の行見出しと数字を、表全体の「薄い文字の色」にする", value: true });
    breakdownTag = new formattingSettings.ToggleSwitch({ name: "breakdownTag", displayName: "「うち」の字を付ける", description: "うちの行の行見出しの前に小さく「うち」と付ける", value: true });
    othersCount = new formattingSettings.ToggleSwitch({
        name: "othersCount",
        displayName: "その他の件数を出す",
        description: "「その他（3件）」のように、まとめた行の数を行見出しに添える",
        value: true,
    });
    boldAggregates = new formattingSettings.ToggleSwitch({
        name: "boldAggregates",
        displayName: "小計の行の数字も太字",
        description: "小計の行（区分・中分類など）の数字も太字にする。切ると、数字の太字は計算行（利益・合計）だけ",
        value: false,
    });
    nameSize = new formattingSettings.NumUpDown({ name: "nameSize", displayName: "行見出しの文字サイズ", description: "行見出しの列の字（セグメントの名前は「セグメント」カード）", value: DEFAULT_TEXT_SIZES.name });
    allNamesBold = new formattingSettings.ToggleSwitch({
        name: "allNamesBold",
        displayName: "行見出しをすべて太字",
        description: "行見出しを、行の種類によらず太字にする",
        value: false,
    });

    /** 適用先（「数値」と同じ形。2026-10-17 ユーザー）。すべての行なら下は表全体の設定、レベル・行を選ぶと「配置」だけをその対象で書き換える（arrangeCards） */
    scope = new formattingSettings.ItemDropdown({
        name: "scope",
        displayName: "対象",
        description:
            "すべての行は表全体の設定。レベル・行を選ぶと、その対象のレイアウト・小計の位置・分割の向きを変える。レベルを選んだ設定はすべての行の設定より強く、行を選んだ設定はレベルより強い（同じ対象が 2 つあれば後の枠）",
        items: [{ value: SCOPE_ALL, displayName: ALL_ROWS_NAME }],
        value: { value: SCOPE_ALL, displayName: ALL_ROWS_NAME },
        visible: false,
    });
    scopeSlotNumber: number | null = null;
    scopeFull = false;
    scopeGroup = new formattingSettings.Group({ name: "rowsScope", displayName: "設定の適用先", slices: [this.scope] });

    layout = new formattingSettings.Group({
        name: "rowsLayout",
        displayName: "並べ方",
        description: "「計算行を親にする」は「計算行」カード",
        slices: [this.accountStyle, this.accountTotal, this.accountDirection, this.stepParents, this.accountRoot, this.hideBlankRows, this.hideZeroRows],
    });
    /** 合計・計算行・小計を太字にするか（行の名前と数字。2026-10-08、太字は合計行にするかしないかだけで足りるとユーザーと決めた。要素ごとの太字の選択をやめてここにまとめた） */
    totalBold = new formattingSettings.ToggleSwitch({
        name: "totalBold",
        displayName: "集計行を太字",
        description: "集計行（合計・計算行・小計の行）の行見出しと、計算行の数字を太字にする。切ると、どの行も太字にしない",
        value: true,
    });
    /** 表の数字をすべて太字（2026-10-08。2.1.7.8 で「シナリオの列」カードから移した。保存値の読み替えは VisualFormattingSettingsModel.applyMoved） */
    valueBold = new formattingSettings.ToggleSwitch({
        name: "valueBold",
        displayName: "数字をすべて太字",
        description: "基準・比較・2 段目の数字を、行の種類によらず太字にする",
        value: false,
    });
    order = new formattingSettings.Group({ name: "rowsOrder", displayName: "合計行", slices: [this.totalRow, this.totalName] });
    // 太字はカードを対象ごとに分ける方針の例外で、表の中の行・数字の太字をここに集める（見出しの太字は各見出しの文字サイズの隣）。
    // 効く順は描画（visual.less の ft-row-bold・ft-row-plain・ft-bold-*）のとおり
    bold = new formattingSettings.Group({
        name: "rowsBold",
        displayName: "太字",
        description:
            "強い順に、「数値」の適用先で計算行・指標の行に決めた太字（太字・太字にしない）→ 行見出しをすべて太字・数字をすべて太字 → 集計行を太字（切るとどの行も太字にせず、小計の行の数字も太字も効かない）→ 小計の行の数字も太字",
        slices: [this.totalBold, this.boldAggregates, this.allNamesBold, this.valueBold],
    });
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
    look = new formattingSettings.Group({
        name: "rowsLook",
        displayName: "行見出しと線",
        slices: [this.rowHeader, this.rowHeaderBold, this.nameSize, this.rowLine, this.subtotalLine, this.totalLine],
    });

    groups = [this.scopeGroup, this.layout, this.order, this.bold, this.box, this.breakdown, this.look];

    /** レベルごとの囲みの色（外側から） */
    bandColors(): string[] {
        return [this.bandColor, this.bandColor2, this.bandColor3, this.bandColor4].map((c, i) => c.value?.value || DEFAULT_BAND_COLORS[i]);
    }
}

/** セグメントの行：並べ方（レイアウト・小計の位置・データの無い行）、合計と小計の名前、囲みの塗りと線、文字。「セグメント」フィールドを入れたときだけ出す */
export class SegmentsCardSettings extends formattingSettings.CompositeCard {
    name = "segments";
    displayName = "セグメント";
    visible = false;

    orgStyle = dropdown("orgStyle", "レイアウト", STYLE_ITEMS, "box", "セグメントと、その下のセグメントの見せ方。分割は子のセグメントごとに表を分ける。既定は囲み");
    orgTotal = dropdown("orgTotal", "小計の位置", TOTAL_ITEMS, "bottom", "子のセグメントを足した小計のブロックを、子の下・上に置くか、置かないか（なしでも畳むと小計を出す）");
    orgDirection = dropdown("orgDirection", "分割の向き", DIRECTION_ITEMS, "vertical", "レイアウトが「分割」のとき、子の表を縦に並べるか左右に並べるか");

    /** レベル 1 のセグメントが 2 つ以上のときの合計のブロックを出すか（行の「合計行」と同じ形。2026-10-16 ユーザー） */
    rootTotal = new formattingSettings.ItemDropdown({
        name: "rootTotal",
        displayName: "合計の表示",
        description: "レベル 1 のセグメントが 2 つ以上あるとき、それを足した合計のブロック。非表示にすると、レベル 1 のセグメントを並べるだけにする（畳めない）",
        items: ROOT_TOTAL_ITEMS,
        value: ROOT_TOTAL_ITEMS[0],
        visible: false,
    });

    rootName = new formattingSettings.TextInput({
        name: "rootName",
        displayName: "合計の名前",
        description: "レベル 1 のセグメントが 2 つ以上あるとき、それを足した合計のブロックの名前",
        value: "",
        placeholder: "合計",
        visible: false,
    });

    subtotalName = new formattingSettings.TextInput({
        name: "subtotalName",
        displayName: "小計の名前",
        description: "子のセグメントを足した小計のブロックの名前。{名前} にセグメントの名前が入る（{名前}計・{名前} 合計）。空なら「{名前}計」",
        value: "",
        placeholder: "{名前}計",
        visible: false,
    });

    hideEmptyAccounts = new formattingSettings.ToggleSwitch({
        name: "hideEmptyAccounts",
        displayName: "セグメントにデータの無い行を隠す",
        description: "セグメントのブロックで、そのセグメントにデータ（ファクト）の無い行と、下の行が全部隠れた小計の行を出さない。データがあって値が 0 の行は出す",
        value: true,
    });

    segmentHeader = new formattingSettings.TextInput({
        name: "segmentHeader",
        displayName: "セグメントの名前の列の見出し",
        description: "セグメントを左の列に並べたとき（横積み）の、列の見出しの文字。空なら「セグメント」",
        value: "",
        placeholder: "セグメント",
    });
    segmentHeaderBold = new formattingSettings.ToggleSwitch({ name: "segmentHeaderBold", displayName: "セグメントの名前の列の見出しを太字", description: "横積みのときのセグメントの列の見出しを太字にする", value: false });

    segmentFill = new formattingSettings.ToggleSwitch({
        name: "segmentFill",
        displayName: "囲みの塗り",
        description: "セグメントの囲みの名前の所を、レベルごとの色で塗る（切ると線だけ）。既定は塗らない",
        value: false,
    });
    segmentColor = new formattingSettings.ColorPicker({
        name: "segmentColor",
        displayName: "囲みの色（レベル 1）",
        description: "レベル 1 のセグメントの名前の所の塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_SEGMENT_COLORS[0] },
    });
    segmentColor2 = new formattingSettings.ColorPicker({
        name: "segmentColor2",
        displayName: "囲みの色（レベル 2）",
        description: "そのレベルのセグメントの名前の所の塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_SEGMENT_COLORS[1] },
    });
    segmentColor3 = new formattingSettings.ColorPicker({
        name: "segmentColor3",
        displayName: "囲みの色（レベル 3）",
        description: "そのレベルのセグメントの名前の所の塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_SEGMENT_COLORS[2] },
    });
    segmentColor4 = new formattingSettings.ColorPicker({
        name: "segmentColor4",
        displayName: "囲みの色（レベル 4 から）",
        description: "そのレベルのセグメントの名前の所の塗り。「囲みの塗り」を入れたときに効く",
        value: { value: DEFAULT_SEGMENT_COLORS[3] },
    });
    blockLine = new formattingSettings.ColorPicker({
        name: "blockLine",
        displayName: "囲み線の色",
        description: "セグメントのレベルごとの縦線と、セグメントの切れ目の横線",
        value: { value: DEFAULT_LINE_COLORS.block },
    });
    orgBold = new formattingSettings.ItemDropdown({
        name: "orgBold",
        displayName: "セグメント名の太字",
        description: "自動はビジュアルの標準の太さ。太字・太字にしないを選ぶと、この文字を全部その太さにそろえる",
        items: BOLD_ITEMS,
        value: BOLD_ITEMS[0],
    });
    orgSize = new formattingSettings.NumUpDown({ name: "orgSize", displayName: "セグメント名の文字サイズ", description: "セグメントの名前の字", value: DEFAULT_TEXT_SIZES.org });

    /** 適用先（「行」と同じ形）。セグメントを選ぶと、その配置と、そのセグメントで隠す行を出す（arrangeCards） */
    scope = new formattingSettings.ItemDropdown({
        name: "scope",
        displayName: "対象",
        description:
            "すべてのセグメントは表全体の設定。レベル・セグメントを選ぶと、その対象のレイアウト・小計の位置・分割の向きを変える。レベルを選んだ設定はすべてのセグメントの設定より強く、セグメントを選んだ設定はレベルより強い（同じ対象が 2 つあれば後の枠）。セグメントを選ぶと、そのセグメントで隠す行も選べる",
        items: [{ value: SCOPE_ALL, displayName: ALL_SEGMENTS_NAME }],
        value: { value: SCOPE_ALL, displayName: ALL_SEGMENTS_NAME },
        visible: false,
    });
    scopeSlotNumber: number | null = null;
    scopeFull = false;
    /** セグメントを選んだときの、隠す行の枠（1 から）：決めた枠と、次に足す枠 */
    hideSlots: { listed: number[]; add: number | null } | null = null;
    scopeGroup = new formattingSettings.Group({ name: "segmentsScope", displayName: "設定の適用先", slices: [this.scope] });

    // 群は「行」カードとそろえる（並べ方・合計・囲みの塗りと線・文字。2026-10-16 ユーザー「セグメントと行で同じ表現に」）。名前は保存に関わらない
    order = new formattingSettings.Group({
        name: "segmentsOrder",
        displayName: "並べ方",
        slices: [this.orgStyle, this.orgTotal, this.orgDirection, this.hideEmptyAccounts],
    });
    total = new formattingSettings.Group({ name: "segmentsTotal", displayName: "合計と小計", slices: [this.rootTotal, this.rootName, this.subtotalName] });
    box = new formattingSettings.Group({
        name: "segmentsBoxes",
        displayName: "囲みの塗りと線",
        slices: [this.segmentFill, this.segmentColor, this.segmentColor2, this.segmentColor3, this.segmentColor4, this.blockLine],
    });
    text = new formattingSettings.Group({ name: "segmentsText", displayName: "文字", slices: [this.segmentHeader, this.segmentHeaderBold, this.orgSize, this.orgBold] });

    groups = [this.scopeGroup, this.order, this.total, this.box, this.text];

    /** 「セグメント」フィールドを入れたときだけ出す。合計の名前は、レベル 1 が 2 つ以上で合計を置くときだけ。小計の名前は、子のあるセグメントがあるときだけ */
    applyOrgs(hasOrgs: boolean, hasRoot: boolean, hasSubtotals = false): void {
        this.visible = hasOrgs;
        this.rootTotal.visible = hasOrgs && hasRoot;
        this.rootName.visible = hasOrgs && hasRoot && this.rootTotal.value.value !== ROOT_TOTAL.off;
        this.subtotalName.visible = hasOrgs && hasSubtotals;
    }

    /** レベルごとのセグメントの囲みの色（外側から） */
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
    { value: CALC_KINDS.ratio, displayName: "割り算" },
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
                "小計：開始位置から終了位置までのレベル 1 の行を足す（営業利益・負債合計・フリー CF など）。割り算：分子の行 ÷ 分母の行（利益率・原価率・1 人当たり・時間当たり）",
            items: CALC_KIND_ITEMS,
            value: CALC_KIND_ITEMS[0],
        });
        this.rowName = new formattingSettings.TextInput({
            name: calcProp("name", slot),
            displayName: "名前",
            description: "表に出す行の名前。設定の適用先の一覧にもこの名前で出る",
            value: "",
            placeholder: "営業利益・売上原価率など",
        });
        this.from = new formattingSettings.ItemDropdown({
            name: calcProp("from", slot),
            displayName: "開始位置",
            description: "小計を足し始めるレベル 1 の行（表の並び）。負債合計は流動負債から",
            items: [NO_ITEM],
            value: NO_ITEM,
        });
        this.after = new formattingSettings.ItemDropdown({
            name: calcProp("after", slot),
            displayName: "挿入位置",
            description: "このレベル 1 の行の後に出す。小計はここまで足す",
            items: [NO_ITEM],
            value: NO_ITEM,
        });
        this.numerator = new formattingSettings.ItemDropdown({
            name: calcProp("numerator", slot),
            displayName: "分子",
            description: "割り算の上に来る行（表の行と、表に出さない指標から選ぶ）",
            items: [NO_ITEM],
            value: NO_ITEM,
        });
        this.denominator = new formattingSettings.ItemDropdown({
            name: calcProp("denominator", slot),
            displayName: "分母",
            description: "割り算の下に来る行。期間ごとに分子の合計 ÷ 分母の合計で割る",
            items: [NO_ITEM],
            value: NO_ITEM,
        });
        this.format = new formattingSettings.TextInput({
            name: calcProp("format", slot),
            displayName: "書式",
            description: "古い置き場所（今は「数値」カードの適用先でその行を選んで決める）。空にすると欄が消え、そこで決めた書式か既定（0.0%）になる",
            value: "",
            placeholder: "0.0%・#,0円",
        });
        this.good = new formattingSettings.ItemDropdown({
            name: calcProp("good", slot),
            displayName: "良し悪し",
            description: "古い置き場所（今は「数値」カードの適用先でその行を選んで決める）。「上がると良い」に戻すと欄が消える",
            items: CALC_GOOD_ITEMS,
            value: CALC_GOOD_ITEMS[0],
        });
        this.slices = [this.kind, this.rowName, this.from, this.after, this.numerator, this.denominator, this.format, this.good];
    }
}

/** 計算行の適用先で、空いた枠に新しい計算行を作る名前 */
export const NEW_CALC_ITEM = "＋ 新しい計算行";

/**
 * 計算行。段階利益・合計（小計）と割り算の行を、書式ペインで足す。区分の名前から型を当てて最初から入れる行は持たない。
 * 一番上の「並べ方」には「行」カードの「計算行を親にする」（保存先 rows.stepParents）を出す（arrangeCards）
 */
export class CalcRowsCardSettings extends formattingSettings.CompositeCard {
    name = "calcRows";
    displayName = "計算行";
    description = "段階利益・合計（小計）と割り算の行（利益率・1 人当たり）を足す。書式・良し悪し・太字は「数値」の適用先でその行を選ぶ";

    items = Array.from({ length: CALC_ROW_SLOTS }, (_, i) => new CalcRowItem(i + 1));
    custom = new formattingSettings.Group({
        name: "calcCustom",
        displayName: "追加行",
        description: "小計（開始位置 〜 終了位置のレベル 1 の行を足す。売上総利益・営業利益・資産合計など）と割り算（分子 ÷ 分母。利益率・自己資本比率など）。種類を選ぶと要る欄が出る",
        slices: [],
        container: new formattingSettings.Container({ displayName: "編集対象", containerItems: this.items }),
    });

    /**
     * 適用先（「数値」「行」と同じ形。2026-10-17 ユーザー「計算行も区分けできる？」）：作った計算行の名前と「＋ 新しい計算行」（空いた枠の
     * 1 本目）。値は枠の番号。表に元からある対象ではないので「すべて」は無い。コンテナーをやめたので、名前を変えても選択は 1 本目に戻らない
     */
    scope = new formattingSettings.ItemDropdown({
        name: "scope",
        displayName: "対象",
        description: "直す計算行。「＋ 新しい計算行」で種類を選ぶと行が足される。「使わない」に戻すと消える",
        items: [NO_ITEM],
        value: NO_ITEM,
    });
    scopeSlotNumber: number | null = null;
    scopeGroup = new formattingSettings.Group({ name: "calcScope", displayName: "設定の適用先", slices: [this.scope] });

    groups = [this.scopeGroup, this.custom];

    /**
     * 挿入位置・分子・分母の選択肢を流し込む。保存値は viewModel が生で読んだもの
     * （データ次第の選択肢は populate の時点では items に無いので value に入らない）。選択肢に無い保存値は「（表に無い）」で残す
     */
    applyCalc(choices: CalcChoices, saved: CalcSaved[], scope = ""): void {
        const kindOf = (i: number) => saved[i]?.kind ?? "";
        const made = this.items.flatMap((item, i) =>
            kindOf(i) === CALC_KINDS.subtotal || kindOf(i) === CALC_KINDS.ratio ? [{ value: String(item.slot), displayName: saved[i].name.trim() || `計算行 ${item.slot}` }] : [],
        );
        const free = this.items.find((_, i) => kindOf(i) !== CALC_KINDS.subtotal && kindOf(i) !== CALC_KINDS.ratio);
        this.scope.items = [...made, ...(free ? [{ value: String(free.slot), displayName: NEW_CALC_ITEM }] : [])];
        this.scope.value = itemOf(this.scope.items, scope);
        this.scopeSlotNumber = this.scope.items.length > 0 ? Number(this.scope.value.value) : null;
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
            // 種類で要る欄だけを出す（欄ごとのグレーアウトは部品に無い）：小計は名前・どこから・挿入位置、割り算は名前・挿入位置・分子・分母・書式・良し悪し
            item.rowName.visible = used;
            item.from.visible = kind === CALC_KINDS.subtotal;
            item.after.visible = used;
            // 小計の「挿入位置」は足す範囲の最後でもある（その区分まで足して、その後ろに置く）。どこからと組で読めるように名前を変える
            item.after.displayName = ratio ? "挿入位置" : "終了位置";
            item.numerator.visible = ratio;
            item.denominator.visible = ratio;
            // 書式・良し悪しは行別の書式に移した（2.1.7.15）。古い保存値があるときだけ、消せるように出す
            item.format.visible = ratio && s.format.trim() !== "";
            item.good.visible = ratio && s.good !== "" && s.good !== "up";
            item.from.items = keep([{ value: CALC_TABLE_START, displayName: "表の最初" }, ...sectionItems], s.from);
            item.from.value = itemOf(item.from.items, s.from || CALC_TABLE_START);
            item.after.items = keep([...sectionItems, { value: CALC_TABLE_END, displayName: "表の最後" }], s.after);
            item.after.value = itemOf(item.after.items, s.after || CALC_TABLE_END);
            // 分子・分母の選択肢は長い（表の行すべて）ので、割り算の行にだけ入れる
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
    /** メジャーの書式（モデルの書式の文字）。書式ペインの書式が空で率の書式なら、その率で出す（2026-10-24） */
    measureFormat?: string;
    good: string;
    /** 行の名前と数字を太字にするか（"on" | "off"） */
    bold: string;
    /** 横持ちのときのイベント（金額のメジャーの queryName か INDICATOR_ALL_EVENTS）。縦持ちでは null（イベントの列で分かれて届く） */
    event: string | null;
    eventChoices: Array<{ value: string; displayName: string }>;
}

/** 指標の欄の説明（メジャーごとのグループの欄で同じものを使う） */
const INDICATOR_HELP = {
    after: "この行の後ろか、この行のうちに出す。区分は区分の合計行（区分の中の行の後ろ）。表の最初は一番上の行の前（レベル 1 の配置を左右分割にすると、指標の表が左に来る）。既定は表の最後",
    placement:
        "表に出さない：行は出さず、計算行の分子・分母に選べる（率の分子と分母のメジャーを入れて、計算行の割り算で率を作る。どの期間でも分子の合計 ÷ 分母の合計になる。人数のような分母も「期間の集計」は合計（延べ人数）にする。期末にすると、四半期の時間を最後の月の人数で割ってしまう）。うち：挿入位置の行の下に 1 段下げて出し、親の行と同じ向き・集計・書式・良し悪しで見せる（値の行のうちは、「値」フィールドと同じ持ち方のメジャー。CALCULATE([値], 製品[区分] = \"新製品\") の形）。親の合計には足さない",
    aggregation: "四半期・通期の値。人数のような残高は期末。利益率のような率のメジャーは、どちらでも期間の率にならないので、分子と分母のメジャーを指標に入れて「表に出さない（計算行で使う）」にし、計算行の「割り算」で作る",
    format: "空なら「値」フィールドと同じ（表示単位で割る。EBITDA など）。#,0人・#,0.0h のように書くと数で出す（表示単位で割らない。字も出す）。0.0% は比率",
    good: "差の色の向き。既定は色を付けない",
    bold: "この指標の行の名前と数字を太字にする。「行」カードの「集計行を太字」を切っていても太字",
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
    description = "「指標」フィールドのメジャー（人数・時間・EBITDA・うち）の挿入位置と見せ方（メジャーごと）";
    visible = false;

    /** 適用先（計算行と同じ形）：メジャーの名前。値は queryName。下はそのメジャーの群だけを出す（arrangeCards） */
    scope = new formattingSettings.ItemDropdown({ name: "scope", displayName: "対象", description: "直す指標（メジャー）", items: [NO_ITEM], value: NO_ITEM });
    /** 選んだメジャーの群の番号（0 から） */
    scopeIndex: number | null = null;
    groups: formattingSettings.Group[] = [];

    applyIndicators(settings: IndicatorSetting[], places: Array<{ value: string; displayName: string }>, scope = ""): void {
        this.visible = settings.length > 0;
        this.scope.items = settings.length > 0 ? settings.map((setting) => ({ value: setting.queryName, displayName: setting.name })) : [NO_ITEM];
        this.scope.value = itemOf(this.scope.items, scope);
        const at = settings.findIndex((setting) => setting.queryName === this.scope.value.value);
        this.scopeIndex = at >= 0 ? at : null;
        const scopeGroup = new formattingSettings.Group({ name: "indicatorsScope", displayName: "設定の適用先", slices: [this.scope] });
        this.groups = [scopeGroup, ...settings.map((setting, i) => {
            const selector = { metadata: setting.queryName };
            const dropdown = (name: string, displayName: string, items: powerbi.IEnumMember[], value: string, description: string) =>
                new formattingSettings.ItemDropdown({ name, displayName, description, items, value: itemOf(items, value), selector });
            // 挿入位置の選択肢は表の最初・表の行・表の最後（自分の行は入れない）。表に無い保存値（フィルターで行が消えたなど）は選択肢に残す
            const places_: powerbi.IEnumMember[] = [
                { value: CALC_TABLE_START, displayName: "表の最初" },
                ...places.filter((p) => p.value !== setting.code),
                { value: CALC_TABLE_END, displayName: "表の最後" },
            ];
            if (!places_.some((p) => p.value === setting.after)) places_.push({ value: setting.after, displayName: `${missingLabel(setting.after)}（表に無い）` });
            const hidden = setting.placement === INDICATOR_MODES.hidden;
            const slices: formattingSettings.Slice[] = [
                // 表に出さない指標は置く場所が要らない
                ...(hidden ? [] : [dropdown("after", "挿入位置", places_, setting.after, INDICATOR_HELP.after)]),
                dropdown("placement", "配置", INDICATOR_MODE_ITEMS, setting.placement, INDICATOR_HELP.placement),
            ];
            // うちは親の行に合わせるので、集計の欄を出さない。表に出さない指標は集計だけ（計算行の分子・分母の期間の足し方）。
            // 書式・良し悪し・太字は「数値」カードに移した（2.1.7.15）。古い保存値は行の既定として効き、「数値」カードの設定が勝つ。
            // 古い欄は書式ペインに出さない（2.1.7.19、ユーザー「前の置き場所を消したい」）
            if (setting.placement !== INDICATOR_MODES.under) slices.push(dropdown("aggregation", "期間の集計", INDICATOR_AGGREGATION_ITEMS, setting.aggregation, INDICATOR_HELP.aggregation));
            // 横持ちだけ：指標の値のシナリオ
            if (setting.event !== null) {
                const events: powerbi.IEnumMember[] = [...setting.eventChoices, { value: INDICATOR_ALL_EVENTS, displayName: "すべてのシナリオ" }];
                slices.push(dropdown("event", "シナリオ", events, setting.event, INDICATOR_HELP.event));
            }
            return new formattingSettings.Group({ name: `indicator${i + 1}`, displayName: setting.name, slices });
        })];
    }
}

/** 基準・比較の値のマイナス・プラスの見せ方（numberFormat.ts の VALUE_SHOWS） */
export const VALUE_SHOW_ITEMS: powerbi.IEnumMember[] = [
    { value: "show", displayName: "そのまま" },
    { value: "blank", displayName: "出さない（空欄）" },
    { value: "dash", displayName: "「―」にする" },
];

/** 「数値」の適用先：すべての行（表全体） */
const SCOPE_ALL_ITEM: powerbi.IEnumMember = { value: SCOPE_ALL, displayName: ALL_ROWS_NAME };

/** 行別の数値書式の枠の数（決め打ち。計算行・イベントの金額の持ち方と同じ形） */
export const ROW_NUMBER_SLOTS = 8;
/** 行別の数値書式の枠で上書きする設定（名前は「数値」カードと同じ。書式は指標の書式と同じ書き方） */
export const ROW_NUMBER_PROPS = ["target", "unitType", "precision", "negativeStyle", "zeroStyle", "diffZeroStyle", "negativeZero", "negativeValue", "positiveValue", "good", "format", "currency", "caption", "bold"] as const;
export type RowNumberProp = (typeof ROW_NUMBER_PROPS)[number];
/** 枠 i（1 から）の設定の名前。1 つの object（rowNumbers）に target1・unitType1… と並べる */
export const rowNumberProp = (prop: RowNumberProp, slot: number): string => `${prop}${slot}`;
/** 行別の数値書式の枠の保存値（生の文字。空は ""。表全体に合わせるは ROW_NUMBER_TABLE） */
export type RowNumberSaved = Record<RowNumberProp, string>;
/** 「すべての行と同じ」（表全体の設定に合わせる）の値 */
export const ROW_NUMBER_TABLE = "table";
const TABLE_ITEM: powerbi.IEnumMember = { value: ROW_NUMBER_TABLE, displayName: "すべての行と同じ" };
/** 行で選べる表示単位は固定の単位だけ（自動は表全体の 1 つ。行ごとに自動で変わると縦に比べられない） */
const ROW_UNIT_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...UNIT_TYPES.filter((u) => u.value !== "auto")];
const ROW_PRECISION_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...PRECISIONS.filter((p) => p.value !== "auto")];
const ROW_NEGATIVE_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...NEGATIVE_STYLE_ITEMS];
const ROW_ZERO_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...ZERO_STYLE_ITEMS];
const ROW_VALUE_SHOW_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...VALUE_SHOW_ITEMS];
/** 行の良し悪し（差の色の向き）。既定は行の向き（収益は増えると良い、費用は減ると良い、残高は色を付けない） */
const ROW_GOOD_ITEMS: powerbi.IEnumMember[] = [{ value: ROW_NUMBER_TABLE, displayName: "行の向きのまま" }, ...CALC_GOOD_ITEMS];
/** 行の太字。計算行・後ろに置いた指標の行だけに効く（科目の木の行は「行」カードの太字） */
const ROW_BOLD_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, { value: "on", displayName: "太字" }, { value: "off", displayName: "太字にしない" }];
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
        negativeValue: ROW_VALUE_SHOW_ITEMS,
        positiveValue: ROW_VALUE_SHOW_ITEMS,
        good: ROW_GOOD_ITEMS,
        bold: ROW_BOLD_ITEMS,
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
    negativeValue: formattingSettings.ItemDropdown;
    positiveValue: formattingSettings.ItemDropdown;
    good: formattingSettings.ItemDropdown;
    bold: formattingSettings.ItemDropdown;
    format: formattingSettings.TextInput;
    currency: formattingSettings.TextInput;
    caption: formattingSettings.TextInput;

    constructor(readonly slot: number) {
        super();
        this.name = `rowNumber${slot}`;
        // 表示名は固定（計算行と同じ。名前を変えると Desktop で編集する枠の選択が 1 本目に戻った）
        this.displayName = `設定 ${slot}`;
        const dropdown = (prop: RowNumberProp, displayName: string, items: powerbi.IEnumMember[], description?: string) =>
            new formattingSettings.ItemDropdown({ name: rowNumberProp(prop, slot), displayName, description, items, value: items[0] });
        this.target = dropdown("target", "対象", [NO_ITEM], "レベル（そのレベルの行すべて）か、行（その行と下の行）、計算行・指標の行（その行だけ。レベルは効かない）。行はレベルより後に効き、下の行ほど後に効く");
        this.unitType = dropdown("unitType", "表示単位", ROW_UNIT_ITEMS, "表全体と違う単位にすると、親と単位の違う行の行見出しに単位を添える（「販管費（千円）」）");
        const same = "表全体の同じ名前の欄を、この行だけ変える";
        this.precision = dropdown("precision", "小数点以下の桁数", ROW_PRECISION_ITEMS, "この行の桁。すべての行と同じなら表全体の桁");
        this.negativeStyle = dropdown("negativeStyle", "マイナスの表記", ROW_NEGATIVE_ITEMS, same);
        this.zeroStyle = dropdown("zeroStyle", "0 の表記", ROW_ZERO_ITEMS, same);
        this.diffZeroStyle = dropdown("diffZeroStyle", "差 0 の表記", ROW_ZERO_ITEMS, same);
        this.negativeZero = dropdown("negativeZero", "丸めた 0 の符号", ROW_NEGATIVE_ZERO_ITEMS, same);
        this.negativeValue = dropdown("negativeValue", "マイナスの値の表示", ROW_VALUE_SHOW_ITEMS, "基準・比較の値がマイナスのとき、出さないか「―」にする");
        this.positiveValue = dropdown("positiveValue", "プラスの値の表示", ROW_VALUE_SHOW_ITEMS, "基準・比較の値がプラスのとき、出さないか「―」にする");
        this.good = dropdown("good", "良し悪し", ROW_GOOD_ITEMS, "差の色の向き。既定は行の向き（収益は増えると良い、費用は減ると良い、残高は色を付けない）。下の行も同じ向きになる");
        this.bold = dropdown("bold", "太字", ROW_BOLD_ITEMS, "計算行・指標の行の行見出しと数字を太字にするか。「行」カードの太字より強い。ほかの行は「行」カードの太字で決める");
        this.format = new formattingSettings.TextInput({
            name: rowNumberProp("format", slot),
            displayName: "書式",
            description: "#,0h・#,0台 のように書くと、表示単位で割らずに字を付けて出す（時間・台数の行）。0.0% は比率。空は金額（計算行・指標の行は、その行の既定のまま。指標の行の既定は、メジャーの書式が % なら % で、ほかは金額。「金額」と書けば金額）",
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
        this.caption = new formattingSettings.TextInput({
            name: rowNumberProp("caption", slot),
            displayName: "右上の見出し",
            description: "この行がレベル 1 にある表（左右・縦に分けた表）の右上に出す字（「台数：台」など）。空なら表の行から決める（金額の行があれば表全体の単位、台数だけなら「台数：台」）",
            value: "",
            placeholder: "台数：台",
        });
        this.slices = [this.target, this.unitType, this.precision, this.format, this.currency, this.caption, this.negativeStyle, this.zeroStyle, this.diffZeroStyle, this.negativeZero, this.negativeValue, this.positiveValue, this.good, this.bold];
    }
}

/**
 * 行別の書式（「数値」カードの設定を、段か行で上書きする）。対象は行のコード（区分・中分類・科目・計算行・指標）か段で覚える（枠の番号や
 * 並びでは覚えない）。計算行と後ろに置いた指標の行は、その行だけに効く（段は効かない。区分をまたぐ行なので）。2.1.7.15 で計算行・指標の
 * 書式・良し悪し・太字もここに集めた（2026-10-16 ユーザー決定。置き場所がばらけて、できることが違っていた）
 */
export class RowNumbersCardSettings extends formattingSettings.CompositeCard {
    name = "rowNumbers";
    displayName = "行別の書式";
    description = "「数値」カードの設定を、レベルか行ごとに上書きする。計算行・指標の行の書式・良し悪し・太字もここで決める。上書きしない設定は上のレベル・行か表全体に合わせる";
    visible = false;

    items = Array.from({ length: ROW_NUMBER_SLOTS }, (_, i) => new RowNumberItem(i + 1));
    group = new formattingSettings.Group({
        name: "rowNumbersGroup",
        displayName: "行ごと",
        description: "設定 1〜8 のそれぞれで、対象を 1 つ選び、変えるものを選ぶ。計算行（段階利益・合計・割り算）と後ろに置いた指標の行は、その行を選んだときだけ変わる（レベルは効かない）",
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
            set(item.negativeValue, slot?.negativeValue);
            set(item.positiveValue, slot?.positiveValue);
            set(item.good, slot?.good);
            set(item.bold, slot?.bold);
            item.format.value = slot?.format ?? "";
            item.currency.value = slot?.currency ?? "";
        });
    }
}

export class NumbersCardSettings extends formattingSettings.CompositeCard {
    name = "numbers";
    displayName = "数値";

    amountSign = new formattingSettings.ItemDropdown({
        name: "amountSign",
        displayName: "符号の持ち方",
        description:
            "メジャーが返す符号の持ち方。自動は、シナリオごとに「貸方フラグ」フィールドが 1 の行（貸方の行）と 0 の行（借方の行）の合計の符号で見分ける。表は小計の行の向き（その中の行の「貸方フラグ」の値の多数）でプラスに見せる。「貸方フラグ」フィールドが無ければ届いた値のまま",
        items: AMOUNT_SIGN_ITEMS,
        value: AMOUNT_SIGN_ITEMS[0],
    });

    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        description: "「値」フィールドの行だけを割る（% の行・時間・人数の行は割らない）。自動は、一番大きい値が 4 けた以上残る単位を表全体で 1 つ選ぶ",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0],
    });

    unitNotation = new formattingSettings.ItemDropdown({
        name: "unitNotation",
        displayName: "単位の表記",
        description: "表示単位の字を日本語（千・百万・億）か英語（K・M・bn）で書く。「円」は単位文字",
        items: UNIT_NOTATIONS,
        value: UNIT_NOTATIONS[0],
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        description: "「値」フィールドの行の桁。自動は、単位で割った値が 100 未満なら 1 桁、ほかは 0 桁（表全体でそろえる）",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    currency = new formattingSettings.TextInput({
        name: "currency",
        displayName: "単位文字",
        description: "右上の「単位：百万円」の「円」。空にすると付けない（台数・人数などの値のとき）",
        value: DEFAULT_CURRENCY,
        placeholder: DEFAULT_CURRENCY,
    });

    negativeStyle = new formattingSettings.ItemDropdown({
        name: "negativeStyle",
        displayName: "マイナスの表記",
        description: "マイナスの数字の書き方（-1,234・▲1,234・△1,234・(1,234)）",
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

    negativeValue = new formattingSettings.ItemDropdown({
        name: "negativeValue",
        displayName: "マイナスの値の表示",
        description: "基準・比較の値がマイナスのとき、出さないか「―」にする。差・比・率と、小計・計算行の足し算は値のまま",
        items: VALUE_SHOW_ITEMS,
        value: VALUE_SHOW_ITEMS[0],
    });

    positiveValue = new formattingSettings.ItemDropdown({
        name: "positiveValue",
        displayName: "プラスの値の表示",
        description: "基準・比較の値がプラスのとき、出さないか「―」にする。差・比・率と、小計・計算行の足し算は値のまま",
        items: VALUE_SHOW_ITEMS,
        value: VALUE_SHOW_ITEMS[0],
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
        description: "これを超える比・率と、% の行（計算行の割り算・% の指標）の値・差は「≧999%」のように上限で止める（100〜99999）",
        value: DEFAULT_RATIO_CAP,
    });

    /**
     * 適用先：すべての行（表全体の設定）か、行別の書式の枠 1 本。下の群は選んだ適用先の値を見せて書き換える（arrangeCards が欄の保存先を
     * 枠に差し替える。2026-10-17 ユーザー「一番上で対象を選び、すべてか個別か」）。選んだ値も保存し、開き直すと最後の適用先が出る
     */
    scope = new formattingSettings.ItemDropdown({
        name: "scope",
        displayName: "対象",
        description: "すべての行は表全体の書式。行（レベル・行・計算行・指標の行）を選ぶと、下の欄はその行だけを変える（「すべての行と同じ」の欄は表全体に合わせる）。決めた対象には「（変更あり）」が付く。対象は 8 つまで",
        items: [SCOPE_ALL_ITEM],
        value: SCOPE_ALL_ITEM,
        visible: false,
    });

    // グループの名前は書式ペインの見出しだけ。保存先（numbers の各項目）は変えない
    scopeGroup = new formattingSettings.Group({ name: "numbersScope", displayName: "設定の適用先", slices: [this.scope] });
    unitGroup = new formattingSettings.Group({
        name: "numbersUnit",
        displayName: "単位",
        description: "単位の置き場所・文字サイズ・色は「表全体 › 表の上」、「単位：」の字は「用語 › 単位の見出し」",
        slices: [this.unitType, this.unitNotation, this.currency, this.alignTails],
    });
    // 比・率の上限は表全体だけの欄。適用先で行を選ぶと、この群は枠の欄に差し替わる（arrangeCards）ので出ない
    digitsGroup = new formattingSettings.Group({
        name: "numbersDigits",
        displayName: "桁と記号",
        slices: [this.precision, this.negativeStyle, this.zeroStyle, this.diffZeroStyle, this.negativeZero, this.ratioCap],
    });
    hideGroup = new formattingSettings.Group({ name: "numbersHide", displayName: "値を隠す", slices: [this.negativeValue, this.positiveValue] });
    signGroup = new formattingSettings.Group({ name: "numbersSign", displayName: "データの符号", slices: [this.amountSign] });

    groups = [this.scopeGroup, this.unitGroup, this.digitsGroup, this.hideGroup, this.signGroup];

    /** 適用先に選んだ対象をつなぐ枠（1 から）。すべての行なら null */
    scopeSlotNumber: number | null = null;

    /** 適用先の選択肢：すべての行と、行別の書式の対象すべて（レベル・行・計算行・指標の行）。行別の書式が出ないときは出さない（resolveScope） */
    applyScope(rows: RowNumbersCardSettings, targets: Array<powerbi.IEnumMember & { own?: boolean }>, saved: RowNumberSaved[], scope: string): ScopeChoice {
        const slots = rows.items.map((_, i) => ({
            target: saved[i]?.target ?? "",
            used: ROW_NUMBER_PROPS.some((p) => p !== "target" && (saved[i]?.[p] ?? "") !== "" && saved[i]?.[p] !== ROW_NUMBER_TABLE),
        }));
        const choice = resolveScope(SCOPE_ALL_ITEM.displayName as string, rows.visible ? targets : [], slots, rows.visible ? scope : SCOPE_ALL);
        this.scope.visible = rows.visible;
        this.scope.items = choice.items;
        this.scope.value = itemOf(choice.items, choice.value);
        this.scopeSlotNumber = choice.slot;
        this.scopeFull = choice.value !== SCOPE_ALL && choice.slot === null;
        this.scopeOwnRow = targets.some((t) => t.value === choice.value && t.own === true);
        return choice;
    }

    /** 適用先に選んだ対象に、空いた枠が無い（設定できる対象の数を超えた） */
    scopeFull = false;
    /** 適用先が計算行・指標の行（太字が効く） */
    scopeOwnRow = false;

    /** 選んだ適用先の枠の番号（1 から）。すべての行なら null */
    scopeSlot(): number | null {
        return this.scopeSlotNumber;
    }
}

/** 表全体：表の上の文字、フォントと色、余白、見出しの背景と線、外枠、見出しの列の区切り、画像のコピー */
export class TableCardSettings extends formattingSettings.CompositeCard {
    name = "table";
    displayName = "表全体";

    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "フォント", description: "表の中の字のフォント", value: DEFAULT_FONT_FAMILY });
    textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "文字の色",
        description: "表の中の文字（行見出し・数字・列の見出し）と、上のボタン・メニューの文字の色。空ならテーマの文字の色。良い差・悪い差の色と、濃い塗りの上の白い字はそのまま",
        value: { value: "" },
    });
    mutedColor = new formattingSettings.ColorPicker({
        name: "mutedColor",
        displayName: "薄い文字の色",
        description: "単位・見出しの 2 段目・シナリオの名前などの薄い文字の色。空ならテーマの薄い文字の色",
        value: { value: "" },
    });
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "年度・単位の文字サイズ",
        description: "表のすぐ上の年度と単位（表の右上に置いた単位も）。上のボタン・メニュー・ダイアログの文字は変わらない。表の中の文字は、行・列のカードで要素ごとに変える",
        value: DEFAULT_HEADING_SIZE,
    });
    headingColor = new formattingSettings.ColorPicker({
        name: "headingColor",
        displayName: "年度・単位の色",
        description: "表のすぐ上の年度と単位の文字の色。空なら表の文字と同じ色。左上の角に置いた年度は、期間の見出しと同じ大きさと色",
        value: { value: "" },
    });
    // 年度の数え方・年度の位置（2.1.7.8 で「期間の列」から）、単位の位置（「数値」から）。保存値の読み替えは applyMoved
    yearLabel = new formattingSettings.ItemDropdown({
        name: "yearLabel",
        displayName: "年度の数え方",
        description: "年をまたぐ年度を、始まりの年（2025年度）と終わりの年（FY2026）のどちらで呼ぶか",
        items: YEAR_LABEL_ITEMS,
        value: YEAR_LABEL_ITEMS[0],
    });

    titlePlace = new formattingSettings.ItemDropdown({
        name: "titlePlace",
        displayName: "年度の位置",
        description: "左上の角にすると、行見出しの列の上（期間の見出しと同じ段）に置く",
        items: TITLE_PLACE_ITEMS,
        value: TITLE_PLACE_ITEMS[0],
    });
    unitPlace = new formattingSettings.ItemDropdown({
        name: "unitPlace",
        displayName: "単位の位置",
        description: "「単位：百万円」を置く所。行見出しの横にすると、値の行は行見出しに（百万円）、台数・時間のような指標は数字から単位の字を外して行見出しに（台）を添える（% の行の % は数字に残す）",
        items: UNIT_PLACE_ITEMS,
        value: UNIT_PLACE_ITEMS[0],
    });
    headBackground = new formattingSettings.ColorPicker({
        name: "headBackground",
        displayName: "見出しの背景",
        description: "期間と列の見出し、左上の角、セグメントの列の見出し",
        value: { value: DEFAULT_HEAD_BACKGROUND },
    });
    headLine = new formattingSettings.ColorPicker({ name: "headLine", displayName: "見出しの下線の色", description: "期間・列の見出しの下の線", value: { value: DEFAULT_LINE_COLORS.head } });
    nameLine = new formattingSettings.ColorPicker({
        name: "nameLine",
        displayName: "見出しの列の区切り線の色",
        description: "セグメントの列と行見出しの列の右の縦線（左上の角も）",
        value: { value: DEFAULT_LINE_COLORS.name },
    });
    outerLine = new formattingSettings.ColorPicker({ name: "outerLine", displayName: "表の外枠の色", description: "表のまわりの線", value: { value: DEFAULT_LINE_COLORS.outer } });
    copyButton = new formattingSettings.ToggleSwitch({
        name: "copyButton",
        displayName: "画像のコピー",
        description: "マウスを乗せたとき、右下に「画像としてコピー」のボタンを出す。押すと、見えている表を画像としてクリップボードに入れ、PowerPoint などに貼れる",
        value: true,
    });

    rowHeight = new formattingSettings.ItemDropdown({
        name: "rowHeight",
        displayName: "行の余白",
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
    top = new formattingSettings.Group({
        name: "tableTop",
        displayName: "表の上（年度と単位）",
        description:
            "年度と単位の書き方の文字（「FY{年2}」「単位の見出し」など）は「用語」カード。左上の角に置いたときは、ここの文字サイズと色は効かない。年度は「期間の列 › 期間の見出しの文字サイズ」と期間の見出しの色、単位は「シナリオの列 › 基準列の見出しの文字サイズ」に従う（少し小さく出す）",
        slices: [this.yearLabel, this.titlePlace, this.unitPlace, this.fontSize, this.headingColor],
    });
    text = new formattingSettings.Group({ name: "tableText", displayName: "文字", slices: [this.fontFamily, this.textColor, this.mutedColor] });
    spacing = new formattingSettings.Group({ name: "tableSpacing", displayName: "余白", slices: [this.rowHeight, this.columnPadding] });
    head = new formattingSettings.Group({ name: "tableHead", displayName: "見出しと枠", slices: [this.headBackground, this.headLine, this.nameLine, this.outerLine] });
    // トグルはグループの見出しに置く（グループ名と項目名が同じになるのを避ける）
    copy = new formattingSettings.Group({ name: "tableCopy", displayName: "画像のコピー", description: this.copyButton.description, topLevelSlice: this.copyButton, slices: [] });

    groups = [this.top, this.text, this.spacing, this.head, this.copy];

    /** 年度の数え方・年度の位置は、「月」フィールドが無ければ出さない（「期間の列」の年度と同じ） */
    applyPeriods(hasPeriods: boolean): void {
        this.yearLabel.visible = hasPeriods;
        this.titlePlace.visible = hasPeriods;
    }
}

/** update() のたびにデータから流し込むもの */
export interface DataDrivenFormat {
    hierarchy?: HierarchyFormat;
    /** セグメントごとに隠す行：セグメントと行の選択肢、枠の保存値 */
    segmentRows?: SegmentRowsFormat;
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
    rowNumbers?: { targets: Array<{ value: string; displayName: string; own?: boolean }>; saved: RowNumberSaved[]; scope?: string };
    /** 計算行（枠の番号）と指標（メジャーの queryName）の適用先の保存値（生。選択肢はデータ次第なので populate では入らない） */
    scopes?: { calcRows?: string; indicators?: string };
    /** 組織の欄を入れたか（組織の設定を出す）と、一番上の段が 2 つ以上で全社を置いたか（全社の名前を出す） */
    hasOrgs?: boolean;
    hasRoot?: boolean;
    /** 子のある組織（根のほか）があるか（セグメントの計の名前を出す） */
    hasSubtotals?: boolean;
    /** 月の欄を入れたか（無ければ期間の列のカードの年度と出す列を出さない） */
    hasPeriods?: boolean;
}

/** 用語：ビジュアルが決めている言葉（最新見込み・前年同期・差・比・率・期間の名前・年度・単位の見出し）を書き換える。空なら既定の字（words.ts） */
export class WordsCardSettings extends SimpleCard {
    name = "words";
    displayName = "用語";
    description = "表に出る言葉を書き換える。空なら既定の字";
    latest = new formattingSettings.TextInput({ name: "latest", displayName: "最新値", description: "シナリオを月ごとに確かなものから採った基準の名前（基準・比較のメニュー。列の見出しでは、比較の列がほかの比較の列と同じ名前になるときに使う）", value: "", placeholder: "最新値" });
    priorYear = new formattingSettings.TextInput({
        name: "priorYear",
        displayName: "前年同期",
        description: "前の年の同じ月と比べる比較の名前。データのシナリオではなく、同じシナリオの 12 か月前の値をビジュアルが取る",
        value: "",
        placeholder: "前年同期",
    });
    diff = new formattingSettings.TextInput({ name: "diff", displayName: "差", description: "比較の差の列の名前の後ろに付く字（予定差）", value: "", placeholder: "差" });
    ratio = new formattingSettings.TextInput({ name: "ratio", displayName: "比", description: "比較の比の列の名前の後ろに付く字（予定比）", value: "", placeholder: "比" });
    rate = new formattingSettings.TextInput({ name: "rate", displayName: "率", description: "比較の率の列の名前の後ろに付く字（予定率）", value: "", placeholder: "率" });
    quarter = new formattingSettings.TextInput({ name: "quarter", displayName: "四半期", description: "四半期の名前。{n} に 1〜4 が入る（第{n}四半期・Q{n}）", value: "", placeholder: "{n}Q" });
    firstHalf = new formattingSettings.TextInput({ name: "firstHalf", displayName: "上期", description: "前半の 6 か月の名前", value: "", placeholder: "上期" });
    secondHalf = new formattingSettings.TextInput({ name: "secondHalf", displayName: "下期", description: "後半の 6 か月の名前", value: "", placeholder: "下期" });
    fullYear = new formattingSettings.TextInput({ name: "fullYear", displayName: "通期", description: "年度の 12 か月の名前", value: "", placeholder: "通期" });
    ytd = new formattingSettings.TextInput({ name: "ytd", displayName: "累計", description: "累計の列（期首から値のある最後の月まで）の名前", value: "", placeholder: "累計" });
    sum = new formattingSettings.TextInput({ name: "sum", displayName: "合計の列", description: "見る人が足した月の合計の列の後ろに付く字（4〜6月計）", value: "", placeholder: "計" });
    allPeriods = new formattingSettings.TextInput({ name: "allPeriods", displayName: "全期間", description: "「月」フィールドを入れないときの 1 列の名前", value: "", placeholder: "全期間" });
    year = new formattingSettings.TextInput({
        name: "year",
        displayName: "年度の書き方",
        description: "年度の形。{年} に 4 桁、{年2} に下 2 桁の年が入る（{年}年度・FY{年2}）。どちらの年かは「表全体」カードの「年度の数え方」。空なら年度の数え方のまま",
        value: "",
        placeholder: "{年}年度",
    });
    unit = new formattingSettings.TextInput({ name: "unit", displayName: "単位の見出し", description: "「単位：百万円」の「：」より前の字。「値」フィールドが金額なら「金額」、台数なら「数量」など", value: "", placeholder: "単位" });
    slices = [this.latest, this.priorYear, this.diff, this.ratio, this.rate, this.quarter, this.firstHalf, this.secondHalf, this.fullYear, this.ytd, this.sum, this.allPeriods, this.year, this.unit];
}

export class VisualFormattingSettingsModel extends Model {
    hierarchyOverrides = new HierarchyOverrides();
    segmentRows = new SegmentRowsCard();
    /** 適用先に選んだ対象を、空いた枠に書き込む（applyData が決め、ビジュアルが persistProperties する。scope.ts） */
    reservations: Reservation[] = [];
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
    words = new WordsCardSettings();

    // 対象ごとのカード（2.0）：列（シナリオ・期間）→ セグメント → 行（科目・計算・指標）→ 数値 → 表全体。
    // セグメントを行の前にし、行・計算行・指標を並べた（2026-10-16 ユーザー）。段・行ごとの配置と隠す行は、組み立てたあと行・セグメントのカードに入れる（arrangeCards）
    cards = [
        this.comparison,
        this.events,
        this.periods,
        this.segments,
        this.rows,
        this.calcRows,
        this.indicators,
        this.hierarchyOverrides,
        this.segmentRows,
        this.numbers,
        this.rowNumbers,
        this.table,
        this.words,
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

    /**
     * 置き場所を移した項目（2.1.7.8）。新しい置き場所に保存値が無く、前の置き場所に保存値があれば、それを使う（今のレポートの設定を残す）。
     * 前の置き場所の項目は書式ペインに出さないが、保存値を読むために capabilities.json に残す。populateFormattingSettingsModel のあとに呼ぶ
     */
    applyMoved(objects: powerbi.DataViewObjects | undefined): void {
        const moved: Array<[string, formattingSettings.ToggleSwitch | formattingSettings.ItemDropdown, string]> = [
            [this.rows.name, this.rows.valueBold, "comparison"],
            [this.table.name, this.table.yearLabel, "periods"],
            [this.table.name, this.table.titlePlace, "periods"],
            [this.table.name, this.table.unitPlace, "numbers"],
        ];
        for (const [object, slice, oldObject] of moved) {
            if (objects?.[object]?.[slice.name] !== undefined) continue;
            const old = objects?.[oldObject]?.[slice.name];
            if (old === undefined) continue;
            if (slice instanceof formattingSettings.ToggleSwitch) slice.value = old === true;
            else slice.value = itemOf(slice.items, String(old));
        }
    }

    /** 「行」「セグメント」の適用先。枠（hierarchyOverrides）は行とセグメントで分け合う。セグメントを選んだら隠す行の枠（segmentRows）も決める */
    private applyHierarchyScopes(data: DataDrivenFormat): void {
        const h = data.hierarchy;
        const slots = this.hierarchyOverrides.items.map((_, i) => {
            const saved = h?.saved[i];
            return { target: saved?.target ?? "", used: !!saved && [saved.style, saved.total, saved.direction].some((v) => v !== "" && v !== "inherit") };
        });
        const keep = new Set([h?.scopes?.account ?? "", h?.scopes?.org ?? ""].filter((v) => v !== ""));
        const cards = [
            ["account", this.rows, ALL_ROWS_NAME],
            ["org", this.segments, ALL_SEGMENTS_NAME],
        ] as const;
        for (const [kind, card, allName] of cards) {
            const own = (t: string) => t.startsWith(`${kind}:`);
            const choices = (h?.targets ?? []).filter((t) => own(String(t.value))).map((t) => ({ ...t, displayName: String(t.displayName).replace(/^(行|セグメント)：/, "") }));
            const choice = resolveScope(allName, choices, slots, h?.scopes?.[kind] ?? "", { own, keep });
            card.scope.visible = choices.length > 0;
            card.scope.items = choice.items;
            card.scope.value = itemOf(choice.items, choice.value);
            card.scopeSlotNumber = choice.slot;
            card.scopeFull = choice.value !== SCOPE_ALL && choice.slot === null;
            if (choice.reserve !== null) {
                slots[choice.reserve - 1] = { target: choice.value, used: false };
                const item = this.hierarchyOverrides.items[choice.reserve - 1];
                item.kind = kind;
                const found = choices.find((c) => c.value === choice.value);
                if (found && !item.target.items.some((i) => i.value === choice.value)) item.target.items = [...item.target.items, found];
                item.target.value = itemOf(item.target.items, choice.value);
                this.reservations.push({ objectName: "hierarchyOverrides", properties: { [`target${choice.reserve}`]: choice.value } });
            }
        }
        // セグメントで隠す行：そのセグメントの枠と、行を足す枠（無ければ空いた枠にセグメントを書き込む）
        this.segments.hideSlots = null;
        const scope = String(this.segments.scope.value.value);
        const saved = data.segmentRows?.saved ?? [];
        if (this.segmentRows.visible && scope.startsWith("org:node:")) {
            const org = scope.slice("org:node:".length);
            const listed = saved.flatMap((s, i) => (s.org === org && s.row !== "" ? [i + 1] : []));
            let add = saved.findIndex((s) => s.org === org && s.row === "");
            if (add < 0) {
                const empty = saved.findIndex((s) => s.org === "");
                add = empty >= 0 ? empty : saved.findIndex((s) => s.row === "");
                if (add >= 0) {
                    const item = this.segmentRows.items[add];
                    item.org.value = itemOf(item.org.items, org);
                    this.reservations.push({ objectName: "segmentRows", properties: { [`org${add + 1}`]: org } });
                }
            }
            this.segments.hideSlots = { listed, add: add >= 0 ? add + 1 : null };
        }
    }

    /** 切った設定に従う欄は出さない（通し確認 2026-10-17：色なしでも差の色、区切りなしでも二重線と色、塗りなしでも囲みの色が出ていた） */
    private applyToggles(): void {
        const tone = String(this.comparison.toneMode.value.value);
        this.comparison.good.visible = tone === TONE_MODES.both;
        this.comparison.bad.visible = tone !== TONE_MODES.none;
        this.periods.aggregateDouble.visible = this.periods.periodLines.value;
        this.periods.periodLine.visible = this.periods.periodLines.value;
        for (const color of [this.rows.bandColor, this.rows.bandColor2, this.rows.bandColor3, this.rows.bandColor4]) color.visible = this.rows.bandFill.value;
        for (const color of [this.segments.segmentColor, this.segments.segmentColor2, this.segments.segmentColor3, this.segments.segmentColor4]) color.visible = this.segments.segmentFill.value;
    }

    applyData(data: DataDrivenFormat): void {
        this.applyToggles();
        this.hierarchyOverrides.apply(data.hierarchy);
        this.segmentRows.apply(data.segmentRows);
        this.comparison.applyEvents(data.mainItems, data.main);
        this.events.applyMeasures(data.measureSettings ?? []);
        this.periods.applyYears(data.years, data.fiscalYear, data.hasPeriods ?? true);
        this.table.applyPeriods(data.hasPeriods ?? true);
        this.calcRows.applyCalc(data.calcChoices ?? { sections: [], refs: [], places: [] }, data.calcSaved ?? [], data.scopes?.calcRows ?? "");
        this.indicators.applyIndicators(data.indicatorSettings ?? [], data.calcChoices?.places ?? [], data.scopes?.indicators ?? "");
        this.segments.applyOrgs(data.hasOrgs ?? false, data.hasRoot ?? false, data.hasSubtotals ?? false);
        this.rowNumbers.applyRowNumbers(data.rowNumbers?.targets ?? [], data.rowNumbers?.saved ?? []);
        this.reservations = [];
        const numbers = this.numbers.applyScope(this.rowNumbers, data.rowNumbers?.targets ?? [], data.rowNumbers?.saved ?? [], data.rowNumbers?.scope ?? "");
        this.applyHierarchyScopes(data);
        if (numbers.reserve !== null) {
            const item = this.rowNumbers.items[numbers.reserve - 1];
            item.target.value = itemOf(item.target.items, numbers.value);
            this.reservations.push({ objectName: "rowNumbers", properties: { [rowNumberProp("target", numbers.reserve)]: numbers.value } });
        }
    }
}

/** 「数値」「行」の適用先で、表全体の設定を選ぶ名前 */
export const ALL_ROWS_ITEM = ALL_ROWS_NAME;
/** 「セグメント」の適用先で、全部のセグメントの既定を選ぶ名前 */
export const ALL_SEGMENTS_ITEM = ALL_SEGMENTS_NAME;

type FCard = powerbi.visuals.FormattingCard;
type FGroup = powerbi.visuals.FormattingGroup;
type FSlice = powerbi.visuals.FormattingSlice;

/**
 * 書式ペインのカードを、対象ごとにまとめて見せる（2026-10-16〜17 ユーザー。表全体の設定と段・行ごとの上書きが別のカードで分かりにくかった）。
 * 保存先は今のまま（numbers・rowNumbers・hierarchyOverrides・segmentRows）で、組み立てた書式モデルの上で並べ替えるだけ
 * （保存先は欄ごとの objectName で決まるので、カードをまたいで置いても保存値は変わらず、読み替えが要らない）。
 * 「数値」「行」「セグメント」は一番上の「適用先」で対象を選ぶ（scope.ts）。すべてなら表全体の群のまま。対象を選ぶと、下の群はその対象の枠の欄に
 * 差し替え、表全体にしか無い群は出さない。書式ペインのコンテナー（編集対象）は使わない（ビジュアルから選択を動かせず、対象を選んでも
 * 「新しい設定」のままに見えた。2026-10-17 ユーザー）。
 * - 「数値」：単位・桁と記号・値を隠すを rowNumbers の枠の欄に、行だけの欄（良し悪しと太字）を足す
 * - 「行」「セグメント」：「配置」（レイアウト・小計の位置・分割の向き）を hierarchyOverrides の枠の欄に
 * - 「セグメント」：セグメントを選んだら「隠す行」（segmentRows の、そのセグメントの枠と、行を足す枠）
 * - 「計算行」：一番上の「並べ方」に、「行」カードの「計算行を親にする」（保存先 rows.stepParents）を移す
 */
export function arrangeCards(model: powerbi.visuals.FormattingModel, settings: VisualFormattingSettingsModel): powerbi.visuals.FormattingModel {
    const isCard = (c: unknown): c is FCard => typeof c === "object" && c !== null && "groups" in c;
    const cardOf = (name: string) => model.cards.find((c): c is FCard => isCard(c) && c.uid === `${name}-card`);
    const groupsOf = (card: FCard | undefined) => (card?.groups ?? []) as FGroup[];
    // 欄を保存先（object と項目の名前）で引く。隠れたカード（行別の書式・段・行ごとの配置）の欄も組み立てに入っている
    const byKey = new Map<string, FSlice>();
    for (const card of model.cards)
        if (isCard(card))
            for (const group of groupsOf(card))
                for (const slice of [...((group.slices ?? []) as FSlice[]), ...(group.container?.containerItems ?? []).flatMap((item) => (item.slices ?? []) as FSlice[])]) {
                    const descriptor = (slice as powerbi.visuals.SimpleVisualFormattingSlice).control?.properties as { descriptor?: powerbi.visuals.FormattingDescriptor } | undefined;
                    if (descriptor?.descriptor) byKey.set(`${descriptor.descriptor.objectName}.${descriptor.descriptor.propertyName}`, slice);
                }
    const pick = (objectName: string, slot: number, props: readonly string[]): FSlice[] =>
        props.flatMap((p): FSlice[] => {
            const slice = byKey.get(`${objectName}.${p}${slot}`);
            return slice ? [slice] : [];
        });
    const replaced = new Map<FCard, FCard | null>();
    const drop = (name: string) => {
        const card = cardOf(name);
        if (card) replaced.set(card, null);
        return card;
    };
    const rowNumbers = drop("rowNumbers");
    const overrides = drop("hierarchyOverrides");
    const segmentRows = drop("segmentRows");
    // 既定に戻すは、カードに入れた枠の保存値も戻す（配置の枠は対象の種類で行とセグメントに分ける）
    const kindOfSlot = (d: powerbi.visuals.FormattingDescriptor) => settings.hierarchyOverrides.items[Number(/(\d+)$/.exec(d.propertyName)?.[1] ?? 0) - 1]?.kind;
    const overrideDescriptors = (kind: HierarchyKind) => (overrides?.revertToDefaultDescriptors ?? []).filter((d) => kindOfSlot(d) === kind);
    const keep = (groups: FGroup[]) => groups.filter((g) => (g.slices ?? []).length > 0 || g.container);
    /** 適用先で対象を選んだら、適用先の群と、枠の欄に差し替えた群だけにする */
    const HIERARCHY_FULL = `レベル・行・セグメントごとの配置は合わせて ${HIERARCHY_SLOTS} まで。どれかを「すべての行と同じ」「すべてのセグメントと同じ」に戻すと空きます`;
    // 枠が足りないとき（full）は適用先だけを出し、見出しで知らせる。表全体の群を出すと、そこを変えて表全体が変わった（通し確認 2026-10-17）
    const arrange = (card: FCard | undefined, slot: number | null, scopeGroup: string, scoped: FGroup[], reverts: powerbi.visuals.FormattingDescriptor[], full = { on: false, message: "" }) => {
        if (!card) return;
        const groups = groupsOf(card);
        const scopeOnly = groups.filter((g) => g.uid === scopeGroup);
        replaced.set(card, {
            ...card,
            groups: keep(
                full.on
                    ? scopeOnly.map((g) => ({ ...g, displayName: full.message }))
                    : slot === null
                      ? groups
                      : [...scopeOnly, ...scoped],
            ),
            revertToDefaultDescriptors: [...(card.revertToDefaultDescriptors ?? []), ...reverts],
        });
    };

    // 数値：表全体の群と同じ見出しで、欄を枠の欄に差し替える。表全体だけの欄（比・率の上限・データの符号）は出さない
    const numberSlot = settings.numbers.scopeSlot();
    arrange(
        cardOf("numbers"),
        numberSlot,
        "numbersScope-group",
        numberSlot === null
            ? []
            : [
                  { uid: "numbersUnit-group", displayName: "単位", slices: pick("rowNumbers", numberSlot, ["unitType", "format", "currency", "caption"]) },
                  { uid: "numbersDigits-group", displayName: "桁と記号", slices: pick("rowNumbers", numberSlot, ["precision", "negativeStyle", "zeroStyle", "diffZeroStyle", "negativeZero"]) },
                  { uid: "numbersHide-group", displayName: "値を隠す", slices: pick("rowNumbers", numberSlot, ["negativeValue", "positiveValue"]) },
                  // 太字が効くのは計算行・指標の行だけ（段・科目の木の行は「行」カードの太字。通し確認 2026-10-17）
                  {
                      uid: "numbersRowStyle-group",
                      displayName: settings.numbers.scopeOwnRow ? "良し悪しと太字" : "良し悪し",
                      slices: pick("rowNumbers", numberSlot, settings.numbers.scopeOwnRow ? ["good", "bold"] : ["good"]),
                  },
              ],
        rowNumbers?.revertToDefaultDescriptors ?? [],
        { on: settings.numbers.scopeFull, message: `行ごとの設定は ${ROW_NUMBER_SLOTS} つまで。どれかを「すべての行と同じ」に戻すと空きます` },
    );

    // 行・セグメント：配置を枠の欄に
    const layout = (slot: number, uid: string): FGroup => ({ uid, displayName: "配置", slices: pick("hierarchyOverrides", slot, ["style", "total", "direction"]) });
    const rowSlot = settings.rows.scopeSlotNumber;
    arrange(cardOf("rows"), rowSlot, "rowsScope-group", rowSlot === null ? [] : [layout(rowSlot, "rowsArrange-group")], overrideDescriptors("account"), { on: settings.rows.scopeFull, message: HIERARCHY_FULL });
    const segmentSlot = settings.segments.scopeSlotNumber;
    const hide = settings.segments.hideSlots;
    const hideSlices: FSlice[] = hide
        ? [
              ...hide.listed.flatMap((slot) => pick("segmentRows", slot, ["row", "mode"])),
              ...(hide.add !== null ? pick("segmentRows", hide.add, ["row"]).map((s) => ({ ...s, displayName: "隠す行を足す" }) as FSlice) : []),
          ]
        : [];
    arrange(
        cardOf("segments"),
        segmentSlot,
        "segmentsScope-group",
        segmentSlot === null ? [] : [layout(segmentSlot, "segmentsArrange-group"), { uid: "segmentsHide-group", displayName: "隠す行", slices: hideSlices }],
        [...overrideDescriptors("org"), ...(segmentRows?.revertToDefaultDescriptors ?? [])],
        { on: settings.segments.scopeFull, message: HIERARCHY_FULL },
    );

    // 計算行：一番上に「並べ方」（「行」カードの計算行を親にする。保存先は rows.stepParents のまま）。選んだ枠の欄を、種類と名前・位置・分子と分母に
    // 分ける（コンテナーは使わない）
    const calcSlot = settings.calcRows.scopeSlotNumber;
    const calc = cardOf("calcRows");
    const stepParents = byKey.get(`${settings.rows.name}.${settings.rows.stepParents.name}`);
    const isStepParents = (d: powerbi.visuals.FormattingDescriptor) => d.objectName === settings.rows.name && d.propertyName === settings.rows.stepParents.name;
    if (calc) {
        const arrangeGroup: FGroup[] = stepParents ? [{ uid: "calcArrange-group", displayName: "並べ方", slices: [stepParents] }] : [];
        replaced.set(calc, {
            ...calc,
            groups: keep([
                ...arrangeGroup,
                ...(calcSlot === null
                    ? groupsOf(calc)
                    : [
                          ...groupsOf(calc).filter((g) => g.uid === "calcScope-group"),
                          { uid: "calcDefine-group", displayName: "種類と名前", slices: pick("calcRows", calcSlot, ["kind", "name"]) },
                          { uid: "calcPlace-group", displayName: "位置", slices: pick("calcRows", calcSlot, ["from", "after"]) },
                          { uid: "calcRatio-group", displayName: "分子と分母", slices: pick("calcRows", calcSlot, ["numerator", "denominator"]) },
                      ]),
            ]),
            // 既定に戻すも、出しているカードで戻す
            revertToDefaultDescriptors: [
                ...(calc.revertToDefaultDescriptors ?? []),
                ...(stepParents ? (cardOf("rows")?.revertToDefaultDescriptors ?? []).filter(isStepParents) : []),
            ],
        });
    }

    // 指標：選んだメジャーの群の欄を、位置・集計・古い置き場所に分ける（保存はメジャーの selector のまま）
    const indicators = cardOf("indicators");
    const at = settings.indicators.scopeIndex;
    if (indicators && at !== null) {
        const chosen = groupsOf(indicators).find((g) => g.uid === `indicator${at + 1}-group`);
        const slices = (chosen?.slices ?? []) as FSlice[];
        const nameOf = (slice: FSlice) =>
            ((slice as powerbi.visuals.SimpleVisualFormattingSlice).control?.properties as { descriptor?: powerbi.visuals.FormattingDescriptor } | undefined)?.descriptor?.propertyName ?? "";
        const part = (uid: string, displayName: string, props: string[]): FGroup => ({ uid, displayName, slices: slices.filter((s) => props.includes(nameOf(s))) });
        replaced.set(indicators, {
            ...indicators,
            groups: keep([
                ...groupsOf(indicators).filter((g) => g.uid === "indicatorsScope-group"),
                part("indicatorsPlace-group", "位置", ["after", "placement"]),
                part("indicatorsSum-group", "集計", ["aggregation", "event"]),
            ]),
        });
    }

    // 「計算行を親にする」は計算行のカードにだけ出す（「行」カードの並べ方からは外す）
    if (stepParents) {
        const rows = cardOf("rows");
        const shown = rows ? (replaced.get(rows) ?? rows) : undefined;
        if (rows && shown)
            replaced.set(rows, {
                ...shown,
                groups: keep(groupsOf(shown).map((g) => ({ ...g, slices: ((g.slices ?? []) as FSlice[]).filter((s) => s !== stepParents) }))),
                revertToDefaultDescriptors: (shown.revertToDefaultDescriptors ?? []).filter((d) => !isStepParents(d)),
            });
    }

    const cards = model.cards.flatMap((c) => (isCard(c) && replaced.has(c) ? (replaced.get(c) ? [replaced.get(c)!] : []) : [c]));
    return { ...model, cards };
}
