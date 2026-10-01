"use strict";

/**
 * 書式ペイン。カードの name とスライスの name は capabilities.json の objects と完全に一致させる
 * 。
 *
 * 主と比較・年度の選択肢と、横持ちのメジャーごとの比較順、計算の行の置く場所・分子・分母の選択肢、指標のメジャーごとの設定は
 * データ次第なので、update() のたびに viewModel の結果から流し込む（applyData）。
 */

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import SimpleCard = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;
import { HierarchySettings, HierarchyOverrides, HierarchyFormat } from "./hierarchySettings";

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
import { MeasureSetting, ORDER_NONE, SIGN_TABLE } from "./events";
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
export const UNIT_PLACES = { right: "right", corner: "corner", name: "name" } as const;
const UNIT_PLACE_ITEMS: powerbi.IEnumMember[] = [
    { value: UNIT_PLACES.right, displayName: "表の右上" },
    { value: UNIT_PLACES.corner, displayName: "左上の角（行の名前の上）" },
    { value: UNIT_PLACES.name, displayName: "行の名前の横（売上高（百万円）・台数（台））" },
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
    { value: "debit", displayName: "借方プラス（貸方の科目がマイナス）" },
    { value: "credit", displayName: "貸方プラス（借方の科目がマイナス。足すと利益）" },
    { value: "positive", displayName: "すべてプラス" },
];

/** 横持ちのメジャーごとの金額の持ち方。既定は表全体（「行」カードの金額の持ち方）に合わせる */
const MEASURE_SIGN_ITEMS: powerbi.IEnumMember[] = [{ value: SIGN_TABLE, displayName: "表全体に合わせる" }, ...AMOUNT_SIGN_ITEMS];

export const YEAR_LABEL_ITEMS: powerbi.IEnumMember[] = [
    { value: "start", displayName: "始まりの年（2025年度）" },
    { value: "end", displayName: "終わりの年（FY2026）" },
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
 * 主と比較。主は 1 つ。比較の列は書式ペインに持たない
 */
export class ComparisonCardSettings extends formattingSettings.CompositeCard {
    name = "comparison";
    displayName = "主と比較";

    main = new formattingSettings.ItemDropdown({
        name: "main",
        displayName: "主",
        description:
            "表の値に出すもの。「最新見込み」はセグメント × 月ごとに、数字のあるシナリオのうち比較順の一番大きいもの（比較順が 2 つ以上のシナリオにあるとき）。比較順の無いシナリオ（前年実績など）は主にしない。見る人は表の上のメニューで変えられる",
        items: [NO_ITEM],
        value: NO_ITEM,
    });

    diffSwap = new formattingSettings.ToggleSwitch({
        name: "diffSwap",
        displayName: "2 段の上下を入れ替える",
        description: "差と比・差と率の 2 段で、差を下、比・率を上に出す",
        value: false,
    });

    general = new formattingSettings.Group({ name: "comparisonMain", displayName: "主", slices: [this.main] });

    /** 2 段の見せ方の上下（比較の列は見る人が選ぶ） */
    twoRows = new formattingSettings.Group({ name: "comparisonTwoRows", displayName: "2 段の見せ方", slices: [this.diffSwap] });

    groups = [this.general, this.twoRows];

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
    description = "メジャーを積んだときの、メジャーごとの比較順と符号の持ち方";
    visible = false;

    orderGroup = new formattingSettings.Group({
        name: "eventsOrder",
        displayName: "比較順",
        description: "大きいほど確かで、最新見込みで先に採る。「比較だけ」は主にせず比較にだけ使う（前年実績など）。既定は欄に入れた順（後ろほど確か）",
        slices: [],
    });

    signGroup = new formattingSettings.Group({
        name: "eventsSign",
        displayName: "符号の持ち方",
        description: "メジャーが返す値の符号。既定は「行」カードの符号の持ち方（表全体）に合わせる",
        slices: [],
    });

    groups = [this.orderGroup, this.signGroup];

    applyMeasures(measures: MeasureSetting[]): void {
        this.visible = measures.length > 0;
        this.orderGroup.slices = measures.map((measure) => {
            const items: powerbi.IEnumMember[] = [{ value: ORDER_NONE, displayName: "比較だけ" }, ...measure.orderChoices.map((value) => ({ value, displayName: value }))];
            return new formattingSettings.ItemDropdown({
                name: "order",
                displayName: measure.name,
                items,
                value: itemOf(items, measure.order),
                selector: { metadata: measure.queryName },
            });
        });
        this.signGroup.slices = measures.map(
            (measure) =>
                new formattingSettings.ItemDropdown({
                    name: "sign",
                    displayName: measure.name,
                    items: MEASURE_SIGN_ITEMS,
                    value: itemOf(MEASURE_SIGN_ITEMS, measure.sign),
                    selector: { metadata: measure.queryName },
                })
        );
    }
}

export class PeriodsCardSettings extends SimpleCard {
    name = "periods";
    displayName = "期間";

    fiscalYear = new formattingSettings.ItemDropdown({
        name: "fiscalYear",
        displayName: "年度",
        items: [{ value: LATEST_YEAR, displayName: "最新" }],
        value: { value: LATEST_YEAR, displayName: "最新" },
    });

    fiscalStartMonth = new formattingSettings.ItemDropdown({
        name: "fiscalStartMonth",
        displayName: "期首の月",
        items: FISCAL_START_ITEMS,
        value: FISCAL_START_ITEMS[3],
    });

    yearLabel = new formattingSettings.ItemDropdown({
        name: "yearLabel",
        displayName: "年度の呼び方",
        items: YEAR_LABEL_ITEMS,
        value: YEAR_LABEL_ITEMS[0],
    });

    showMonths = new formattingSettings.ToggleSwitch({ name: "showMonths", displayName: "月", value: true });
    showQuarters = new formattingSettings.ToggleSwitch({ name: "showQuarters", displayName: "四半期", value: true });
    showHalves = new formattingSettings.ToggleSwitch({ name: "showHalves", displayName: "半期", value: false });
    showYear = new formattingSettings.ToggleSwitch({ name: "showYear", displayName: "通期", value: true });
    showYtd = new formattingSettings.ToggleSwitch({
        name: "showYtd",
        displayName: "累計",
        description: "期首から、比較順の一番大きいシナリオ（実績など）に値のある最後の月まで",
        value: false,
    });

    scrollStart = new formattingSettings.ItemDropdown({
        name: "scrollStart",
        displayName: "スクロールの最初の位置",
        description: "列がはみ出すとき、開いたときに左端（先頭）と右端（末尾）のどちらから見せるか",
        items: SCROLL_START_ITEMS,
        value: SCROLL_START_ITEMS[0],
    });

    slices = [
        this.fiscalYear,
        this.fiscalStartMonth,
        this.yearLabel,
        this.showMonths,
        this.showQuarters,
        this.showHalves,
        this.showYear,
        this.showYtd,
        this.scrollStart,
    ];

    /** 年度の選択肢（新しい順）。値は期首の月の「年-月」、先頭は「最新」 */
    applyYears(years: Array<{ value: string; displayName: string }>, selected: string): void {
        this.fiscalYear.items = [{ value: LATEST_YEAR, displayName: "最新" }, ...years];
        this.fiscalYear.value = itemOf(this.fiscalYear.items, selected);
    }
}

export class RowsCardSettings extends SimpleCard {
    name = "rows";
    displayName = "行";

    amountSign = new formattingSettings.ItemDropdown({
        name: "amountSign",
        displayName: "符号の持ち方",
        description:
            "メジャーが返す符号の持ち方。自動は、シナリオごとに貸方の科目と借方の科目の合計の符号で見分ける。表は区分の向き（区分の中の科目の貸方フラグの値の多数）でプラスに見せる",
        items: AMOUNT_SIGN_ITEMS,
        value: AMOUNT_SIGN_ITEMS[0],
    });

    parentPosition = new formattingSettings.ItemDropdown({
        name: "parentPosition",
        displayName: "小計の位置",
        description: "区分・中分類の小計の行を、科目の下に置くか上に置くか。内訳（うち）はいつも親の下",
        items: PARENT_POSITION_ITEMS,
        value: PARENT_POSITION_ITEMS[0],
    });

    orgTotalPosition = new formattingSettings.ItemDropdown({
        name: "orgTotalPosition",
        displayName: "セグメントの小計の位置",
        description: "子のセグメントを足した小計のブロック（開いた製造部など）を、子のセグメントの下に置くか上に置くか。区分の小計の位置とは別に選ぶ",
        items: ORG_TOTAL_POSITION_ITEMS,
        value: ORG_TOTAL_POSITION_ITEMS[0],
        visible: false,
    });

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
        displayName: "データの無い科目を隠す",
        description: "セグメントのブロックで、そのセグメントにデータ（ファクト）の無い科目の行と、科目が全部隠れた区分・中分類の行を出さない。データがあって値が 0 の科目は出す",
        value: true,
        visible: false,
    });

    bands = new formattingSettings.ToggleSwitch({
        name: "bands",
        displayName: "囲み",
        description: "区分・中分類とその科目を入れ子の箱で囲む。子の箱は親の箱の中に 1 段ずらして置き、親の色の帯が左に残る",
        value: true,
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
    bandColor4 = new formattingSettings.ColorPicker({
        name: "bandColor4",
        displayName: "囲みの色（4 段目から）",
        value: { value: DEFAULT_BAND_COLORS[3] },
    });

    segmentFill = new formattingSettings.ToggleSwitch({
        name: "segmentFill",
        displayName: "セグメントの塗り",
        description: "セグメントの箱の名前の所を、段ごとの色で塗る（切ると線だけ）",
        value: true,
        visible: false,
    });

    segmentColor = new formattingSettings.ColorPicker({
        name: "segmentColor",
        displayName: "セグメントの色（1 段目）",
        description: "一番外の段のセグメントの箱の塗り",
        value: { value: DEFAULT_SEGMENT_COLORS[0] },
        visible: false,
    });
    segmentColor2 = new formattingSettings.ColorPicker({ name: "segmentColor2", displayName: "セグメントの色（2 段目）", value: { value: DEFAULT_SEGMENT_COLORS[1] }, visible: false });
    segmentColor3 = new formattingSettings.ColorPicker({ name: "segmentColor3", displayName: "セグメントの色（3 段目）", value: { value: DEFAULT_SEGMENT_COLORS[2] }, visible: false });
    segmentColor4 = new formattingSettings.ColorPicker({
        name: "segmentColor4",
        displayName: "セグメントの色（4 段目から）",
        value: { value: DEFAULT_SEGMENT_COLORS[3] },
        visible: false,
    });

    slices = [
        this.amountSign,
        this.parentPosition,
        this.orgTotalPosition,
        this.rootName,
        this.hideEmptyAccounts,
        this.bands,
        this.bandFill,
        this.bandColor,
        this.bandColor2,
        this.bandColor3,
        this.bandColor4,
        this.segmentFill,
        this.segmentColor,
        this.segmentColor2,
        this.segmentColor3,
        this.segmentColor4,
    ];

    /** 組織の欄を入れたときだけ、組織の設定を出す。全社の名前は、一番上の段が 2 つ以上で全社を置くときだけ */
    applyOrgs(hasOrgs: boolean, hasRoot: boolean): void {
        this.orgTotalPosition.visible = hasOrgs;
        for (const slice of [this.segmentFill, this.segmentColor, this.segmentColor2, this.segmentColor3, this.segmentColor4]) slice.visible = hasOrgs;
        this.rootName.visible = hasOrgs && hasRoot;
        this.hideEmptyAccounts.visible = hasOrgs;
    }

    /** 段ごとの囲みの色（外側から） */
    bandColors(): string[] {
        return [this.bandColor, this.bandColor2, this.bandColor3, this.bandColor4].map((c, i) => c.value?.value || DEFAULT_BAND_COLORS[i]);
    }

    /** 段ごとのセグメントの箱の色（外側から） */
    segmentColors(): string[] {
        return [this.segmentColor, this.segmentColor2, this.segmentColor3, this.segmentColor4].map((c, i) => c.value?.value || DEFAULT_SEGMENT_COLORS[i]);
    }
}

/** 書式ペインで書ける計算の行の数（書式ペインの部品では足し消しできないので決め打ち）。損益と貸借対照表を 1 つの表に入れても足りる数 */
export const CALC_ROW_SLOTS = 20;
export const CALC_KINDS = { none: "none", subtotal: "subtotal", ratio: "ratio" } as const;
const CALC_KIND_ITEMS: powerbi.IEnumMember[] = [
    { value: CALC_KINDS.none, displayName: "使わない" },
    { value: CALC_KINDS.subtotal, displayName: "小計" },
    { value: CALC_KINDS.ratio, displayName: "比率" },
];

/** 縦持ちのイベントごとの金額の持ち方の枠の数（書式ペインの部品では足し消しできないので決め打ち。計算の行と同じ形） */
export const EVENT_SIGN_SLOTS = 8;
/** 枠 i（1 から）の設定の名前。1 つの object（eventSigns）に event1・sign1… と並べる */
export const eventSignProp = (prop: "event" | "sign", slot: number): string => `${prop}${slot}`;
/** 縦持ちのイベントごとの金額の持ち方の枠の保存値（生の文字。空は ""） */
export interface EventSignSaved {
    event: string;
    sign: string;
}

/** 金額の持ち方の枠 1 本（書式ペインのコンテナーの項目） */
class EventSignItem extends SimpleCard {
    event: formattingSettings.ItemDropdown;
    sign: formattingSettings.ItemDropdown;

    constructor(readonly slot: number) {
        super();
        this.name = `eventSign${slot}`;
        // 表示名は固定（計算の行と同じ。名前を変えると Desktop で編集する枠の選択が 1 本目に戻った）
        this.displayName = `設定 ${slot}`;
        this.event = new formattingSettings.ItemDropdown({ name: eventSignProp("event", slot), displayName: "シナリオ", items: [NO_ITEM], value: NO_ITEM });
        this.sign = new formattingSettings.ItemDropdown({
            name: eventSignProp("sign", slot),
            displayName: "符号の持ち方",
            items: MEASURE_SIGN_ITEMS,
            value: MEASURE_SIGN_ITEMS[0],
        });
        this.slices = [this.event, this.sign];
    }
}

/**
 * 縦持ち（イベントの列）のときの、イベントごとの金額の持ち方。イベントの名前で覚える（イベントの節点の identity は組織ごとに分かれ、
 * 組織をまたいで効かない・フィルターで組織が外れると消えるので使わない）。データに無いイベントの保存値は「（データに無い）」で残し、知らせる
 */
export class EventSignsCardSettings extends formattingSettings.CompositeCard {
    name = "eventSigns";
    displayName = "シナリオごとの符号の持ち方";
    description =
        "データのシナリオ（実績・予算・見通しなど）ごとに、値をどちらの符号で持っているか（借方の科目をプラスで持つか、貸方の科目をプラスで持つか）を決める。ふつうは自動で見分けるので、見分けが外れたシナリオだけ決める。決めないシナリオは「行」カードの符号の持ち方に合わせる";
    visible = false;

    items = Array.from({ length: EVENT_SIGN_SLOTS }, (_, i) => new EventSignItem(i + 1));
    group = new formattingSettings.Group({
        name: "eventSignsGroup",
        displayName: "シナリオごと",
        description: "設定 1〜8 のそれぞれで、シナリオを 1 つ選び、その符号の持ち方を選ぶ。使わない設定はシナリオを（なし）のままにする",
        slices: [],
        container: new formattingSettings.Container({ displayName: "編集する設定", containerItems: this.items }),
    });

    groups = [this.group];

    /** イベントの選択肢と保存値を流し込む。縦持ちのときだけ出す。選択肢に無い保存値は「（データに無い）」で残す */
    applyEventSigns(visible: boolean, events: string[], saved: EventSignSaved[]): void {
        this.visible = visible;
        const base: powerbi.IEnumMember[] = [NO_ITEM, ...events.map((e) => ({ value: e, displayName: e }))];
        this.items.forEach((item, i) => {
            const value = saved[i]?.event ?? "";
            item.event.items = value === "" || events.includes(value) ? base : [...base, { value, displayName: `${value}（データに無い）` }];
            item.event.value = itemOf(item.event.items, value);
            item.sign.value = itemOf(MEASURE_SIGN_ITEMS, saved[i]?.sign || SIGN_TABLE);
        });
    }
}

export const CALC_GOOD_ITEMS: powerbi.IEnumMember[] = [
    { value: "up", displayName: "上がると良い" },
    { value: "down", displayName: "上がると悪い" },
    { value: "neutral", displayName: "色を付けない" },
];
/** どこからの「表の最初」と、置く場所の「表の最後」（区分の名前とぶつからない値） */
export const CALC_TABLE_START = "__start__";
export const CALC_TABLE_END = "__end__";
/** 表に無い保存値の表示名：書式ペインの小計は「計算の行 N」、区分・指標は名前か queryName（キーの頭を外す）、科目はコード */
const missingLabel = (value: string) => (value.startsWith(CALC_KEY) ? `計算の行 ${value.slice(CALC_KEY.length)}` : value.replace(/^§[a-z]+:/, ""));
export const CALC_PROPS = ["kind", "name", "from", "after", "numerator", "denominator", "format", "good"] as const;
export type CalcProp = (typeof CALC_PROPS)[number];
/** 書式ペインの計算の行の保存値（生の文字。空は ""） */
export type CalcSaved = Record<CalcProp, string>;
/** 計算の行 i（1 から）の設定の名前。1 つの object（calcRows）に kind1・name1… と並べる */
export const calcProp = (prop: CalcProp, slot: number): string => `${prop}${slot}`;
export const EMPTY_CALC: CalcSaved = { kind: "", name: "", from: "", after: "", numerator: "", denominator: "", format: "", good: "" };

/** 計算の行の 1 本（書式ペインのコンテナーの項目。公式のサンプルどおり SimpleCard を継ぐ。name は capabilities のカード名でなく識別子） */
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
        this.displayName = `計算の行 ${slot}`;
        this.kind = new formattingSettings.ItemDropdown({
            name: calcProp("kind", slot),
            displayName: "種類",
            description:
                "小計：どこからの区分から置く場所の区分までを足す（営業利益・負債合計・フリー CF など）。比率：分子の行 ÷ 分母の行（利益率・原価率・1 人当たり・時間当たり）",
            items: CALC_KIND_ITEMS,
            value: CALC_KIND_ITEMS[0],
        });
        this.rowName = new formattingSettings.TextInput({ name: calcProp("name", slot), displayName: "名前", value: "", placeholder: "営業利益・売上原価率など" });
        this.from = new formattingSettings.ItemDropdown({
            name: calcProp("from", slot),
            displayName: "どこから",
            description: "小計を足し始める区分（表の並び）。負債合計は流動負債から",
            items: [NO_ITEM],
            value: NO_ITEM,
        });
        this.after = new formattingSettings.ItemDropdown({
            name: calcProp("after", slot),
            displayName: "置く場所",
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
 * 計算の行。段階利益・合計・比率を、書式ペインで足す（小計と比率）。区分の名前から型を当てて最初から入れる行は持たない
 *
 */
export class CalcRowsCardSettings extends formattingSettings.CompositeCard {
    name = "calcRows";
    displayName = "計算の行";
    description = "段階利益・合計（小計）と比率を足す";

    items = Array.from({ length: CALC_ROW_SLOTS }, (_, i) => new CalcRowItem(i + 1));
    custom = new formattingSettings.Group({
        name: "calcCustom",
        displayName: "足す行",
        description: "小計（どこから 〜 置く場所の区分を足す。売上総利益・営業利益・資産合計など）と比率（分子 ÷ 分母。利益率・自己資本比率など）。種類を選ぶと要る欄が出る",
        slices: [],
        container: new formattingSettings.Container({ displayName: "編集する行", containerItems: this.items }),
    });

    groups = [this.custom];

    /**
     * 置く場所・分子・分母の選択肢を流し込む。保存値は viewModel が生で読んだもの
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
            // 種類で要る欄だけを出す（欄ごとのグレーアウトは部品に無い）：小計は名前・どこから・置く場所、比率は名前・置く場所・分子・分母・書式・良し悪し
            item.rowName.visible = used;
            item.from.visible = kind === CALC_KINDS.subtotal;
            item.after.visible = used;
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

/** 指標の置き方：置く場所の行の後ろか、行の「うち」か */
export const INDICATOR_MODES = { after: "after", under: "under" } as const;
const INDICATOR_MODE_ITEMS: powerbi.IEnumMember[] = [
    { value: INDICATOR_MODES.after, displayName: "行の後ろ" },
    { value: INDICATOR_MODES.under, displayName: "行のうち（親の行と同じ見せ方）" },
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
    /** 行のコード（置く場所の選択肢から自分を外す） */
    code: string;
    /** 置く場所（CalcChoices.places の value か CALC_TABLE_END） */
    after: string;
    placement: string;
    aggregation: string;
    format: string;
    good: string;
    /** 横持ちのときのイベント（金額のメジャーの queryName か INDICATOR_ALL_EVENTS）。縦持ちでは null（イベントの列で分かれて届く） */
    event: string | null;
    eventChoices: Array<{ value: string; displayName: string }>;
}

/**
 * 指標。指標の欄のメジャーを、どの行の後ろか・どの行のうちに置くかと、期間の集計・書式・良し悪しをメジャーごとに決める。
 * 保存はメジャーの queryName の selector。設定ごとのグループに、メジャーごとの欄を並べる（横持ちの「イベント（メジャーごと）」と同じ形）。
 * メジャーごとの項目を入れ物（container）に並べる形は、Desktop で書式ペインの値を変えると、別のメジャー（最初の項目）の selector に
 * 書かれ、ほかのメジャーの保存値が消えた。
 * メジャーの数と名前はデータ次第なので、欄は applyIndicators で作る
 */
export class IndicatorsCardSettings extends formattingSettings.CompositeCard {
    name = "indicators";
    displayName = "指標";
    description = "指標の欄のメジャー（人数・時間・EBITDA・うち）を置く場所と見せ方（メジャーごと）";
    visible = false;

    afterGroup = new formattingSettings.Group({
        name: "indicatorsAfter",
        displayName: "置く場所",
        description: "この行の後ろか、この行のうちに出す（置き方）。区分は区分の合計の行（区分の科目の後ろ）。既定は表の最後",
        slices: [],
    });

    placementGroup = new formattingSettings.Group({
        name: "indicatorsPlacement",
        displayName: "置き方",
        description:
            "うち：置く場所の行の下に 1 段下げて出し、親の行と同じ向き・集計・書式・良し悪しで見せる（科目の行のうちは、値の欄と同じ持ち方のメジャー。CALCULATE([値], 製品[区分] = \"新製品\") の形）。親の合計には足さない",
        slices: [],
    });

    aggregationGroup = new formattingSettings.Group({
        name: "indicatorsAggregation",
        displayName: "期間の集計",
        description: "四半期・通期の値。人数のような残高は期末（うちは親の行に合わせるので出さない）",
        slices: [],
    });

    formatGroup = new formattingSettings.Group({
        name: "indicatorsFormat",
        displayName: "書式",
        description: "空なら値の欄と同じ（表示単位で割る。EBITDA など）。#,0人・#,0.0h のように書くと数で出す（表示単位で割らない。字も出す）。0.0% は比率（うちは親の行に合わせるので出さない）",
        slices: [],
    });

    goodGroup = new formattingSettings.Group({
        name: "indicatorsGood",
        displayName: "良し悪し",
        description: "差の色の向き。既定は色を付けない（うちは親の行に合わせるので出さない）",
        slices: [],
    });

    eventGroup = new formattingSettings.Group({
        name: "indicatorsEvent",
        displayName: "シナリオ",
        description: "メジャーを積んだ（横持ち）ときの、指標の値のシナリオ。既定は比較順の一番大きいメジャー。「すべてのシナリオ」はシナリオに依らない値（営業日数など）",
        slices: [],
        visible: false,
    });

    groups = [this.afterGroup, this.placementGroup, this.aggregationGroup, this.formatGroup, this.goodGroup, this.eventGroup];

    applyIndicators(settings: IndicatorSetting[], places: Array<{ value: string; displayName: string }>): void {
        this.visible = settings.length > 0;
        const selectorOf = (setting: IndicatorSetting) => ({ metadata: setting.queryName });
        this.afterGroup.slices = settings.map((setting) => {
            // 置く場所の選択肢は表の行と表の最後（自分の行は入れない）。表に無い保存値（フィルターで行が消えたなど）は選択肢に残す
            const items: powerbi.IEnumMember[] = [...places.filter((p) => p.value !== setting.code), { value: CALC_TABLE_END, displayName: "表の最後" }];
            if (!items.some((i) => i.value === setting.after)) items.push({ value: setting.after, displayName: `${missingLabel(setting.after)}（表に無い）` });
            return new formattingSettings.ItemDropdown({ name: "after", displayName: setting.name, items, value: itemOf(items, setting.after), selector: selectorOf(setting) });
        });
        this.placementGroup.slices = settings.map(
            (setting) =>
                new formattingSettings.ItemDropdown({
                    name: "placement",
                    displayName: setting.name,
                    items: INDICATOR_MODE_ITEMS,
                    value: itemOf(INDICATOR_MODE_ITEMS, setting.placement),
                    selector: selectorOf(setting),
                })
        );
        // うちは親の行に合わせるので、集計・書式・良し悪しの欄を出さない
        const own = settings.filter((setting) => setting.placement !== INDICATOR_MODES.under);
        this.aggregationGroup.visible = own.length > 0;
        this.aggregationGroup.slices = own.map(
            (setting) =>
                new formattingSettings.ItemDropdown({
                    name: "aggregation",
                    displayName: setting.name,
                    items: INDICATOR_AGGREGATION_ITEMS,
                    value: itemOf(INDICATOR_AGGREGATION_ITEMS, setting.aggregation),
                    selector: selectorOf(setting),
                })
        );
        this.formatGroup.visible = own.length > 0;
        this.formatGroup.slices = own.map(
            (setting) =>
                new formattingSettings.TextInput({ name: "format", displayName: setting.name, value: setting.format, placeholder: "#,0人・#,0.0h", selector: selectorOf(setting) })
        );
        this.goodGroup.visible = own.length > 0;
        this.goodGroup.slices = own.map(
            (setting) =>
                new formattingSettings.ItemDropdown({
                    name: "good",
                    displayName: setting.name,
                    items: CALC_GOOD_ITEMS,
                    value: itemOf(CALC_GOOD_ITEMS, setting.good),
                    selector: selectorOf(setting),
                })
        );
        // 横持ちだけ：指標の値のイベント
        const horizontal = settings.filter((setting) => setting.event !== null);
        this.eventGroup.visible = horizontal.length > 0;
        this.eventGroup.slices = horizontal.map((setting) => {
            const items: powerbi.IEnumMember[] = [...setting.eventChoices, { value: INDICATOR_ALL_EVENTS, displayName: "すべてのシナリオ" }];
            return new formattingSettings.ItemDropdown({ name: "event", displayName: setting.name, items, value: itemOf(items, setting.event!), selector: selectorOf(setting) });
        });
    }
}


/** 区分ごとの数値の枠の数（決め打ち。計算の行・イベントの金額の持ち方と同じ形） */
export const SECTION_NUMBER_SLOTS = 8;
/** 区分ごとの数値の枠で上書きする設定（名前は「数値」カードと同じ） */
export const SECTION_NUMBER_PROPS = ["section", "unitType", "precision", "negativeStyle", "zeroStyle", "diffZeroStyle", "negativeZero"] as const;
export type SectionNumberProp = (typeof SECTION_NUMBER_PROPS)[number];
/** 枠 i（1 から）の設定の名前。1 つの object（sectionNumbers）に section1・unitType1… と並べる */
export const sectionNumberProp = (prop: SectionNumberProp, slot: number): string => `${prop}${slot}`;
/** 区分ごとの数値の枠の保存値（生の文字。空は ""。表全体に合わせるは SECTION_TABLE） */
export type SectionNumberSaved = Record<SectionNumberProp, string>;
/** 「表全体に合わせる」の値 */
export const SECTION_TABLE = "table";
const TABLE_ITEM: powerbi.IEnumMember = { value: SECTION_TABLE, displayName: "表全体に合わせる" };
/** 区分で選べる表示単位は固定の単位だけ（自動は表全体の 1 つ。区分ごとに自動で変わると縦に比べられない） */
const SECTION_UNIT_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...UNIT_TYPES.filter((u) => u.value !== "auto")];
const SECTION_PRECISION_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...PRECISIONS.filter((p) => p.value !== "auto")];
const SECTION_NEGATIVE_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...NEGATIVE_STYLE_ITEMS];
const SECTION_ZERO_ITEMS: powerbi.IEnumMember[] = [TABLE_ITEM, ...ZERO_STYLE_ITEMS];
const SECTION_NEGATIVE_ZERO_ITEMS: powerbi.IEnumMember[] = [
    TABLE_ITEM,
    { value: "on", displayName: "残す（▲0）" },
    { value: "off", displayName: "残さない（0）" },
];

/** 区分ごとの数値の枠の、設定ごとの選択肢の値（「表全体に合わせる」を除く）。保存値が選択肢にあるかを見る */
export function sectionNumberAllowed(prop: SectionNumberProp): string[] {
    const items: Record<SectionNumberProp, powerbi.IEnumMember[]> = {
        section: [],
        unitType: SECTION_UNIT_ITEMS,
        precision: SECTION_PRECISION_ITEMS,
        negativeStyle: SECTION_NEGATIVE_ITEMS,
        zeroStyle: SECTION_ZERO_ITEMS,
        diffZeroStyle: SECTION_ZERO_ITEMS,
        negativeZero: SECTION_NEGATIVE_ZERO_ITEMS,
    };
    return items[prop].map((i) => String(i.value)).filter((v) => v !== SECTION_TABLE);
}

/** 区分ごとの数値の枠 1 本（書式ペインのコンテナーの項目） */
class SectionNumberItem extends SimpleCard {
    section: formattingSettings.ItemDropdown;
    unitType: formattingSettings.ItemDropdown;
    precision: formattingSettings.ItemDropdown;
    negativeStyle: formattingSettings.ItemDropdown;
    zeroStyle: formattingSettings.ItemDropdown;
    diffZeroStyle: formattingSettings.ItemDropdown;
    negativeZero: formattingSettings.ItemDropdown;

    constructor(readonly slot: number) {
        super();
        this.name = `sectionNumber${slot}`;
        // 表示名は固定（計算の行と同じ。名前を変えると Desktop で編集する枠の選択が 1 本目に戻った）
        this.displayName = `設定 ${slot}`;
        const dropdown = (prop: SectionNumberProp, displayName: string, items: powerbi.IEnumMember[], description?: string) =>
            new formattingSettings.ItemDropdown({ name: sectionNumberProp(prop, slot), displayName, description, items, value: items[0] });
        this.section = dropdown("section", "区分", [NO_ITEM]);
        this.unitType = dropdown("unitType", "表示単位", SECTION_UNIT_ITEMS, "表全体と違う単位にすると、区分の行の名前に単位を添える（「売上高（千円）」）");
        this.precision = dropdown("precision", "小数点以下の桁数", SECTION_PRECISION_ITEMS);
        this.negativeStyle = dropdown("negativeStyle", "マイナスの書き方", SECTION_NEGATIVE_ITEMS);
        this.zeroStyle = dropdown("zeroStyle", "0 の書き方", SECTION_ZERO_ITEMS);
        this.diffZeroStyle = dropdown("diffZeroStyle", "差が 0 の書き方", SECTION_ZERO_ITEMS);
        this.negativeZero = dropdown("negativeZero", "丸めて 0 のマイナスに符号を残す", SECTION_NEGATIVE_ZERO_ITEMS);
        this.slices = [this.section, this.unitType, this.precision, this.negativeStyle, this.zeroStyle, this.diffZeroStyle, this.negativeZero];
    }
}

/**
 * 区分ごとの数値（「数値」カードの設定を区分で上書きする）。区分の名前で覚える（区分は科目の属性で、節点の identity が無い）。
 * 効くのはその区分の小計・中分類の小計・科目・科目のうち。計算の行と指標の行は表全体のまま（区分をまたぐ行なので）
 */
export class SectionNumbersCardSettings extends formattingSettings.CompositeCard {
    name = "sectionNumbers";
    displayName = "区分ごとの数値";
    description = "「数値」カードの設定を区分ごとに上書きする。上書きしない設定は表全体に合わせる";
    visible = false;

    items = Array.from({ length: SECTION_NUMBER_SLOTS }, (_, i) => new SectionNumberItem(i + 1));
    group = new formattingSettings.Group({
        name: "sectionNumbersGroup",
        displayName: "区分ごと",
        description: "設定 1〜8 のそれぞれで、区分を 1 つ選び、その区分だけ変えるものを選ぶ。計算の行（段階利益・合計・比率）と指標の行は表全体のまま",
        slices: [],
        container: new formattingSettings.Container({ displayName: "編集する設定", containerItems: this.items }),
    });

    groups = [this.group];

    /** 区分の選択肢と保存値を流し込む。区分が無ければ出さない。選択肢に無い保存値は「（表に無い）」で残す */
    applySectionNumbers(sections: string[], saved: SectionNumberSaved[]): void {
        this.visible = sections.length > 0;
        const base: powerbi.IEnumMember[] = [NO_ITEM, ...sections.map((s) => ({ value: s, displayName: s }))];
        this.items.forEach((item, i) => {
            const slot = saved[i];
            const value = slot?.section ?? "";
            item.section.items = value === "" || sections.includes(value) ? base : [...base, { value, displayName: `${value}（表に無い）` }];
            item.section.value = itemOf(item.section.items, value);
            const set = (dropdown: formattingSettings.ItemDropdown, v: string | undefined) => (dropdown.value = itemOf(dropdown.items, v || SECTION_TABLE));
            set(item.unitType, slot?.unitType);
            set(item.precision, slot?.precision);
            set(item.negativeStyle, slot?.negativeStyle);
            set(item.zeroStyle, slot?.zeroStyle);
            set(item.diffZeroStyle, slot?.diffZeroStyle);
            set(item.negativeZero, slot?.negativeZero);
        });
    }
}

export class NumbersCardSettings extends SimpleCard {
    name = "numbers";
    displayName = "数値";

    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        description: "科目の行だけを割る（比率・時間・人数の行は割らない）。自動は、一番大きい値が 4 けた以上残る単位を表全体で 1 つ選ぶ",
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
        description: "科目の行の桁。自動は、単位で割った値が 100 未満なら 1 桁、ほかは 0 桁（表全体でそろえる）",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    currency = new formattingSettings.TextInput({
        name: "currency",
        displayName: "通貨の字",
        description: "右上の「単位：百万円」の「円」",
        value: DEFAULT_CURRENCY,
        placeholder: DEFAULT_CURRENCY,
    });

    negativeStyle = new formattingSettings.ItemDropdown({
        name: "negativeStyle",
        displayName: "マイナスの書き方",
        items: NEGATIVE_STYLE_ITEMS,
        value: itemOf(NEGATIVE_STYLE_ITEMS, NEGATIVE_STYLES.triangle),
    });

    zeroStyle = new formattingSettings.ItemDropdown({
        name: "zeroStyle",
        displayName: "0 の書き方",
        description: "値が 0（丸めて 0 を含む）のときの書き方",
        items: ZERO_STYLE_ITEMS,
        value: itemOf(ZERO_STYLE_ITEMS, ZERO_STYLES.zero),
    });

    diffZeroStyle = new formattingSettings.ItemDropdown({
        name: "diffZeroStyle",
        displayName: "差が 0 の書き方",
        description: "差・率が 0（変わらない）のときの書き方",
        items: ZERO_STYLE_ITEMS,
        value: itemOf(ZERO_STYLE_ITEMS, ZERO_STYLES.plusMinus),
    });

    negativeZero = new formattingSettings.ToggleSwitch({
        name: "negativeZero",
        displayName: "丸めて 0 のマイナスに符号を残す",
        description: "▲0 のように、丸めると 0 になるマイナスにも符号を付ける",
        value: true,
    });

    unitPlace = new formattingSettings.ItemDropdown({
        name: "unitPlace",
        displayName: "単位の置き場所",
        description: "「単位：百万円」を置く所。行の名前の横にすると、値の行は名前に（百万円）、台数・時間のような指標は数字から単位の字を外して名前に（台）を添える（比率の % は数字に残す）",
        items: UNIT_PLACE_ITEMS,
        value: UNIT_PLACE_ITEMS[0],
    });

    // options は付けない（素の numeric に付けると書式ペインが空になった記録がある）。範囲は viewModel でクランプする
    ratioCap = new formattingSettings.NumUpDown({
        name: "ratioCap",
        displayName: "比・率の上限（%）",
        description: "これを超える比・率は「≧999%」のように上限で止める（100〜99999）",
        value: DEFAULT_RATIO_CAP,
    });

    slices = [
        this.unitType,
        this.unitNotation,
        this.precision,
        this.currency,
        this.unitPlace,
        this.negativeStyle,
        this.zeroStyle,
        this.diffZeroStyle,
        this.negativeZero,
        this.ratioCap,
    ];
}

export class ColorsCardSettings extends SimpleCard {
    name = "colors";
    displayName = "色";

    toneMode = new formattingSettings.ItemDropdown({
        name: "toneMode",
        displayName: "差の色",
        description: "差の列を、良い向き（収益は増えて良い、費用は増えて悪い）で塗る",
        items: DIFF_TONE_MODE_ITEMS,
        value: DIFF_TONE_MODE_ITEMS.find((i) => i.value === TONE_MODES.both)!,
    });

    good = new formattingSettings.ColorPicker({
        name: "good",
        displayName: "良い差",
        value: { value: DEFAULT_GOOD_COLOR },
    });

    bad = new formattingSettings.ColorPicker({
        name: "bad",
        displayName: "悪い差",
        value: { value: DEFAULT_BAD_COLOR },
    });

    headBackground = new formattingSettings.ColorPicker({
        name: "headBackground",
        displayName: "見出しの背景",
        description: "期間と列の見出し、左上の角、セグメントの列の見出し",
        value: { value: DEFAULT_HEAD_BACKGROUND },
    });

    slices = [this.toneMode, this.good, this.bad, this.headBackground];
}

export class LinesCardSettings extends SimpleCard {
    name = "lines";
    displayName = "罫線";

    row = new formattingSettings.ColorPicker({ name: "row", displayName: "行の間の線", value: { value: DEFAULT_LINE_COLORS.row } });
    subtotal = new formattingSettings.ColorPicker({ name: "subtotal", displayName: "小計の上の線", value: { value: DEFAULT_LINE_COLORS.subtotal } });
    total = new formattingSettings.ColorPicker({
        name: "total",
        displayName: "段階の上の二重線",
        description: "売上総利益・営業利益・資産合計など、計算の行の小計の上",
        value: { value: DEFAULT_LINE_COLORS.total },
    });
    head = new formattingSettings.ColorPicker({ name: "head", displayName: "見出しの下の線", value: { value: DEFAULT_LINE_COLORS.head } });
    block = new formattingSettings.ColorPicker({
        name: "block",
        displayName: "セグメントの箱の線",
        description: "セグメントの段ごとの縦線と、セグメントの切れ目の横線",
        value: { value: DEFAULT_LINE_COLORS.block },
    });
    band = new formattingSettings.ColorPicker({
        name: "band",
        displayName: "囲みの線",
        description: "区分・中分類の囲み（箱）の段ごとの縦線。箱の横の辺は行の間の線",
        value: { value: DEFAULT_LINE_COLORS.band },
    });
    outer = new formattingSettings.ColorPicker({ name: "outer", displayName: "表の外枠", value: { value: DEFAULT_LINE_COLORS.outer } });
    periods = new formattingSettings.ToggleSwitch({ name: "periods", displayName: "期間の区切り", description: "期間のあいだに縦の線を引く", value: true });
    period = new formattingSettings.ColorPicker({ name: "period", displayName: "期間の区切りの色", value: { value: DEFAULT_LINE_COLORS.period } });
    nameEdge = new formattingSettings.ColorPicker({
        name: "nameEdge",
        displayName: "左の列の区切りの線",
        description: "セグメントの列と行の名前の列の右の縦線（見出しの左上の角も）",
        value: { value: DEFAULT_LINE_COLORS.name },
    });

    slices = [this.row, this.subtotal, this.total, this.head, this.block, this.band, this.nameEdge, this.outer, this.periods, this.period];
}

/** 表のほかの見せ方（今は「画像としてコピー」のボタンだけ） */
export class DisplayCardSettings extends SimpleCard {
    name = "display";
    displayName = "表示";

    copyButton = new formattingSettings.ToggleSwitch({
        name: "copyButton",
        displayName: "画像のコピー",
        description: "マウスを乗せたとき、右下に「画像としてコピー」のボタンを出す。押すと、見えている表を画像としてクリップボードに入れ、PowerPoint などに貼れる",
        value: true,
    });

    slices = [this.copyButton];
}

export class TextCardSettings extends SimpleCard {
    name = "text";
    displayName = "文字";

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "フォント",
        value: DEFAULT_FONT_FAMILY,
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "表の上の文字サイズ",
        description: "表の上のバー（題名・メニュー・単位）とダイアログ。表の中は下の項目で要素ごとに変える",
        value: DEFAULT_FONT_SIZE,
    });

    orgSize = new formattingSettings.NumUpDown({ name: "orgSize", displayName: "セグメントの名前", value: DEFAULT_TEXT_SIZES.org });
    nameSize = new formattingSettings.NumUpDown({ name: "nameSize", displayName: "行の名前", value: DEFAULT_TEXT_SIZES.name });
    periodSize = new formattingSettings.NumUpDown({ name: "periodSize", displayName: "期間の見出し", value: DEFAULT_TEXT_SIZES.period });
    // 主の列と比較の列の見出しは別々に変える。保存の名前 columnSize は主の列の見出し
    columnSize = new formattingSettings.NumUpDown({
        name: "columnSize",
        displayName: "主の列の見出し",
        description: "実績・最新見込みなどの列の見出しと、左上の角に置いた単位",
        value: DEFAULT_TEXT_SIZES.column,
    });
    compareHeadSize = new formattingSettings.NumUpDown({
        name: "compareHeadSize",
        displayName: "比較の列の見出し",
        description: "期初予算差・見通し比などの列の見出し（2 行なら 2 行とも）",
        value: DEFAULT_TEXT_SIZES.compareHead,
    });
    mainSize = new formattingSettings.NumUpDown({ name: "mainSize", displayName: "主の数字", value: DEFAULT_TEXT_SIZES.main });
    compareSize = new formattingSettings.NumUpDown({
        name: "compareSize",
        displayName: "比較の数字",
        description: "比較の列（比較の値・差・比・率）。2 段なら上の段",
        value: DEFAULT_TEXT_SIZES.compare,
    });
    subSize = new formattingSettings.NumUpDown({ name: "subSize", displayName: "2 段目の数字", value: DEFAULT_TEXT_SIZES.sub });

    slices = [this.fontFamily, this.fontSize, this.orgSize, this.nameSize, this.periodSize, this.columnSize, this.compareHeadSize, this.mainSize, this.compareSize, this.subSize];
}

/** update() のたびにデータから流し込むもの */
export interface DataDrivenFormat {
    hierarchy?: HierarchyFormat;
    mainItems: Array<{ value: string; displayName: string }>;
    main: string;
    /** 横持ちのメジャーごとの比較順と金額の持ち方（縦持ちなら空） */
    measureSettings: MeasureSetting[];
    /** 縦持ちのイベントごとの金額の持ち方：出すか、イベントの選択肢、枠の保存値（生） */
    eventSigns?: { visible: boolean; events: string[]; saved: EventSignSaved[] };
    years: Array<{ value: string; displayName: string }>;
    fiscalYear: string;
    /** 計算の行：置く場所・分子・分母の選択肢、書式ペインの保存値（生） */
    calcChoices: CalcChoices;
    calcSaved: CalcSaved[];
    /** 指標のメジャーごとの設定（置く場所の選択肢は calcChoices.places） */
    indicatorSettings: IndicatorSetting[];
    /** 区分ごとの数値：区分の選択肢（表の並び）と、枠の保存値（生） */
    sectionNumbers?: { sections: string[]; saved: SectionNumberSaved[] };
    /** 組織の欄を入れたか（組織の設定を出す）と、一番上の段が 2 つ以上で全社を置いたか（全社の名前を出す） */
    hasOrgs?: boolean;
    hasRoot?: boolean;
}

export class VisualFormattingSettingsModel extends Model {
    hierarchy = new HierarchySettings();
    hierarchyOverrides = new HierarchyOverrides();
    comparison = new ComparisonCardSettings();
    events = new EventsCardSettings();
    eventSigns = new EventSignsCardSettings();
    periods = new PeriodsCardSettings();
    rows = new RowsCardSettings();
    calcRows = new CalcRowsCardSettings();
    indicators = new IndicatorsCardSettings();
    numbers = new NumbersCardSettings();
    sectionNumbers = new SectionNumbersCardSettings();
    colors = new ColorsCardSettings();
    lines = new LinesCardSettings();
    text = new TextCardSettings();

    display = new DisplayCardSettings();

    cards = [
        this.comparison,
        this.events,
        this.eventSigns,
        this.periods,
        this.rows,
        this.hierarchy,
        this.hierarchyOverrides,
        this.calcRows,
        this.indicators,
        this.numbers,
        this.sectionNumbers,
        this.colors,
        this.lines,
        this.text,
        this.display,
    ];

    /**
     * 書式ペインの色のうち、保存値の無いもの（saved が false）の既定をテーマの色にする。書式ペインにもテーマの色が出る。
     * populateFormattingSettingsModel のあと、表を組む前に呼ぶ
     */
    applyTheme(theme: Theme, saved: (object: string, property: string) => boolean): void {
        const set = (object: string, picker: formattingSettings.ColorPicker, value: string) => {
            if (!saved(object, picker.name)) picker.value = { value };
        };
        set(this.colors.name, this.colors.good, theme.good);
        set(this.colors.name, this.colors.bad, theme.bad);
        set(this.colors.name, this.colors.headBackground, theme.headBackground);
        const lines = this.lines;
        set(lines.name, lines.row, theme.lines.row);
        set(lines.name, lines.subtotal, theme.lines.subtotal);
        set(lines.name, lines.total, theme.lines.total);
        set(lines.name, lines.head, theme.lines.head);
        set(lines.name, lines.block, theme.lines.block);
        set(lines.name, lines.band, theme.lines.band);
        set(lines.name, lines.outer, theme.lines.outer);
        set(lines.name, lines.period, theme.lines.period);
        set(lines.name, lines.nameEdge, theme.lines.name);
        const rows = this.rows;
        [rows.bandColor, rows.bandColor2, rows.bandColor3, rows.bandColor4].forEach((picker, i) => set(rows.name, picker, theme.bandColors[i]));
        [rows.segmentColor, rows.segmentColor2, rows.segmentColor3, rows.segmentColor4].forEach((picker, i) => set(rows.name, picker, theme.segmentColors[i]));
    }

    applyData(data: DataDrivenFormat): void {
        this.hierarchyOverrides.visible = this.hierarchy.enabled.value;
        this.hierarchyOverrides.apply(data.hierarchy);
        this.comparison.applyEvents(data.mainItems, data.main);
        this.events.applyMeasures(data.measureSettings ?? []);
        this.eventSigns.applyEventSigns(data.eventSigns?.visible ?? false, data.eventSigns?.events ?? [], data.eventSigns?.saved ?? []);
        this.periods.applyYears(data.years, data.fiscalYear);
        this.calcRows.applyCalc(data.calcChoices ?? { sections: [], refs: [], places: [] }, data.calcSaved ?? []);
        this.indicators.applyIndicators(data.indicatorSettings ?? [], data.calcChoices?.places ?? []);
        this.rows.applyOrgs(data.hasOrgs ?? false, data.hasRoot ?? false);
        this.sectionNumbers.applySectionNumbers(data.sectionNumbers?.sections ?? [], data.sectionNumbers?.saved ?? []);
    }
}
