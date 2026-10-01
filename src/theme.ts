/**
 * レポートのテーマの色。
 *
 * Power BI がビジュアルに渡すテーマの色（ISandboxExtendedColorPalette）から、表の既定の色を決める。書式ペインで色を決めていれば
 * （保存値があれば）その色のまま（VisualFormattingSettingsModel.applyTheme）。字・副文字・地は書式ペインに無いので、いつもテーマの色。
 * 今の既定の色は Power BI の標準テーマの値なので、標準テーマのレポートでは見た目が変わらない。テーマが無いときと
 * ハイコントラストのとき（Power BI が前景と背景を決める）は DEFAULT_THEME
 */
import powerbi from "powerbi-visuals-api";

import { blend, contrastRatio } from "./shared/color";
import { DEFAULT_BAD_COLOR, DEFAULT_BAND_COLORS, DEFAULT_GOOD_COLOR, DEFAULT_HEAD_BACKGROUND, DEFAULT_LINE_COLORS, DEFAULT_SEGMENT_COLORS } from "./settings";

import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;

export interface Theme {
    /** 字・副文字（単位・2 行目の見出し）・地 */
    text: string;
    muted: string;
    background: string;
    /** 以下は書式ペインの色の既定 */
    headBackground: string;
    good: string;
    bad: string;
    lines: typeof DEFAULT_LINE_COLORS;
    bandColors: string[];
    segmentColors: string[];
}

/**
 * 良い差・悪い差の字が地に対して読める濃さ（4.5:1。小さい字の読みやすさの目安、WCAG AA）。足りなければ字の色へ寄せて暗くする
 * （標準テーマの緑 #1AAB40 は 3.0:1 → #1D8337、記事用のテーマの黄 #E3B42F は 1.9:1 → #827131）。
 * 「4.5:1 まで暗く」
 */
export const SENTIMENT_CONTRAST = 4.5;

export const DEFAULT_THEME: Theme = {
    text: "#252423",
    muted: "#605E5C",
    background: "#FFFFFF",
    headBackground: DEFAULT_HEAD_BACKGROUND,
    good: readableOn(DEFAULT_GOOD_COLOR, "#FFFFFF", "#252423", SENTIMENT_CONTRAST),
    bad: readableOn(DEFAULT_BAD_COLOR, "#FFFFFF", "#252423", SENTIMENT_CONTRAST),
    lines: DEFAULT_LINE_COLORS,
    bandColors: DEFAULT_BAND_COLORS,
    segmentColors: DEFAULT_SEGMENT_COLORS,
};

/**
 * 塗りの段ごとの、地に寄せる量（外ほど濃い）。標準テーマの 1 番目のデータの色 #118DFF から、既定の囲みの色 #D4EAFF 系になる。
 * 塗るのはまとまりを締める行だけ（App.tsx）。「囲みの塗りを薄く」
 */
const TINTS = [0.82, 0.9, 0.95, 0.97];
/** セグメントの塗りの寄せる量。標準テーマの 3 番目のデータの色 #E66C37 から、今の既定のセグメントの色 #F4B183 系に近くなる */
const SEGMENT_TINTS = [0.45, 0.65, 0.85, 0.93];


/** 字に使う色を、地に対して minimum のコントラスト比になるまで字の色へ寄せる（足りていればそのまま） */
export function readableOn(color: string, background: string, text: string, minimum: number): string {
    for (let amount = 0; amount <= 1; amount += 0.05) {
        const candidate = amount === 0 ? color : blend(color, text, amount);
        if (contrastRatio(candidate, background) >= minimum) return candidate;
    }
    return text;
}

/** テーマの色から表の色を決める。無いか読めない色は DEFAULT_THEME の同じ所 */
export function themeOf(palette: ISandboxExtendedColorPalette | undefined): Theme {
    if (!palette || palette.isHighContrast) return DEFAULT_THEME;
    const valid = (info: { value?: string } | undefined, fallback: string) => (info?.value && /^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/i.test(info.value) ? info.value : fallback);
    const text = valid(palette.foreground, DEFAULT_THEME.text);
    const background = valid(palette.background, DEFAULT_THEME.background);
    const muted = valid(palette.foregroundNeutralSecondary, DEFAULT_THEME.muted);
    const line = valid(palette.backgroundNeutral, DEFAULT_LINE_COLORS.row);
    // データの色は getColor で順に取る（初めての名前には、テーマのデータの色が並びの順に割り当たる。同じ名前はいつも同じ色）
    const data = ["financialTable.theme.0", "financialTable.theme.1", "financialTable.theme.2"].map((key) => valid(palette.getColor(key), ""));
    const tints = (base: string, amounts: number[], fallback: string[]) => (base ? amounts.map((amount) => blend(base, background, amount)) : fallback);
    return {
        text,
        muted,
        background,
        headBackground: valid(palette.backgroundLight, DEFAULT_HEAD_BACKGROUND),
        good: readableOn(valid(palette.positive, DEFAULT_GOOD_COLOR), background, text, SENTIMENT_CONTRAST),
        bad: readableOn(valid(palette.negative, DEFAULT_BAD_COLOR), background, text, SENTIMENT_CONTRAST),
        lines: {
            row: line,
            subtotal: muted,
            total: text,
            head: text,
            // セグメントの箱・表の外枠は字と地のあいだ（標準テーマで #8A8886 あたり）
            block: blend(text, background, 0.46),
            band: line,
            // 期間の区切りは行の線より薄く（標準テーマで #E1DFDD あたり）
            period: blend(line, background, 0.44),
            outer: blend(text, background, 0.46),
            name: line,
        },
        bandColors: tints(data[0], TINTS, DEFAULT_BAND_COLORS),
        segmentColors: tints(data[2], SEGMENT_TINTS, DEFAULT_SEGMENT_COLORS),
    };
}
