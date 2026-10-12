/**
 * レポートのテーマの色。
 *
 * Power BI がビジュアルに渡すテーマの色（ISandboxExtendedColorPalette）から、表の既定の色を決める。書式ペインで色を決めていれば
 * （保存値があれば）その色のまま（VisualFormattingSettingsModel.applyTheme）。字・副文字・地は書式ペインに無いので、いつもテーマの色。
 * 見出しの地と囲みの段は、テーマの「背景（明るい）」「背景（中間）」でなく、テーマの字と地から作る（基本テーマ Fluent 2 では
 * その 2 つがどちらも濃い灰 #C8C6C4・#C7C7C7 で、見出しと囲みがすべて濃く塗られた）。テーマが無いときと
 * ハイコントラストのとき（Power BI が前景と背景を決める）は DEFAULT_THEME
 */
import powerbi from "powerbi-visuals-api";

import { blend, readableKeepingHue } from "./shared/color";
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
 * 良い差・悪い差の字が地に対して読める濃さ（4.5:1。小さい字の読みやすさの目安、WCAG AA）。足りなければ、テーマの良い・悪いの色の
 * 色相と彩度を保ったまま暗くする（Fluent 2 の淡い緑 #9FD89F → #378637、淡い赤 #EEACB2 → #D73C4A）。字の色へ混ぜると灰に近づき、
 * 黒い数字と見分けにくくなった（#9FD89F が #5B755B に、#EEACB2 が #89686B になった）
 */
export const SENTIMENT_CONTRAST = 4.5;

export const DEFAULT_THEME: Theme = {
    text: "#252423",
    muted: "#605E5C",
    background: "#FFFFFF",
    headBackground: DEFAULT_HEAD_BACKGROUND,
    good: readableKeepingHue(DEFAULT_GOOD_COLOR, "#FFFFFF", SENTIMENT_CONTRAST, "#252423"),
    bad: readableKeepingHue(DEFAULT_BAD_COLOR, "#FFFFFF", SENTIMENT_CONTRAST, "#252423"),
    lines: DEFAULT_LINE_COLORS,
    bandColors: DEFAULT_BAND_COLORS,
    segmentColors: DEFAULT_SEGMENT_COLORS,
};

/**
 * 見出しの地と、囲みの段の両端。テーマの地から字へ寄せる量（地 #FFFFFF・字 #242424 なら見出しの地 #F5F5F5、濃い端 #C9C9C9）。
 * 暗い地のテーマでも同じ段になる。標準の表とマトリックスは見出しを塗らないので、見出しの地はごく薄くする
 */
const HEAD_STEP = 0.045;
const NEUTRAL_STEP = 0.25;

/**
 * 囲み・セグメントの塗りの段ごとの濃さ（外ほど濃い）。見出しの地から濃い端へ寄せる量。
 * データの色（系列の色）は使わない（表の色は地と字のあいだの灰の段で作る）。塗るのはまとまりを締める行だけ（App.tsx）
 */
const BAND_STEPS = [0.45, 0.3, 0.15, 0];
/** セグメントの塗りは、囲みよりひと段濃く */
const SEGMENT_STEPS = [0.75, 0.55, 0.35, 0.15];


/** テーマの色から表の色を決める。無いか読めない色は DEFAULT_THEME の同じ所 */
export function themeOf(palette: ISandboxExtendedColorPalette | undefined): Theme {
    if (!palette || palette.isHighContrast) return DEFAULT_THEME;
    const valid = (info: { value?: string } | undefined, fallback: string) => (info?.value && /^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/i.test(info.value) ? info.value : fallback);
    const text = valid(palette.foreground, DEFAULT_THEME.text);
    const background = valid(palette.background, DEFAULT_THEME.background);
    const muted = valid(palette.foregroundNeutralSecondary, DEFAULT_THEME.muted);
    const soft = blend(background, "#000000", 0.1);
    const light = blend(background, text, HEAD_STEP);
    const neutral = blend(background, text, NEUTRAL_STEP);
    const steps = (amounts: number[]) => amounts.map((amount) => blend(light, neutral, amount));
    return {
        text,
        muted,
        background,
        headBackground: light,
        good: readableKeepingHue(valid(palette.positive, DEFAULT_GOOD_COLOR), background, SENTIMENT_CONTRAST, text),
        bad: readableKeepingHue(valid(palette.negative, DEFAULT_BAD_COLOR), background, SENTIMENT_CONTRAST, text),
        // 線の既定（2026-10-08 ユーザーが Desktop で決めた色をテーマの言葉で）：行の中の線はテーマの背景を 10% 暗く
        // （書式ペインのテーマの色の「白、10% 暗く」）、セグメントの箱・期間の区切り・外枠はテーマの文字の色。見出しの下の線は文字の色のまま
        lines: {
            row: soft,
            subtotal: soft,
            total: soft,
            head: text,
            block: text,
            band: soft,
            period: text,
            outer: text,
            name: soft,
        },
        bandColors: steps(BAND_STEPS),
        segmentColors: steps(SEGMENT_STEPS),
    };
}
