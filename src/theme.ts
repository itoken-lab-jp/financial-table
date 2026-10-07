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
import { DEFAULT_BAD_COLOR, DEFAULT_BAND_COLORS, DEFAULT_GOOD_COLOR, DEFAULT_HEAD_BACKGROUND, DEFAULT_LINE_COLORS, DEFAULT_NEUTRAL_BACKGROUND, DEFAULT_SEGMENT_COLORS } from "./settings";

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
 * 囲み・セグメントの塗りの段ごとの濃さ（外ほど濃い）。テーマの「背景（明るい）」（見出しの背景と同じ）から「背景（中間）」へ寄せる量。
 * データの色（系列の色）は使わない（2026-10-08 ユーザーと決めた。表の色はテーマの背景の段で作る）。塗るのはまとまりを締める行だけ（App.tsx）
 */
const BAND_STEPS = [0.45, 0.3, 0.15, 0];
/** セグメントの塗りは、囲みよりひと段濃く */
const SEGMENT_STEPS = [0.75, 0.55, 0.35, 0.15];


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
    const soft = blend(background, "#000000", 0.1);
    const light = valid(palette.backgroundLight, DEFAULT_HEAD_BACKGROUND);
    const neutral = valid(palette.backgroundNeutral, DEFAULT_NEUTRAL_BACKGROUND);
    const steps = (amounts: number[]) => amounts.map((amount) => blend(light, neutral, amount));
    return {
        text,
        muted,
        background,
        headBackground: valid(palette.backgroundLight, DEFAULT_HEAD_BACKGROUND),
        good: readableOn(valid(palette.positive, DEFAULT_GOOD_COLOR), background, text, SENTIMENT_CONTRAST),
        bad: readableOn(valid(palette.negative, DEFAULT_BAD_COLOR), background, text, SENTIMENT_CONTRAST),
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
