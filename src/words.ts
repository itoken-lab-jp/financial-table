/**
 * ビジュアルが決めている言葉（最新値・前年同期・差・比・率・期間の名前・年度・単位の見出し）。書式ペインの「用語」カードで書き換える。
 * 空ならここの既定。四半期は {n}（1〜4）、年度は {年}（4 桁）・{年2}（下 2 桁）の所に数字を入れる
 */
export interface Words {
    latest: string;
    priorYear: string;
    diff: string;
    ratio: string;
    rate: string;
    quarter: string;
    firstHalf: string;
    secondHalf: string;
    fullYear: string;
    ytd: string;
    sum: string;
    allPeriods: string;
    /** 年度の形。空なら「年度の表記」の選択（2025年度・FY2026） */
    year: string;
    unit: string;
}

export const DEFAULT_WORDS: Words = {
    latest: "最新値",
    priorYear: "前年同期",
    diff: "差",
    ratio: "比",
    rate: "率",
    quarter: "{n}Q",
    firstHalf: "上期",
    secondHalf: "下期",
    fullYear: "通期",
    ytd: "累計",
    sum: "計",
    allPeriods: "全期間",
    year: "",
    unit: "単位",
};

export const WORD_KEYS = Object.keys(DEFAULT_WORDS) as Array<keyof Words>;

/** 書式ペインの字（空・空白だけは既定） */
export function wordsOf(read: (key: keyof Words) => string | undefined | null): Words {
    return Object.fromEntries(WORD_KEYS.map((key) => [key, read(key)?.trim() || DEFAULT_WORDS[key]])) as unknown as Words;
}

/** {n} を埋める（四半期の番号）。印が無ければ数を付けない（「第1四半期」と書けば 4 つとも同じ字になるので、印を書くよう説明に書く） */
export function fillWord(template: string, mark: "{n}", value: number): string {
    return template.split(mark).join(String(value));
}

/** 年度の形の {年2}（下 2 桁、FY27）と {年}（4 桁）を埋める */
export function fillYear(template: string, year: number): string {
    return template.split("{年2}").join(String(year % 100).padStart(2, "0")).split("{年}").join(String(year));
}
