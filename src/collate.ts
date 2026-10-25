/**
 * 名前・コードの並べ替え。数字は数として比べる（第2回 < 第10回、4100-2 < 4100-10）。並びは localeCompare(b, "ja", { numeric: true }) と同じ。
 * localeCompare は比べるたびに照合の書式を作り直すので、作った Intl.Collator を使い回す
 */
const JA_NUMERIC = new Intl.Collator("ja", { numeric: true });

export const compareJa = (a: string, b: string): number => JA_NUMERIC.compare(a, b);
