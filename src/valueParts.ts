export interface ValueParts { number: string; suffix: string; closing: string }

/** 書式が持つ接尾辞を明示して分ける。単位の中の数字や記号を推測で削らない。 */
export function valueParts(text: string, suffixes: string[]): ValueParts {
    const closing = text.endsWith(")") && text.includes("(") ? ")" : "";
    const body = closing ? text.slice(0, -1) : text;
    const suffix = suffixes.filter(Boolean).sort((a, b) => b.length - a.length).find(s => body.endsWith(s)) ?? "";
    return { number: suffix ? body.slice(0, -suffix.length) : body, suffix, closing };
}
