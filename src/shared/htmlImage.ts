/**
 * HTML で描いた表などを画像にする（copyImage の chartImage の HTML 版。CopyImageButton の capture に渡す）。
 *
 * 入れ物の中の見えている要素を、画面の位置のまま canvas に描く：地の色、枠線（実線・点線・破線・二重線）、字。SVG は chartImage と同じく描く。
 * スクロールする箱の中は見えている範囲で切る。重なりの順は、位置を持つ要素（position が static でない）とその z-index で決める
 * （固定した見出し・列が、スクロールした中身の上に来る）。描かないもの：疑似要素（::before・::after）、影、角の丸み、背景の画像、
 * 字の飾り（下線など）、フォームの部品の中身（select の選んだ字など）。写したくない所は SKIP_ATTRIBUTE を付ける。
 * 画像の大きさは、描いたものの見えている範囲を合わせた四角（入れ物の余白は切る）
 */
import { ChartImage, IMAGE_SCALE, Rect, SKIP_ATTRIBUTE, dataUrlWithResolution, drawSvg, intersect, rectOf, visibleRect } from "./copyImage";

/** 重なりの順の鍵：位置を持つ祖先（と自分）ごとに [z-index, 1]。自分が位置を持たなければ最後に [0, 0] */
export type Layer = Array<[number, number]>;

/**
 * 重なりの順を比べる（小さいほうを先に描く）。鍵が前の部分で同じなら短いほう（入れ物）が先。同じ入れ物の中では、位置を持たない要素、
 * z-index の小さい順（同じなら位置を持たないほうが先）
 */
export function compareLayers(a: Layer, b: Layer): number {
    const n = Math.max(a.length, b.length);
    for (let i = 0; i < n; i++) {
        const x = a[i];
        const y = b[i];
        if (!x) return -1;
        if (!y) return 1;
        if (x[0] !== y[0]) return x[0] - y[0];
        if (x[1] !== y[1]) return x[1] - y[1];
    }
    return 0;
}

/** 色が見えないか（透明・アルファ 0） */
export function transparent(color: string): boolean {
    const c = color.trim();
    return c === "" || c === "transparent" || /^rgba\(.*,\s*0(\.0+)?\s*\)$/.test(c) || /\/\s*0(\.0+)?\s*\)$/.test(c);
}

/** 1 本の枠線を描く四角（点線・破線は線、ほかは塗り）。二重線は外と内の 2 本 */
export interface BorderStroke {
    kind: "fill" | "dash";
    /** fill：塗る四角。dash：線の始まりと終わり（中心線）と太さ・破線の長さ */
    rect?: Rect;
    from?: [number, number];
    to?: [number, number];
    width?: number;
    dash?: [number, number];
}

type Side = "top" | "right" | "bottom" | "left";

/** 枠線の 1 辺を、描く形に直す（幅・線の種類・要素の四角から） */
export function borderStrokes(side: Side, width: number, style: string, box: Rect): BorderStroke[] {
    if (width <= 0 || style === "none" || style === "hidden") return [];
    const horizontal = side === "top" || side === "bottom";
    // 辺の四角（外側から width の太さ）
    const strip = (w: number, inset = 0): Rect => {
        if (side === "top") return { left: box.left, right: box.right, top: box.top + inset, bottom: box.top + inset + w };
        if (side === "bottom") return { left: box.left, right: box.right, top: box.bottom - inset - w, bottom: box.bottom - inset };
        if (side === "left") return { top: box.top, bottom: box.bottom, left: box.left + inset, right: box.left + inset + w };
        return { top: box.top, bottom: box.bottom, left: box.right - inset - w, right: box.right - inset };
    };
    if (style === "double" && width >= 3) {
        const line = Math.max(1, Math.round(width / 3));
        return [
            { kind: "fill", rect: strip(line) },
            { kind: "fill", rect: strip(line, width - line) },
        ];
    }
    if (style === "dotted" || style === "dashed") {
        const r = strip(width);
        const from: [number, number] = horizontal ? [r.left, (r.top + r.bottom) / 2] : [(r.left + r.right) / 2, r.top];
        const to: [number, number] = horizontal ? [r.right, (r.top + r.bottom) / 2] : [(r.left + r.right) / 2, r.bottom];
        const length = style === "dotted" ? width : width * 3;
        return [{ kind: "dash", from, to, width, dash: [length, length] }];
    }
    return [{ kind: "fill", rect: strip(width) }];
}

interface Item {
    layer: Layer;
    /** 同じ重なりの中では、地と枠線（0）を字（1）より先に描く */
    kind: 0 | 1;
    order: number;
    paint: (ctx: CanvasRenderingContext2D, origin: Rect) => Promise<void> | void;
}

/** 要素の重なりの鍵（root から要素まで） */
function layerOf(el: Element, root: HTMLElement, cache: Map<Element, Layer>): Layer {
    const cached = cache.get(el);
    if (cached) return cached;
    const parent = el === root || !el.parentElement ? [] : layerOf(el.parentElement, root, cache).filter((part) => part[1] === 1);
    const style = getComputedStyle(el);
    const positioned = style.position !== "static";
    const z = Number.parseInt(style.zIndex, 10);
    const layer: Layer = [...parent, positioned ? [Number.isFinite(z) ? z : 0, 1] : [0, 0]];
    cache.set(el, layer);
    return layer;
}

/** 祖先の opacity を掛けたもの */
function opacityOf(el: Element, root: HTMLElement): number {
    let value = 1;
    for (let node: Element | null = el; node; node = node.parentElement) {
        value *= Number.parseFloat(getComputedStyle(node).opacity) || 0;
        if (node === root) break;
    }
    return value;
}

function clipTo(ctx: CanvasRenderingContext2D, clip: Rect, origin: Rect): void {
    ctx.beginPath();
    ctx.rect(clip.left - origin.left, clip.top - origin.top, clip.right - clip.left, clip.bottom - clip.top);
    ctx.clip();
}

/** 字の 1 行を、画面の四角（字の行の箱）の中に、同じフォントで描く */
function drawLine(ctx: CanvasRenderingContext2D, text: string, rect: Rect, origin: Rect): void {
    const metrics = ctx.measureText(text);
    const ascent = metrics.fontBoundingBoxAscent || metrics.actualBoundingBoxAscent;
    const descent = metrics.fontBoundingBoxDescent || metrics.actualBoundingBoxDescent;
    const height = rect.bottom - rect.top;
    ctx.fillText(text, rect.left - origin.left, rect.top + (height - (ascent + descent)) / 2 + ascent - origin.top);
}

/** 字の節の、画面の行ごとの字と四角（折り返していれば 1 字ずつ測って行に分ける） */
function linesOf(node: Text): Array<{ text: string; rect: Rect }> {
    const range = document.createRange();
    range.selectNodeContents(node);
    const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
    if (rects.length === 0) return [];
    const collapse = (s: string) => s.replace(/[\s ]+/g, " ").trim();
    if (rects.length === 1) {
        const text = collapse(node.data);
        const r = rects[0];
        return text ? [{ text, rect: { left: r.left, top: r.top, right: r.right, bottom: r.bottom } }] : [];
    }
    const lines: Array<{ text: string; rect: Rect }> = [];
    for (let i = 0; i < node.data.length; i++) {
        range.setStart(node, i);
        range.setEnd(node, i + 1);
        const r = range.getClientRects()[0];
        if (!r || r.width === 0) continue;
        const last = lines[lines.length - 1];
        if (last && Math.abs(last.rect.top - r.top) < 1) {
            last.text += node.data[i];
            last.rect = { ...last.rect, right: Math.max(last.rect.right, r.right) };
        } else {
            lines.push({ text: node.data[i], rect: { left: r.left, top: r.top, right: r.right, bottom: r.bottom } });
        }
    }
    return lines.map((line) => ({ ...line, text: collapse(line.text) })).filter((line) => line.text);
}

/** root の中の見えている HTML（と SVG）を画像にする。skip は写さない要素のセレクター（SKIP_ATTRIBUTE のほかに）。描くものが無ければ null */
export async function htmlImage(root: HTMLElement, background = "#ffffff", skip = ""): Promise<ChartImage | null> {
    const layers = new Map<Element, Layer>();
    const items: Item[] = [];
    let bounds: Rect | null = null;
    const grow = (r: Rect) => {
        bounds = bounds ? { left: Math.min(bounds.left, r.left), top: Math.min(bounds.top, r.top), right: Math.max(bounds.right, r.right), bottom: Math.max(bounds.bottom, r.bottom) } : r;
    };
    const skipSelector = [`[${SKIP_ATTRIBUTE}]`, skip].filter(Boolean).join(", ");
    const skipped = (el: Element) => el.closest(skipSelector) !== null;
    let order = 0;
    const elements = [root, ...Array.from(root.querySelectorAll("*"))];
    for (const el of elements) {
        order++;
        // SVG の中の要素は SVG ごと描く
        if (skipped(el) || el.parentElement?.closest("svg")) continue;
        if (el instanceof SVGSVGElement) {
            const visible = visibleRect(el, root);
            if (!visible) continue;
            grow(visible);
            items.push({ layer: layerOf(el, root, layers), kind: 0, order, paint: (ctx, origin) => drawSvg(ctx, el, root, origin) });
            continue;
        }
        if (!(el instanceof HTMLElement)) continue;
        const style = getComputedStyle(el);
        if (style.display === "none" || style.visibility === "hidden") continue;
        const clip = visibleRect(el, root);
        if (!clip) continue;
        const box = rectOf(el);
        const fill = transparent(style.backgroundColor) ? null : style.backgroundColor;
        const strokes = (["top", "right", "bottom", "left"] as Side[]).flatMap((side) => {
            const color = style.getPropertyValue(`border-${side}-color`);
            if (transparent(color)) return [];
            const width = Number.parseFloat(style.getPropertyValue(`border-${side}-width`)) || 0;
            return borderStrokes(side, width, style.getPropertyValue(`border-${side}-style`), box).map((stroke) => ({ stroke, color }));
        });
        const layer = layerOf(el, root, layers);
        if (fill || strokes.length > 0) {
            // 入れ物そのものの地は背景で塗るので、写す範囲には数えない（余白まで写さない）
            if (el !== root) grow(clip);
            const alpha = opacityOf(el, root);
            items.push({
                layer,
                kind: 0,
                order,
                paint: (ctx, origin) => {
                    ctx.save();
                    clipTo(ctx, clip, origin);
                    ctx.globalAlpha = alpha;
                    if (fill && el !== root) {
                        ctx.fillStyle = fill;
                        ctx.fillRect(box.left - origin.left, box.top - origin.top, box.right - box.left, box.bottom - box.top);
                    }
                    for (const { stroke, color } of strokes) {
                        if (stroke.kind === "fill" && stroke.rect) {
                            ctx.fillStyle = color;
                            ctx.fillRect(stroke.rect.left - origin.left, stroke.rect.top - origin.top, stroke.rect.right - stroke.rect.left, stroke.rect.bottom - stroke.rect.top);
                        } else if (stroke.from && stroke.to) {
                            ctx.strokeStyle = color;
                            ctx.lineWidth = stroke.width ?? 1;
                            ctx.setLineDash(stroke.dash ?? []);
                            ctx.beginPath();
                            ctx.moveTo(stroke.from[0] - origin.left, stroke.from[1] - origin.top);
                            ctx.lineTo(stroke.to[0] - origin.left, stroke.to[1] - origin.top);
                            ctx.stroke();
                        }
                    }
                    ctx.restore();
                },
            });
        }
        // 字：この要素の直下の字の節（子の要素の字は子が描く）
        for (const node of Array.from(el.childNodes)) {
            if (!(node instanceof Text) || !node.data.trim().replace(/ /g, "")) continue;
            const lines = linesOf(node).flatMap((line) => {
                const visible = intersect(line.rect, clip);
                return visible ? [{ ...line, visible }] : [];
            });
            if (lines.length === 0) continue;
            lines.forEach((line) => grow(line.visible));
            const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
            const color = style.color;
            const alpha = opacityOf(el, root);
            items.push({
                layer,
                kind: 1,
                order,
                paint: (ctx, origin) => {
                    ctx.save();
                    clipTo(ctx, clip, origin);
                    ctx.globalAlpha = alpha;
                    ctx.font = font;
                    ctx.fillStyle = color;
                    ctx.textBaseline = "alphabetic";
                    for (const line of lines) drawLine(ctx, line.text, line.rect, origin);
                    ctx.restore();
                },
            });
        }
    }
    const area = bounds as Rect | null;
    if (!area) return null;
    const origin: Rect = { left: Math.floor(area.left), top: Math.floor(area.top), right: Math.ceil(area.right), bottom: Math.ceil(area.bottom) };
    const width = origin.right - origin.left;
    const height = origin.bottom - origin.top;
    if (width <= 0 || height <= 0) return null;
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(width * IMAGE_SCALE);
    canvas.height = Math.round(height * IMAGE_SCALE);
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.scale(IMAGE_SCALE, IMAGE_SCALE);
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, width, height);
    // 重なりの順（同じ重なりの中では地と枠線を字より先に、そのあとは文書の順）
    items.sort((a, b) => compareLayers(a.layer, b.layer) || a.kind - b.kind || a.order - b.order);
    for (const item of items) await item.paint(ctx, origin);
    return { url: dataUrlWithResolution(canvas.toDataURL("image/png"), 96 * IMAGE_SCALE), width, height };
}
