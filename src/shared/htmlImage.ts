/**
 * HTML で描いた表などを画像にする（copyImage の chartImage の HTML 版。CopyImageButton の capture に渡す）。
 *
 * 入れ物の中の見えている要素を、画面の位置のまま canvas に描く：地の色、枠線（実線・点線・破線・二重線）、ぼかしの無い内側の影（inset の
 * box-shadow で引いた線）、字。SVG は chartImage と同じく描く。
 * スクロールする箱の中は見えている範囲で切る。重なりの順は、位置を持つ要素（position が static でない）とその z-index で決める
 * （固定した見出し・列が、スクロールした中身の上に来る）。描かないもの：疑似要素（::before・::after）、外側の影とぼかしのある影、角の丸み、背景の画像、
 * 字の飾り（下線など）、フォームの部品の中身（select の選んだ字など）。写したくない所は SKIP_ATTRIBUTE を付ける。
 * 画像の大きさは、描いたものの見えている範囲を合わせた四角（入れ物の余白は切る）。
 * Power BI はビジュアルを拡大して描くことがあり、位置（getBoundingClientRect）は拡大したあと、字の大きさ・枠線の太さ（getComputedStyle）は
 * 拡大する前の値で来る。字と線は入れ物の倍率を掛けて描き、右・中央にそろえた字は四角の右端・中央を基準に描く（字の幅が画面と
 * canvas で少し違っても、そろえた端がずれない。canvas は tabular-nums を効かせない）
 */
import { ChartImage, IMAGE_SCALE, Rect, SKIP_ATTRIBUTE, canvasPng, clipTracker, drawSvg, intersect, withoutScrollbars } from "./copyImage";

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

/** 1 つの box-shadow（getComputedStyle の値を分けたもの。長さは px） */
export interface BoxShadow {
    color: string;
    x: number;
    y: number;
    blur: number;
    spread: number;
    inset: boolean;
}

/** getComputedStyle の box-shadow を影ごとに分ける（none は空）。色は rgb(…) などの関数か名前、長さは px の数 */
export function parseBoxShadows(value: string): BoxShadow[] {
    if (!value || value.trim() === "none") return [];
    const parts: string[] = [];
    let depth = 0;
    let current = "";
    for (const ch of value) {
        if (ch === "(") depth++;
        if (ch === ")") depth--;
        if (ch === "," && depth === 0) {
            parts.push(current);
            current = "";
        } else current += ch;
    }
    parts.push(current);
    return parts.flatMap((part) => {
        const tokens = part.trim().match(/[a-z-]+\([^)]*\)|\S+/gi) ?? [];
        const lengths: number[] = [];
        let color = "";
        let inset = false;
        for (const token of tokens) {
            if (token === "inset") inset = true;
            else if (/^-?[\d.]+(px)?$/.test(token)) lengths.push(Number.parseFloat(token));
            else color = token;
        }
        if (lengths.length < 2) return [];
        const [x, y, blur = 0, spread = 0] = lengths;
        return [{ color, x, y, blur, spread, inset }];
    });
}

/** 四角 a から b を除いた部分（4 つまでの四角） */
export function subtractRect(a: Rect, b: Rect): Rect[] {
    const hole = intersect(a, b);
    if (!hole) return [a];
    const rects: Rect[] = [
        { left: a.left, right: a.right, top: a.top, bottom: hole.top },
        { left: a.left, right: a.right, top: hole.bottom, bottom: a.bottom },
        { left: a.left, right: hole.left, top: hole.top, bottom: hole.bottom },
        { left: hole.right, right: a.right, top: hole.top, bottom: hole.bottom },
    ];
    return rects.filter((r) => r.right - r.left > 0 && r.bottom - r.top > 0);
}

/**
 * ぼかしの無い外側の影が塗る四角（ずらして広げた箱から、要素の箱 border box を除いた部分）。表は外側の影で線を引いたり、地の色で
 * 隙間を覆ったりしている（固定した見出しの上・左に残す列の左・残す合計行の下の線）。ぼかしのある影と内側の影は描かない（空）。
 * 長さは scale を掛けてから使う
 */
export function outerShadowRects(shadow: BoxShadow, box: Rect, scale = 1): Rect[] {
    if (shadow.inset || shadow.blur !== 0 || transparent(shadow.color)) return [];
    const spread = shadow.spread * scale;
    const outer: Rect = {
        left: box.left + shadow.x * scale - spread,
        top: box.top + shadow.y * scale - spread,
        right: box.right + shadow.x * scale + spread,
        bottom: box.bottom + shadow.y * scale + spread,
    };
    if (outer.right <= outer.left || outer.bottom <= outer.top) return [];
    return subtractRect(outer, box);
}

/**
 * ぼかしの無い内側の影が塗る四角（内側の縁の箱 padding box から、ずらして縮めた箱を除いた部分）。ぼかしのある影と外側の影は描かない（空）。
 * 長さは scale を掛けてから使う
 */
export function insetShadowRects(shadow: BoxShadow, box: Rect, scale = 1): Rect[] {
    if (!shadow.inset || shadow.blur !== 0 || transparent(shadow.color)) return [];
    const inner: Rect = {
        left: box.left + (shadow.x + shadow.spread) * scale,
        top: box.top + (shadow.y + shadow.spread) * scale,
        right: box.right + (shadow.x - shadow.spread) * scale,
        bottom: box.bottom + (shadow.y - shadow.spread) * scale,
    };
    const top = Math.min(Math.max(inner.top, box.top), box.bottom);
    const bottom = Math.max(Math.min(inner.bottom, box.bottom), top);
    const rects: Rect[] = [
        { left: box.left, right: box.right, top: box.top, bottom: top },
        { left: box.left, right: box.right, top: bottom, bottom: box.bottom },
        { left: box.left, right: Math.min(Math.max(inner.left, box.left), box.right), top, bottom },
        { left: Math.max(Math.min(inner.right, box.right), box.left), right: box.right, top, bottom },
    ];
    return rects.filter((r) => r.right - r.left > 0 && r.bottom - r.top > 0);
}

/** 線とみなす太さ（CSS px）。これより細い四角は、画面の画素で何画素の線かを見て太さを決める */
const LINE_WIDTH = 2;

/**
 * 線の太さを、画面の画素で何画素かにそろえて画像の画素に直す。画面は線を端末の画素にそろえて描くので、Desktop の 125% では
 * 1px の枠線も 1 画素の線（表の上・見出しの下の線、--ft-px）も同じ 1 画素に見える。画像（画面の IMAGE_SCALE 倍）でそのまま描くと、
 * 枠線は 2 画素・1 画素の線は 1 画素と太さが食い違ったので、画面の画素の数を画像の倍率に直す
 */
export function lineWidth(width: number, ratio: number, scale = IMAGE_SCALE): number {
    const screen = Math.max(1, Math.round(width * ratio));
    return Math.max(1, Math.round((screen * scale) / ratio)) / scale;
}

/**
 * 四角の端を画像の画素の境にそろえる（地・枠線・影）。画面は枠線を端末の画素にそろえて描くが、canvas は端数の位置で 2 画素ににじませ、
 * 線が薄く太く見えた。どの端も同じ丸め方なので、隣り合う四角は隙間なく並ぶ。細い四角（線）は、太さを lineWidth で決める
 * （ratio は画面の devicePixelRatio）
 */
export function snapRect(r: Rect, ratio = 1, scale = IMAGE_SCALE): Rect {
    const snap = (v: number) => Math.round(v * scale) / scale;
    const width = r.right - r.left;
    const height = r.bottom - r.top;
    let left = snap(r.left);
    let right = snap(r.right);
    let top = snap(r.top);
    let bottom = snap(r.bottom);
    if (width > 0 && width < LINE_WIDTH) right = left + lineWidth(width, ratio, scale);
    if (height > 0 && height < LINE_WIDTH) bottom = top + lineWidth(height, ratio, scale);
    return { left, top, right, bottom };
}

/** 四角を塗る（画素の境にそろえて） */
function fillRect(ctx: CanvasRenderingContext2D, rect: Rect, origin: Rect, color: string): void {
    const r = snapRect(rect, window.devicePixelRatio || 1);
    ctx.fillStyle = color;
    ctx.fillRect(r.left - origin.left, r.top - origin.top, r.right - r.left, r.bottom - r.top);
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

/** 祖先の opacity を掛けたもの（root まで。祖先ごとに覚える） */
function opacityOf(el: Element, root: HTMLElement, cache: Map<Element, number>): number {
    const cached = cache.get(el);
    if (cached !== undefined) return cached;
    const own = Number.parseFloat(getComputedStyle(el).opacity) || 0;
    const value = el === root || !el.parentElement ? own : own * opacityOf(el.parentElement, root, cache);
    cache.set(el, value);
    return value;
}

function clipTo(ctx: CanvasRenderingContext2D, clip: Rect, origin: Rect): void {
    ctx.beginPath();
    ctx.rect(clip.left - origin.left, clip.top - origin.top, clip.right - clip.left, clip.bottom - clip.top);
    ctx.clip();
}

type Align = "left" | "right" | "center";

/** 字のそろえ方（text-align から。start・justify などは左） */
function alignOf(textAlign: string): Align {
    if (textAlign === "right" || textAlign === "end" || textAlign === "-webkit-right") return "right";
    if (textAlign === "center" || textAlign === "-webkit-center") return "center";
    return "left";
}

/** 字の 1 行を、画面の四角（字の行の箱）の中に、同じフォントで描く。右・中央にそろえた字は、四角の右端・中央を基準にする */
function drawLine(ctx: CanvasRenderingContext2D, text: string, rect: Rect, origin: Rect, align: Align): void {
    const metrics = ctx.measureText(text);
    const ascent = metrics.fontBoundingBoxAscent || metrics.actualBoundingBoxAscent;
    const descent = metrics.fontBoundingBoxDescent || metrics.actualBoundingBoxDescent;
    const height = rect.bottom - rect.top;
    const x = align === "right" ? rect.right : align === "center" ? (rect.left + rect.right) / 2 : rect.left;
    ctx.textAlign = align;
    ctx.fillText(text, x - origin.left, rect.top + (height - (ascent + descent)) / 2 + ascent - origin.top);
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
    // Power BI の拡大の倍率（位置は拡大したあと、字の大きさと線の太さは拡大する前の値で来る）
    const rootBox = root.getBoundingClientRect();
    const scale = root.offsetWidth > 0 && rootBox.width > 0 ? rootBox.width / root.offsetWidth : 1;
    // 写さない・SVG の中（SVG ごと描く）・見える範囲・opacity は、要素ごとに祖先をたどらず、祖先ごとに覚える（大きな表で秒かかった）
    const clips = clipTracker(root);
    const opacities = new Map<Element, number>();
    const passed = new Map<Element, boolean>();
    const skipped = (el: Element): boolean => {
        const cached = passed.get(el);
        if (cached !== undefined) return cached;
        const parent = el === root ? null : el.parentElement;
        const value = parent
            ? skipped(parent) || parent instanceof SVGSVGElement || el.matches(skipSelector)
            : el.closest(skipSelector) !== null || el.parentElement?.closest("svg") != null;
        passed.set(el, value);
        return value;
    };
    let order = 0;
    const elements = [root, ...Array.from(root.querySelectorAll("*"))];
    for (const el of elements) {
        order++;
        // SVG の中の要素は SVG ごと描く
        if (skipped(el)) continue;
        if (el instanceof SVGSVGElement) {
            const visible = clips.visible(el);
            if (!visible) continue;
            grow(visible);
            items.push({ layer: layerOf(el, root, layers), kind: 0, order, paint: (ctx, origin) => drawSvg(ctx, el, root, origin) });
            continue;
        }
        if (!(el instanceof HTMLElement)) continue;
        // 見えていない要素は計算スタイルを読まずに飛ばす
        const shown = clips.visible(el);
        if (!shown) continue;
        const style = getComputedStyle(el);
        if (style.display === "none" || style.visibility === "hidden") continue;
        // スクロールの箱は、スクロールバーを除いた四角を箱とする（画像にバーは描かないので、枠線を中身に付けて閉じる）
        const box = withoutScrollbars(el);
        const clip = intersect(shown, box);
        if (!clip) continue;
        const fill = transparent(style.backgroundColor) ? null : style.backgroundColor;
        const strokes = (["top", "right", "bottom", "left"] as Side[]).flatMap((side) => {
            const color = style.getPropertyValue(`border-${side}-color`);
            if (transparent(color)) return [];
            const width = (Number.parseFloat(style.getPropertyValue(`border-${side}-width`)) || 0) * scale;
            return borderStrokes(side, width, style.getPropertyValue(`border-${side}-style`), box).map((stroke) => ({ stroke, color }));
        });
        // 内側の影は、枠線の内側（padding box）に塗る
        const border = (side: Side) => (Number.parseFloat(style.getPropertyValue(`border-${side}-width`)) || 0) * scale;
        const padding: Rect = { left: box.left + border("left"), top: box.top + border("top"), right: box.right - border("right"), bottom: box.bottom - border("bottom") };
        const parsedShadows = parseBoxShadows(style.boxShadow);
        const shadows = parsedShadows.flatMap((shadow) => insetShadowRects(shadow, padding, scale).map((rect) => ({ rect, color: shadow.color })));
        // 外側の影は箱の外に出るので、箱では切らず、祖先のはみ出しを切る箱と root で切る。写す範囲には数えない
        const outerClip = el === root ? null : clips.overflow(el);
        const outerShadows = outerClip
            ? parsedShadows
                  // CSS は先に書いた影を上に描くので、後ろの影から塗る
                  .reverse()
                  .flatMap((shadow) => outerShadowRects(shadow, box, scale).flatMap((rect) => {
                      const visible = intersect(rect, outerClip);
                      return visible ? [{ rect: visible, color: shadow.color }] : [];
                  }))
            : [];
        const layer = layerOf(el, root, layers);
        if (fill || strokes.length > 0 || shadows.length > 0 || outerShadows.length > 0) {
            // 入れ物そのものの地は背景で塗るので、写す範囲には数えない（余白まで写さない）
            if (el !== root && (fill || strokes.length > 0 || shadows.length > 0)) grow(clip);
            const alpha = opacityOf(el, root, opacities);
            items.push({
                layer,
                kind: 0,
                order,
                paint: (ctx, origin) => {
                    ctx.save();
                    ctx.globalAlpha = alpha;
                    // 外側の影は地の下
                    for (const { rect, color } of outerShadows) fillRect(ctx, rect, origin, color);
                    // 箱で切る所も、塗る四角と同じくそろえる（そろえて太らせた線を、元の端数の箱で削らない）
                    clipTo(ctx, snapRect(clip, window.devicePixelRatio || 1), origin);
                    // 地は background-clip の所まで（padding-box なら枠線の内側。透明な枠線で上下へ伸ばして地を枠線の内側だけに塗り、
                    // すぐ上の行の罫線を残す帯がある。箱いっぱいに塗ると、画像では罫線が帯の色で消えた）
                    if (fill && el !== root) fillRect(ctx, style.backgroundClip === "padding-box" ? padding : box, origin, fill);
                    for (const { rect, color } of shadows) fillRect(ctx, rect, origin, color);
                    for (const { stroke, color } of strokes) {
                        if (stroke.kind === "fill" && stroke.rect) fillRect(ctx, stroke.rect, origin, color);
                        else if (stroke.from && stroke.to) {
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
            const font = `${style.fontStyle} ${style.fontWeight} ${(Number.parseFloat(style.fontSize) || 0) * scale}px ${style.fontFamily}`;
            const align = alignOf(style.textAlign);
            const color = style.color;
            const alpha = opacityOf(el, root, opacities);
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
                    for (const line of lines) drawLine(ctx, line.text, line.rect, origin, align);
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
    return { url: await canvasPng(canvas), width, height };
}
