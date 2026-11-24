/**
 * 左に残す表（.ft-freeze-left）の横線を、div でなく th の罫線で引く（2026-10-24）。
 *
 * 左右の表の行の位置は DOM ではそろっていても、左の線を div（.ft-name-line・.ft-box-line・.ft-seg-box-line・.ft-box-bottom）、右の線を td の
 * 罫線で引くと、Desktop の 125% で左だけ 1 画素（実寸大）〜2 画素（ページに合わせる）下に描かれた。div の位置と罫線では画素への丸め方が違う。
 * 同じ行の td と同じ太さの th の罫線にすると、ページに合わせる・実寸大のどちらでもそろった（診断版 2.1.9.906 を Desktop で見た）。
 *
 * 線の種類と色を決める CSS（行のクラスで決まる。visual.less）はそのまま div に効かせ、ここで div の計算スタイルを読んで th の罫線に写し、
 * div は描かない（ft-fl-hide）。部品の置き直しのクラスの意味は visual.less の「左に残す表の線を th の罫線で引く」。
 */
import { minOf } from "./shared/minMax";

/** 字下げ 1 段の幅（App.tsx の INDENT_EM と同じ） */
const INDENT_EM = 1;

const shownLine = (e: HTMLElement, side: "Top" | "Bottom") => {
    const cs = getComputedStyle(e);
    return cs.display !== "none" && (parseFloat(cs[`border${side}Width`]) || 0) > 0 && cs[`border${side}Style`] !== "none";
};

/** 置いた値を外して、React と CSS が決めた形に戻す */
function resetCell(th: HTMLElement): void {
    th.style.borderTop = "";
    th.style.borderBottom = "";
    th.style.paddingTop = "";
    th.style.paddingBottom = "";
    th.style.background = "";
    for (const k of ["backgroundImage", "backgroundRepeat", "backgroundOrigin", "backgroundClip", "backgroundPosition", "backgroundSize"] as const) th.style[k] = "";
    th.style.borderRightWidth = "";
    th.style.removeProperty("--ft-fl-br");
    th.style.removeProperty("--ft-fl-bw");
    th.style.removeProperty("--ft-fl-bwb");
    for (const e of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-fl-cover-bottom, :scope > .ft-fl-table-bottom"))) e.remove();
    for (const e of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-fl-hide, :scope > .ft-fl-second, :scope > .ft-fl-lift, :scope > .ft-fl-cover, :scope > .ft-fl-edge, :scope > .ft-fl-inset, :scope > .ft-fl-rise, :scope > .ft-fl-sink"))) {
        e.classList.remove("ft-fl-hide", "ft-fl-second", "ft-fl-lift", "ft-fl-cover", "ft-fl-edge", "ft-fl-inset", "ft-fl-rise", "ft-fl-sink");
    }
}

/**
 * th の上の線（.ft-name-line、箱の中なら .ft-box-line／.ft-seg-box-line の並び）を th の border-top にする。線の種類と色は今までの div、太さは
 * 同じ行の右の td の上の罫線（0 なら引かない）。線を引かない所（上の行から続く帯）は、帯を 1px 上へ重ねて塗る span（NestedBox）を罫線の上に重ねて消す
 */
function topLine(th: HTMLElement, td: HTMLElement, box: string): void {
    const lines = Array.from(th.querySelectorAll<HTMLElement>(`:scope > .ft-name-line, :scope > .${box}-line`)).filter((e) => shownLine(e, "Top"));
    if (lines.length === 0) return;
    const width = parseFloat(getComputedStyle(td).borderTopWidth) || 0;
    if (width <= 0) return;
    // 線の種類は行で決まる（どの span も同じ）。右の端まで引く span（自分の段から右）を見本にする
    const sample = lines.find((e) => e.style.right === "0px") ?? lines[lines.length - 1];
    const ls = getComputedStyle(sample);
    // 締めの行の二重線は、div の上と下の枠線の 2 本（visual.less の .ft-row-total .ft-name-line）。右の td は上の線を枠線、下の線を余白に重ねた
    // 部品で引くので、上の線だけ th の罫線にして、div は下の線だけ残す（ft-fl-second）。どちらも罫線の下の余白から 2px の所で、丸め方が同じ
    const double = (parseFloat(ls.borderBottomWidth) || 0) > 0;
    const was = parseFloat(ls.borderTopWidth) || 0;
    th.style.borderTop = `${width}px ${ls.borderTopStyle} ${ls.borderTopColor}`;
    th.style.setProperty("--ft-fl-bw", `${width}px`);
    // 今までの th は、罫線の代わりに上の余白を線の太さだけ足していた
    th.style.paddingTop = `${Math.max(0, parseFloat(getComputedStyle(th).paddingTop) - was)}px`;
    // 線を引かない所：帯を上へ重ねて塗る span（top が -1px で幅を持つもの）
    for (const e of Array.from(th.querySelectorAll<HTMLElement>(`:scope > .${box}`))) {
        if (e.style.top === "-1px" && e.style.width) e.classList.add("ft-fl-cover");
        else if (e.style.top === "0px") e.classList.add("ft-fl-rise");
    }
    // 行の名前の列の帯（.ft-box）は、地を罫線の 1px 下から塗り（visual.less）、その 1px は th の地を帯の色にして埋める。帯の地が罫線のすぐ下から
    // 始まると、端数の倍率で画素にそろえた帯の地が、2 画素にまたがった罫線の下の画素を塗り、行の名前の列だけ線が細く見えた（2026-11-12 ユーザー）
    // セグメントの列の帯（.ft-seg-box）も同じ（2026-11-12。ハーネスで全体を 0.85 倍にしたとき、合計の行の上の線の下の画素が消えた）
    underLine(th, `:scope > .${box}.ft-fl-rise`);
    // 線の始まり（帯の右）は、div を隠す前に測る
    const cell = th.getBoundingClientRect();
    const scale = cell.width / (th.offsetWidth || cell.width || 1);
    const from = minOf(lines.map((e) => (e.getBoundingClientRect().left - cell.left) / scale));
    for (const e of lines) e.classList.add(double ? "ft-fl-second" : "ft-fl-hide");
    // th の罫線を透明にして（太さは残す）、線のある所（帯が上の行から続くセルは帯の右から）だけ th の背景の絵（罫線の所、border box の上の端）で引く。
    // - 罫線の上に帯の色を重ねて消すと、端数の倍率で罫線が 2 画素にまたがって描かれるのに、重ねた span はセルの外の端（overflow: clip）で画素の途中から
    //   切られ、帯の中に薄い横線が透けた（2026-11-12 ユーザー、Desktop の拡大率が 100% 以外）
    // - 線を div で引くと Desktop では右の td の罫線より 1〜2 画素下に描かれ、横線が列の境でずれ、角が切れた（2026-10-24 と 2026-11-12「角の線の切れとズレ」）
    // - th の罫線のままだと、Desktop で行の名前の列だけ線の上の画素が薄かった（値の列 135 に対して 189。2026-11-12）。背景の絵は値の列とそろった
    if (double) return;
    paintLine(th, "top", Math.max(0, from), width, ls.borderTopColor);
}

/** th の背景の絵で横線を引く（罫線は透明にする）。上か下の端、from（border box の左から）から右の端まで */
function paintLine(th: HTMLElement, edge: "top" | "bottom", from: number, width: number, color: string): void {
    if (edge === "top") th.style.borderTopColor = "transparent";
    else th.style.borderBottomColor = "transparent";
    const add = (key: "backgroundImage" | "backgroundRepeat" | "backgroundOrigin" | "backgroundClip" | "backgroundPosition" | "backgroundSize", value: string) => {
        th.style[key] = th.style[key] ? `${th.style[key]}, ${value}` : value;
    };
    add("backgroundImage", `linear-gradient(${color}, ${color})`);
    add("backgroundRepeat", "no-repeat");
    add("backgroundOrigin", "border-box");
    add("backgroundClip", "border-box");
    add("backgroundPosition", `${from}px ${edge === "top" ? "0" : "100%"}`);
    add("backgroundSize", `calc(100% - ${from}px) ${width}px`);
}

/** 罫線の脇で地を 1px 空けた帯（selector）の色を th の地にする。帯は右の端まで塗るものの、一番内側（左が一番右）の色 */
function underLine(th: HTMLElement, selector: string): void {
    const bands = Array.from(th.querySelectorAll<HTMLElement>(selector)).filter((e) => e.style.right === "0px" || e.classList.contains("ft-fl-inset"));
    let inner: HTMLElement | null = null;
    for (const e of bands) if (!inner || e.getBoundingClientRect().left > inner.getBoundingClientRect().left) inner = e;
    if (!inner) return;
    const color = getComputedStyle(inner).backgroundColor;
    if (color && color !== "rgba(0, 0, 0, 0)") th.style.backgroundColor = color;
}

/**
 * 箱の締めの線（.ft-box-bottom）を th の border-bottom にする。段の内側から引く線は、外の段の帯（下の行へ続く）の所に帯の色を重ねて消す
 */
function bottomLine(th: HTMLElement, td: HTMLElement): void {
    const k = th.querySelector<HTMLElement>(":scope > .ft-box-bottom");
    if (!k || !shownLine(k, "Bottom")) return;
    const width = parseFloat(getComputedStyle(td).borderBottomWidth) || 0;
    if (width <= 0) return;
    const ks = getComputedStyle(k);
    th.style.borderBottom = `${width}px ${ks.borderBottomStyle} ${ks.borderBottomColor}`;
    th.style.setProperty("--ft-fl-bwb", `${width}px`);
    // 罫線の分だけ下の余白を減らし、字の位置を今までと同じにする（減らさないと、縦の真ん中に置いた字が罫線の半分だけ上がる）
    th.style.paddingBottom = `${Math.max(0, parseFloat(getComputedStyle(th).paddingBottom) - width)}px`;
    k.classList.add("ft-fl-hide");
    for (const strip of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-box"))) {
        if (strip.style.top === "0px") strip.classList.add("ft-fl-sink");
    }
    const from = parseFloat(k.style.left) || 0;
    if (from <= 0) return;
    // 帯の本体（right が 0 の span）のうち、締めの線より外の段
    for (const strip of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-box"))) {
        if (strip.style.right !== "0px" || strip.classList.contains("ft-fl-cover")) continue;
        const at = parseFloat(strip.style.left) || 0;
        if (at >= from - 0.001) continue;
        const cover = document.createElement("span");
        cover.className = "ft-box ft-fl-cover-bottom";
        cover.style.left = strip.style.left;
        cover.style.width = `${INDENT_EM}em`;
        if (strip.style.background) cover.style.background = strip.style.background;
        if (strip.style.borderLeftStyle) cover.style.borderLeftStyle = strip.style.borderLeftStyle;
        th.appendChild(cover);
    }
}

/**
 * セグメントの列どうしの境の縦線を、左のセルの右の枠線で引く（右のセルは描く範囲を切るので、セルの外に引く ft-edge-line が見えない）。
 * 左のセルは、右の端が右のセルの左の端と同じで、上下が重なるもの
 */
function orgEdges(table: HTMLTableElement): void {
    const orgs = Array.from(table.querySelectorAll<HTMLElement>(":scope > tbody > tr > th.ft-org"));
    for (const th of orgs) th.style.borderRightColor = "";
    const rects = orgs.map((th) => th.getBoundingClientRect());
    orgs.forEach((th, i) => {
        const edge = th.querySelector<HTMLElement>(":scope > .ft-edge-line");
        if (!edge || getComputedStyle(edge).display === "none") return;
        const color = getComputedStyle(edge).borderLeftColor;
        const r = rects[i];
        orgs.forEach((left, j) => {
            const l = rects[j];
            if (j === i || Math.abs(l.right - r.left) > 0.5 || l.bottom <= r.top + 0.5 || l.top >= r.bottom - 0.5) return;
            // 上の線を罫線にしたセルは、右の透明な枠線を外している（applyFrozenLines）。境を引くセルは枠線を戻す（戻さないと、上の線のある
            // 2 つ目からのブロックで境が消えた。2026-10-26 ユーザー）。色を付けるので、上の線との角で欠けることもない
            const removed = left.style.getPropertyValue("--ft-fl-br");
            if (removed) {
                left.style.borderRightWidth = "";
                left.style.removeProperty("--ft-fl-br");
                for (const e of Array.from(left.querySelectorAll<HTMLElement>(":scope > .ft-fl-inset"))) e.classList.remove("ft-fl-inset");
            }
            if ((parseFloat(getComputedStyle(left).borderRightWidth) || 0) > 0) left.style.borderRightColor = color;
        });
    });
}

/**
 * 左の表の線を、div でなく th の罫線で引く（右の td と同じ要素の罫線にして、位置と画素の丸めを同じにする）。
 *   行の名前のセル（th.ft-name）：上の線（.ft-name-line・.ft-box-line）と、箱の締めの線（.ft-box-bottom。段の内側から引くものも）
 *   セグメントの列（th.ft-org）：ブロックの切れ目の線（.ft-seg-box-line）
 * 左の列との境の縦線（.ft-edge-line）は罫線の分だけ伸ばす。毎回、置いた値を外してから CSS の値を読み直す
 */
export function applyFrozenLines(pairs: ReadonlyArray<readonly [HTMLElement, HTMLElement]>): void {
    const cells: Array<[HTMLElement, HTMLElement, string]> = [];
    for (const [a, b] of pairs) {
        const td = b.querySelector<HTMLElement>(":scope > td");
        if (!td) continue;
        for (const th of Array.from(a.querySelectorAll<HTMLElement>(":scope > th.ft-name"))) cells.push([th, td, "ft-box"]);
        for (const th of Array.from(a.querySelectorAll<HTMLElement>(":scope > th.ft-org"))) cells.push([th, td, "ft-seg-box"]);
    }
    for (const [th] of cells) resetCell(th);
    for (const [th, td, box] of cells) {
        topLine(th, td, box);
        if (box === "ft-box") bottomLine(th, td);
        if (th.style.getPropertyValue("--ft-fl-bw") || th.style.getPropertyValue("--ft-fl-bwb")) {
            for (const e of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-edge-line"))) e.classList.add("ft-fl-edge");
            // 右の枠線が透明（縦線は表の上に重ねた部品が引く）なら、枠線を外す。残すと、上下の罫線の右の端が透明な枠線との角で斜めに切られ、
            // 数字の列との境の縦線の手前で 1 画素欠ける（div の線は右の枠線の手前で終わっていて、1 画素すき間があった）。列の幅は col が決めるので
            // セルの幅は変わらない（余白に足すと、列の幅を測る所が余白を数えて列が 1px 広がった）
            const cs = getComputedStyle(th);
            // 箱の塗りは今までどおり枠線の手前で止める（右の端まで引く span を枠線の幅だけ内へ）
            const right = parseFloat(cs.borderRightWidth) || 0;
            if (right > 0 && cs.borderRightColor === "rgba(0, 0, 0, 0)") {
                th.style.borderRightWidth = "0px";
                th.style.setProperty("--ft-fl-br", `${right}px`);
                for (const e of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-box, :scope > .ft-seg-box"))) {
                    if (e.style.right === "0px") e.classList.add("ft-fl-inset");
                }
            }
        }
    }
    const tables = new Set(pairs.map(([a]) => a.closest("table")).filter((t): t is HTMLTableElement => !!t));
    for (const table of tables) orgEdges(table);
    const all = new Set(pairs.flatMap(([a, b]) => [a.closest("table"), b.closest("table")]).filter((t): t is HTMLTableElement => !!t));
    for (const table of all) tableBottom(table);
}

/**
 * 表の一番下の線を、表の枠線でなく、一番下に届くセルの下の枠線で引く（左右の表とも）。表の枠線のままだと、左に残す表のセル（位置を持つ）の
 * 地や中の帯が、端数の倍率で 2 画素にまたがった枠線の上の画素を塗り、セグメントの列と行の名前の列だけ線が細く薄く見えた（2026-11-12 ユーザー
 * 「一番下の太さのバラつき」）。線を div で引くと Desktop で数字の表の線とずれるので、セルの枠線で引く
 */
function tableBottom(table: HTMLTableElement): void {
    for (const e of Array.from(table.querySelectorAll<HTMLElement>(".ft-fl-bottom"))) e.classList.remove("ft-fl-bottom");
    table.removeAttribute("data-fl-bottom");
    table.style.removeProperty("--ft-fl-bottom");
    table.style.removeProperty("--ft-fl-bottom-w");
    const cs = getComputedStyle(table);
    const width = parseFloat(cs.borderBottomWidth) || 0;
    if (width <= 0 || cs.borderBottomStyle === "none") return;
    const rows = Array.from(table.tBodies).flatMap((b) => Array.from(b.rows));
    const last = rows.length;
    const bottom: HTMLElement[] = [];
    rows.forEach((tr, i) => {
        for (const cell of Array.from(tr.cells)) if (i + (cell.rowSpan || 1) >= last) bottom.push(cell);
    });
    if (!bottom.length) return;
    table.style.setProperty("--ft-fl-bottom", `${width}px ${cs.borderBottomStyle} ${cs.borderBottomColor}`);
    table.style.setProperty("--ft-fl-bottom-w", `${width}px`);
    table.setAttribute("data-fl-bottom", "");
    for (const cell of bottom) {
        cell.classList.add("ft-fl-bottom");
        // セグメントの列の帯は、地を枠線の 1px 上で止め（visual.less）、その 1px は th の地を帯の色にして埋める。帯の地が枠線のすぐ上まで来ると、
        // 端数の倍率で画素にそろえた帯の地が、2 画素にまたがった枠線の上の画素を塗り、セグメントの列だけ線が薄く見えた（2026-11-12 ユーザー）
        if (cell.classList.contains("ft-org")) {
            underLine(cell, ":scope > .ft-seg-box");
            // セグメントの列のセル（行をまたぐ）の下の枠線は、Desktop で 3 画素ににじんで薄かった（2026-11-12）。上の線と同じく背景の絵で引く
            cell.style.borderBottom = `${width}px ${cs.borderBottomStyle} transparent`;
            paintLine(cell, "bottom", 0, width, cs.borderBottomColor);
        }
    }
}
