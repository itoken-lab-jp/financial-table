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
    th.style.borderRightWidth = "";
    th.style.removeProperty("--ft-fl-br");
    th.style.removeProperty("--ft-fl-bw");
    th.style.removeProperty("--ft-fl-bwb");
    for (const e of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-fl-cover-bottom"))) e.remove();
    for (const e of Array.from(th.querySelectorAll<HTMLElement>(":scope > .ft-fl-hide, :scope > .ft-fl-second, :scope > .ft-fl-cover, :scope > .ft-fl-edge, :scope > .ft-fl-inset, :scope > .ft-fl-rise, :scope > .ft-fl-sink"))) {
        e.classList.remove("ft-fl-hide", "ft-fl-second", "ft-fl-cover", "ft-fl-edge", "ft-fl-inset", "ft-fl-rise", "ft-fl-sink");
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
    for (const e of lines) e.classList.add(double ? "ft-fl-second" : "ft-fl-hide");
    // 線を引かない所：帯を上へ重ねて塗る span（top が -1px で幅を持つもの）
    for (const e of Array.from(th.querySelectorAll<HTMLElement>(`:scope > .${box}`))) {
        if (e.style.top === "-1px" && e.style.width) e.classList.add("ft-fl-cover");
        else if (e.style.top === "0px") e.classList.add("ft-fl-rise");
    }
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
}
