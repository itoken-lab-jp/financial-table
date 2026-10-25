/**
 * 小計を内訳の端に残す（「行」カードの「小計を端に残す」。2026-10-25 ユーザー「販売費小計行を下にスティッキー出来たらいいんだけどね。小計行囲みのスティッキーね」）。
 *
 * 小計を下に置く区分・中分類は、内訳が見えているあいだ、その小計の行を表の見えている範囲の下の端に残す。内訳を最後まで送ると小計は自分の位置に収まり、
 * 区分を上へ抜けると残さない。入れ子（販管費の中の人件費の内訳など）は、内側の 1 行だけを残す（モックで比べて 2026-10-25 ユーザー決定）。
 * 小計を上に置く（上囲みの）区分は、同じ決まりを上下逆にして、列の見出しのすぐ下に残す（2026-10-25 ユーザー「小計を上囲みの時は上に sticky しないと」）。
 * 消した「合計行をスクロールしても残す」（2.1.12.0 の前に外した）とは別で、表全体の合計は残さない。
 *
 * 描き方：残すのは 1 つの表につき下に 1 行・上に 1 行までで、その行（左右の表の同じ行）に data-pin-on（bottom・top）を付けて CSS で sticky にする。止まる位置は
 * ブラウザが決めるので、スクロールのたびに描き直さない。スクロールで決めるのはどの行に印を付けるかだけで、行の位置は描いたあとに 1 度測って
 * 覚え、スクロールではスクロールの位置（scrollTop・clientHeight）しか読まない。行をまたぐセル（セグメントの名前）は切らない（残す行は
 * その下から動かず、名前のセルは後ろにそのまま見える）
 */
import { TOTAL_KEY } from "./rows";
import type { TableRow } from "./viewModel";

/**
 * 行ごとの、残せる小計の内訳の行の数。小計を下に置いた行は、すぐ前に続く自分より深い行の数（正）、上に置いた行は、すぐ後ろに続く
 * 自分より深い行の数に - を付けたもの（負）。残せない行は 0。
 * 残せるのは開いていて内訳が 1 行以上ある小計の行。表全体の合計と、行をまたぐセルが始まる行
 * （セグメントの名前・横積みの親の名前。残すとセルが一緒に動く）は残さない
 */
export function pinSpans(rows: readonly TableRow[]): number[] {
    const inside = (row: TableRow, other: TableRow | undefined) => !!other && other.depth > row.depth && (other.orgPath ?? null) === (row.orgPath ?? null);
    return rows.map((row, index) => {
        if (row.type !== "subtotal" || row.open === false || row.code === TOTAL_KEY) return 0;
        if (row.orgCells?.length || row.hierarchyCells?.length) return 0;
        if (row.aggregatePosition === "bottom") {
            let from = index;
            while (inside(row, rows[from - 1])) from--;
            return index - from;
        }
        if (row.aggregatePosition === "top") {
            let to = index;
            while (inside(row, rows[to + 1])) to++;
            return index - to;
        }
        return 0;
    });
}

/**
 * 残せる小計の、スクロールの中の位置（スクロールする箱の中身の上からの CSS px）。上に残す小計は、上下を逆にした座標
 * （値に - を付け、start は内訳の最後の行の下の端）で持ち、下に残す小計と同じ決まりで選ぶ
 */
export interface PinPlace {
    /** 内訳の 1 行目の上の端 */
    start: number;
    /** 小計の行の自分の位置の上と下の端 */
    top: number;
    bottom: number;
    /** 内訳の中の、残せる小計（入れ子）の下の端（上から順） */
    inner: number[];
}

/**
 * 残す小計（places の番号、無ければ -1）。edge は見えている範囲の下の端。
 * 下の端にかかった行を囲む内側の小計（下の端が内訳か小計の行にかかっていて、小計は全部は見えていない）を残す。入れ子の外の小計は、内側の小計より
 * 下の内訳にかかったときだけ残す（内側の小計が見えてから、外の小計を重ねて隠さない）。ただし残す行は、内訳の 1 行目（入れ子のあとなら、
 * 内側の小計のすぐ下の行）から下の端までに収まるときだけ（区分の外や内側の小計に重ねない。収まらないあいだは何も残さない）。
 * 囲む小計が無いときは、自分の位置が下の端のすぐ上（見えている高さの分）にある小計に印を付けておく。止まらないので見た目は変わらず、
 * 上へ送り返したとき、印を付ける前の 1 コマに小計が下の端から沈んで見えないように
 */
export function activePin(places: readonly PinPlace[], edge: number, height: number): number {
    let best = -1;
    let ahead = -1;
    places.forEach((p, i) => {
        if (p.start < edge && p.bottom > edge + 0.5) {
            // 内側：内訳の始まりが下のもの。始まりが同じ（人件費が販管費の最初の内訳）なら、小計が上のもの
            if (best < 0 || p.start > places[best].start || (p.start === places[best].start && p.bottom < places[best].bottom)) best = i;
        } else if (p.bottom <= edge + 0.5 && p.bottom > edge - height && (ahead < 0 || p.bottom > places[ahead].bottom)) {
            ahead = i;
        }
    });
    if (best < 0) return ahead;
    const p = places[best];
    // 内側の小計の中の行にかかっているなら、その内側の小計が best になっている。ここに来るのは、内側の小計より下の内訳
    const from = p.inner.reduce((top, bottom) => (bottom <= edge + 0.5 ? Math.max(top, bottom) : top), p.start);
    return from + (p.bottom - p.top) <= edge + 0.5 ? best : -1;
}

interface Side {
    rows: HTMLTableRowElement[];
    places: PinPlace[];
}

interface Measured {
    bottom: Side;
    top: Side;
    /** 列の見出しの高さ（上に残す行は、その下に止まる） */
    head: number;
}

/** スクロールする箱ごとの、表ごとの残せる小計（描くたびに測り直す） */
const measured = new WeakMap<HTMLElement, Measured[]>();

/** 描いたあと（行の高さをそろえたあと）に、残せる小計の位置を測って覚え、印を付け直す。scale は拡大の倍率（画面の px ÷ CSS px） */
export function measurePins(scroller: HTMLElement | null, scale: number): void {
    if (!scroller) return;
    // 残している行は止まった所に描かれ、その下の行の自分の位置も測れなくなるので、印を外してから測る（同じ処理の中で付け直すので、外した絵は出ない）
    for (const tr of Array.from(scroller.querySelectorAll<HTMLTableRowElement>("tr[data-pin-on]"))) {
        tr.removeAttribute("data-pin-on");
        tr.removeAttribute("data-pin-stuck");
    }
    const box = scroller.getBoundingClientRect();
    const toContent = (y: number) => (y - box.top) / scale - scroller.clientTop + scroller.scrollTop;
    const tables: Measured[] = [];
    for (const table of Array.from(scroller.querySelectorAll<HTMLTableElement>("table"))) {
        const bottom: Side = { rows: [], places: [] };
        const top: Side = { rows: [], places: [] };
        const body = Array.from(table.tBodies).flatMap((b) => Array.from(b.rows));
        body.forEach((tr, index) => {
            const after = Number(tr.dataset.pinTo ?? 0);
            const before = Number(tr.dataset.pinFrom ?? 0);
            const rect = tr.getBoundingClientRect();
            if (before > 0 && index - before >= 0) {
                bottom.rows.push(tr);
                bottom.places.push({ start: toContent(body[index - before].getBoundingClientRect().top), top: toContent(rect.top), bottom: toContent(rect.bottom), inner: [] });
            } else if (after > 0 && index + after < body.length) {
                top.rows.push(tr);
                top.places.push({ start: -toContent(body[index + after].getBoundingClientRect().bottom), top: -toContent(rect.bottom), bottom: -toContent(rect.top), inner: [] });
            }
        });
        // 入れ子：内訳の範囲が自分の内訳の中にある、ほかの残せる小計（下に残す小計は前、上に残す小計は後ろにある）
        for (const side of [bottom, top]) {
            side.places.forEach((p, i) => {
                p.inner = side.places.filter((q, j) => j !== i && q.start >= p.start && q.bottom < p.bottom).map((q) => q.bottom);
            });
        }
        if (bottom.rows.length || top.rows.length) tables.push({ bottom, top, head: (table.tHead?.getBoundingClientRect().height ?? 0) / scale });
    }
    measured.set(scroller, tables);
    placePins(scroller);
}

/** スクロールのたびに、表ごとに残す小計に印を付ける。読むのはスクロールの位置だけで、印が変わるときだけ書く */
export function placePins(scroller: HTMLElement | null): void {
    const tables = scroller ? measured.get(scroller) : undefined;
    if (!scroller || !tables?.length) return;
    const height = scroller.clientHeight;
    const scrollTop = scroller.scrollTop;
    const mark = (side: Side, edge: number, name: string) => {
        const on = activePin(side.places, edge, height);
        side.rows.forEach((tr, i) => {
            const value = i === on ? name : null;
            if (tr.getAttribute("data-pin-on") !== value) {
                if (value) tr.setAttribute("data-pin-on", value);
                else tr.removeAttribute("data-pin-on");
            }
            // 止まっている（自分の位置から動いた）あいだだけ、隣の行を地で覆う。止まっていない行が覆うと、隣の行の縦の線が途切れた
            const stuck = i === on && side.places[i].bottom > edge + 0.5;
            if (stuck !== tr.hasAttribute("data-pin-stuck")) tr.toggleAttribute("data-pin-stuck", stuck);
        });
    };
    for (const { bottom, top, head } of tables) {
        mark(bottom, scrollTop + height, "bottom");
        mark(top, -(scrollTop + head), "top");
    }
}
