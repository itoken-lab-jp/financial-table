/**
 * キーボードで表の中を動く（Power BI の標準のテーブル・マトリックスと同じキー）。
 *
 * - 表は Tab で 1 か所だけ止まり（最後にいたセル）、中は矢印で 1 つずつ動く（見出し・組織の名前・行の見出し・値のセルをまたいで）
 * - Home・End（Ctrl+←・→）で行の端、Ctrl+↑・↓ で列の端、Ctrl+Home・End で表の角、PageUp・PageDown で見えている一番上・一番下の行
 *   （もうそこにいれば、見えている行の数だけ先へ）
 * - Enter・Space で選び、Ctrl（Mac は ⌘）を足すと足し引き。Shift+F10・メニュー キーで Power BI のメニュー、Ctrl+Shift+C で選択を解く
 * - Shift+→・← で行・組織を開き閉じする
 *
 * 行をまたぐ組織の名前・列をまたぐ期間の見出しがあるので、表を升目に広げてから動く。ここは升目の計算だけで、DOM は App.tsx が読む
 */

/** 升目の位置（行・列、どちらも 0 から） */
export interface GridPosition {
    row: number;
    col: number;
}

/** 表の 1 つのセルと、それがまたぐ行・列の数 */
export interface SpannedCell<T> {
    item: T;
    rowSpan: number;
    colSpan: number;
}

/**
 * 行ごとのセル（HTML の表と同じ並び）を升目に広げる。行をまたぐセルは下の行の同じ列を埋め、あとの行のセルはその右に置く。
 * rowSpan・colSpan が 0 以下なら 1 とみなす（HTML の rowspan=0 は使っていない）
 */
export function buildGrid<T>(rows: Array<Array<SpannedCell<T>>>): Array<Array<T | undefined>> {
    const grid: Array<Array<T | undefined>> = rows.map((): Array<T | undefined> => []);
    rows.forEach((cells, r) => {
        let c = 0;
        for (const cell of cells) {
            while (grid[r][c] !== undefined) c++;
            const rowSpan = Math.max(1, cell.rowSpan);
            const colSpan = Math.max(1, cell.colSpan);
            for (let dr = 0; dr < rowSpan && r + dr < grid.length; dr++) {
                for (let dc = 0; dc < colSpan; dc++) grid[r + dr][c + dc] = cell.item;
            }
            c += colSpan;
        }
    });
    return grid;
}

/** 動くキー（修飾キーの組み合わせまで見たもの）。動かないキーは null */
export type Move = "up" | "down" | "left" | "right" | "rowStart" | "rowEnd" | "colStart" | "colEnd" | "first" | "last" | "pageUp" | "pageDown";

export function moveOf(key: string, ctrl: boolean, shift: boolean, alt: boolean): Move | null {
    // Shift+矢印は開き閉じ（行と組織の名前のセル）に使う。Alt の付いたものは Power BI とブラウザーに任せる
    if (shift || alt) return null;
    switch (key) {
        case "ArrowUp":
            return ctrl ? "colStart" : "up";
        case "ArrowDown":
            return ctrl ? "colEnd" : "down";
        case "ArrowLeft":
            return ctrl ? "rowStart" : "left";
        case "ArrowRight":
            return ctrl ? "rowEnd" : "right";
        case "Home":
            return ctrl ? "first" : "rowStart";
        case "End":
            return ctrl ? "last" : "rowEnd";
        case "PageUp":
            return "pageUp";
        case "PageDown":
            return "pageDown";
        default:
            return null;
    }
}

/** 升目の中で item の占める範囲（from を含むなら from を起点にする） */
function extentOf<T>(grid: Array<Array<T | undefined>>, item: T, row: number, col: number): { top: number; bottom: number; left: number; right: number } {
    let top = row;
    while (top > 0 && grid[top - 1][col] === item) top--;
    let bottom = row;
    while (bottom + 1 < grid.length && grid[bottom + 1][col] === item) bottom++;
    let left = col;
    while (left > 0 && grid[row][left - 1] === item) left--;
    let right = col;
    while (right + 1 < grid[row].length && grid[row][right + 1] === item) right++;
    return { top, bottom, left, right };
}

/** item が升目のどこにあるか。起点（前に動いた先）がまだ item の上なら起点を使う（行をまたぐセルを通り抜けても列を保つ） */
export function positionOf<T>(grid: Array<Array<T | undefined>>, item: T, anchor: GridPosition | null): GridPosition | null {
    if (anchor && grid[anchor.row]?.[anchor.col] === item) return anchor;
    for (let row = 0; row < grid.length; row++) {
        const col = grid[row].indexOf(item);
        if (col >= 0) return { row, col };
    }
    return null;
}

/** 見えている体の行（升目の行の番号。上と下の端がどちらも見えている行の、最初と最後） */
export interface VisibleRows {
    first: number;
    last: number;
}

/**
 * 動いた先の位置。行・列をまたぐセルからは、またいだ範囲の外へ出る。端より先へは動かない（同じセルのままなら null）。
 * visible は PageUp・PageDown で使う、見えている行
 */
export function moveInGrid<T>(grid: Array<Array<T | undefined>>, from: GridPosition, move: Move, visible: VisibleRows): GridPosition | null {
    const item = grid[from.row]?.[from.col];
    if (item === undefined) return null;
    const lastRow = grid.length - 1;
    const width = (row: number) => grid[row].length;
    const span = extentOf(grid, item, from.row, from.col);
    let next: GridPosition;
    switch (move) {
        case "up":
            next = { row: span.top - 1, col: from.col };
            break;
        case "down":
            next = { row: span.bottom + 1, col: from.col };
            break;
        case "left":
            next = { row: from.row, col: span.left - 1 };
            break;
        case "right":
            next = { row: from.row, col: span.right + 1 };
            break;
        case "rowStart":
            next = { row: from.row, col: 0 };
            break;
        case "rowEnd":
            next = { row: from.row, col: width(from.row) - 1 };
            break;
        case "colStart":
            next = { row: 0, col: from.col };
            break;
        case "colEnd":
            next = { row: lastRow, col: from.col };
            break;
        case "first":
            next = { row: 0, col: 0 };
            break;
        case "last":
            next = { row: lastRow, col: width(lastRow) - 1 };
            break;
        case "pageUp": {
            const page = Math.max(1, visible.last - visible.first);
            next = { row: span.top > visible.first ? visible.first : Math.max(0, span.top - page), col: from.col };
            break;
        }
        case "pageDown": {
            const page = Math.max(1, visible.last - visible.first);
            next = { row: span.bottom < visible.last ? visible.last : Math.min(lastRow, span.bottom + page), col: from.col };
            break;
        }
    }
    if (next.row < 0 || next.row > lastRow || next.col < 0 || next.col >= width(next.row)) return null;
    const target = grid[next.row][next.col];
    if (target === undefined || target === item) return null;
    return next;
}

/** Power BI のメニューを出すキー（Shift+F10・メニュー キー）。Alt・Ctrl の付いたもの（Alt+Shift+F10 はビジュアルのメニュー）は Power BI に任せる */
export function isMenuKey(key: string, shift: boolean, ctrl: boolean, alt: boolean): boolean {
    if (ctrl || alt) return false;
    return key === "ContextMenu" || (key === "F10" && shift);
}

/** 選ぶキー（Enter・Space） */
export function isSelectKey(key: string): boolean {
    return key === "Enter" || key === " " || key === "Spacebar";
}

/** 選択を解くキー（Ctrl+Shift+C） */
export function isClearKey(key: string, ctrl: boolean, shift: boolean): boolean {
    return ctrl && shift && key.toLowerCase() === "c";
}

/** 開き閉じのキー（Shift+→ で開く、Shift+← で閉じる）。どちらでもなければ null */
export function foldOf(key: string, shift: boolean, ctrl: boolean, alt: boolean): "open" | "close" | null {
    if (!shift || ctrl || alt) return null;
    return key === "ArrowRight" ? "open" : key === "ArrowLeft" ? "close" : null;
}
