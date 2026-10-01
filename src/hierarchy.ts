import type { DisplayRow, RowModel } from "./rows";
import type { TableRow } from "./viewModel";
import type { HierarchyOptions, HierarchyStyle, SplitDirection } from "./hierarchySettings";
import type { OrgCell } from "./orgs";

/** 計算式に手を加えず、表示に使う親子だけを組み立てる。 */
export interface AccountNode { row: DisplayRow; children: AccountNode[]; following: AccountNode[] }
export interface AccountTree { roots: AccountNode[]; parents: string[]; warnings: string[] }

export function accountTree(model: RowModel): AccountTree {
    const nodes = new Map(model.display.map(row => [row.def.code, { row, children: [], following: [] } as AccountNode]));
    const owned = new Set<string>();
    for (const node of nodes.values()) {
        const def = node.row.def;
        const attach = (codes: string[], out: AccountNode[]) => {
            for (const code of codes) {
                const child = nodes.get(code);
                if (!child || child === node || owned.has(code)) continue;
                out.push(child);
                owned.add(code);
            }
        };
        attach(def.children, node.children);
        attach(model.attached.get(def.code) ?? [], def.type === "step" || def.type === "subtotal" ? node.following : node.children);
        attach(model.following.get(def.code) ?? [], node.following);
    }
    // 合計の直後に置いた率は、その合計と一緒に移動する。参照する分子だけでは置き場所を決めない。
    let anchor: AccountNode | undefined;
    for (const node of nodes.values()) {
        if (node.row.def.type === "calc") {
            if (anchor && !owned.has(node.row.def.code)) { anchor.following.push(node); owned.add(node.row.def.code); }
        } else anchor = node.row.def.type === "step" || node.row.def.type === "subtotal" ? node : undefined;
    }
    let roots = Array.from(nodes.values()).filter(n => !owned.has(n.row.def.code) && n.row.def.type !== "step");
    const order = new Map(model.display.map((r, i) => [r.def.code, i]));
    const first = (n: AccountNode): number => Math.min(order.get(n.row.def.code)!, ...n.children.map(first));
    const coverage = (n: AccountNode): Set<string> => n.row.def.type === "step"
        ? new Set(n.row.def.summands.map(s => s.code)) : new Set([n.row.def.code]);
    const warnings: string[] = [];
    for (const step of Array.from(nodes.values()).filter(n => n.row.def.type === "step")) {
        const wanted = coverage(step);
        const selected = roots.filter(n => Array.from(coverage(n)).every(code => wanted.has(code)));
        const found = new Set(selected.flatMap(n => Array.from(coverage(n))));
        // 交差する計算範囲は木にできない。計算行を独立したまま残し、明細は複製しない。
        if (wanted.size > 0 && found.size === wanted.size) {
            step.children = selected;
            roots = roots.filter(n => !selected.includes(n));
        } else if (wanted.size > 0) {
            warnings.push(`「${step.row.def.name}」の計算範囲がほかの合計と交差するため、独立した計算行として表示しました`);
        }
        roots.push(step);
        roots.sort((a, b) => first(a) - first(b));
    }
    const parents = Array.from(nodes.values()).filter(n => n.children.length > 0).map(n => n.row.def.code);
    return { roots, parents, warnings };
}

export interface HierarchyCell { row: TableRow; column: number; rowSpan: number }
export interface LayoutHeader {
    above?: boolean;
    key: string; label: string; kind: "account" | "org"; open?: boolean; fold?: string; org?: string | null; row?: TableRow;
}
export type LayoutPlan =
    | { kind: "table"; key: string; rows: TableRow[]; nameColumns: number; orgColumns?: number; orgHeader?: LayoutHeader }
    | { kind: "stack"; key: string; direction: SplitDirection; children: LayoutPlan[]; preserveTables?: boolean }
    | { kind: "frame"; key: string; style: HierarchyStyle; header: LayoutHeader; content: LayoutPlan };

export function planRows(plan: LayoutPlan): TableRow[] {
    return plan.kind === "table" ? plan.rows : plan.kind === "frame" ? planRows(plan.content) : plan.children.flatMap(planRows);
}
export function planHeaders(plan: LayoutPlan): LayoutHeader[] {
    return plan.kind === "table" ? [] : plan.kind === "frame" ? [plan.header, ...planHeaders(plan.content)] : plan.children.flatMap(planHeaders);
}
export function stack(key: string, children: LayoutPlan[], direction: SplitDirection = "vertical"): LayoutPlan {
    return { kind: "stack", key, children, direction };
}
export function tablePlan(key: string, rows: TableRow[]): LayoutPlan {
    const nameColumns = Math.max(1, ...rows.map(r => (r.nameColumn ?? 0) + 1));
    return { kind: "table", key, nameColumns, rows: rows.map(r => ({ ...r, nameSpan: nameColumns - (r.nameColumn ?? 0) })) };
}

/** 囲みのセグメントは連続した一つの表にまとめ、列見出しを繰り返さない。 */
export function compactOrgBoxes(plan: LayoutPlan): LayoutPlan {
    if (plan.kind === "frame" && plan.style === "columns") return compactOrgColumns(plan);
    const compatible = (p: LayoutPlan): boolean => p.kind === "table" || (p.kind === "stack" ? !p.preserveTables && p.direction === "vertical" && p.children.every(compatible)
        : p.header.kind === "org" && (p.style === "box" || p.key.endsWith(":total")) && compatible(p.content));
    if (plan.kind !== "frame" || plan.header.kind !== "org" || plan.style !== "box" || !compatible(plan)) return plan;
    const rows: TableRow[] = [];
    const cells = new Map<string, OrgCell>();
    let orgHeader: LayoutHeader | undefined;
    const hasOrgFrame = (p: LayoutPlan): boolean => p.kind === "frame" || (p.kind === "stack" && p.children.some(hasOrgFrame));
    const visit = (p: LayoutPlan, depth: number, rails: number[]): void => {
        if (p.kind === "table") { rows.push(...p.rows); return; }
        if (p.kind === "stack") { p.children.forEach(c => visit(c, depth, rails)); return; }
        const start = rows.length;
        const parent = hasOrgFrame(p.content);
        const root = p.header.org === null && parent;
        const hasTotal = (q: LayoutPlan): boolean => q.kind === "stack" && q.children.some(c =>
            c.kind === "frame" && c.header.key === `${p.header.key}:total`);
        const enclosing = parent && (!root || hasTotal(p.content));
        visit(p.content, enclosing ? depth + 1 : depth, enclosing ? [...rails, depth] : rails);
        if (rows.length === start) return;
        const toggle = p.header.fold !== undefined ? { path: p.header.fold, open: p.header.open! } : null;
        if (!parent) {
            const cell: OrgCell = { label: p.header.label, path: p.header.org ?? null, column: 0, colSpan: 1,
                rowSpan: rows.length - start, toggle, total: p.key.endsWith(":total"), depth, rails };
            rows[start] = { ...rows[start], orgCells: [cell] };
            cells.set(p.header.key, cell);
        } else {
            const total = cells.get(`${p.header.key}:total`);
            if (total) {
                total.toggle = toggle; total.label = p.header.label; total.depth = depth; total.rails = rails;
                total.ownAbove = rows[start].orgCells?.[0] !== total;
            }
            else if (root) orgHeader = p.header;
            else {
                const first = rows[start].orgCells?.[0];
                if (first) first.heads = [{ label: p.header.label, path: p.header.org ?? null, depth, toggle }, ...(first.heads ?? [])];
            }
        }
    };
    visit(plan, 0, []);
    const nameColumns = Math.max(1, ...rows.map(r => (r.nameColumn ?? 0) + (r.nameSpan ?? 1)));
    return { kind: "table", key: plan.key, rows: rows.map(r => ({ ...r, nameSpan: nameColumns - (r.nameColumn ?? 0) })), nameColumns, orgColumns: 1, orgHeader };
}

function compactOrgColumns(plan: LayoutPlan): LayoutPlan {
    const compatible = (p: LayoutPlan): boolean => p.kind === "table" || (p.kind === "stack" ? !p.preserveTables && p.direction === "vertical" && p.children.every(compatible)
        : p.header.kind === "org" && (p.style === "columns" || p.key.endsWith(":total")) && compatible(p.content));
    if (!compatible(plan)) return plan;
    const rows: TableRow[] = [];
    const cells: OrgCell[] = [];
    const visit = (p: LayoutPlan, column: number): void => {
        if (p.kind === "table") { rows.push(...p.rows.map(r => ({ ...r }))); return; }
        if (p.kind === "stack") { p.children.forEach(c => visit(c, column)); return; }
        const start = rows.length;
        visit(p.content, column + 1);
        if (rows.length === start) return;
        const cell: OrgCell = { identity: `${p.header.key}:${column}`, label: p.header.label, path: p.header.org ?? null, column, colSpan: 1,
            rowSpan: rows.length - start, toggle: p.header.fold !== undefined ? { path: p.header.fold, open: p.header.open! } : null,
            total: p.key.endsWith(":total"), depth: 0 };
        cells.push(cell);
        rows[start].orgCells = [cell, ...(rows[start].orgCells ?? [])];
    };
    visit(plan, 0);
    const orgColumns = Math.max(0, ...cells.map(c => c.column + 1));
    const nameColumns = Math.max(1, ...rows.map(r => (r.nameColumn ?? 0) + (r.nameSpan ?? 1)));
    // 子のないセルだけを残りの列まで伸ばす。途中の行から数値列がずれないようにする。
    for (const row of rows) {
        for (const cell of row.orgCells ?? []) {
            const deeper = (row.orgCells ?? []).some(other => other.column > cell.column);
            if (!deeper) cell.colSpan = orgColumns - cell.column;
        }
    }
    return { kind: "table", key: plan.key, rows: rows.map(r => ({ ...r, nameSpan: nameColumns - (r.nameColumn ?? 0) })), nameColumns, orgColumns };
}

export function accountPlan(
    roots: AccountNode[], values: Map<string, TableRow>, closed: Set<string>,
    options: (node: AccountNode, level: number) => HierarchyOptions, foldKey: (code: string) => string, key: string, rootDirection = "joined",
): LayoutPlan {
    const value = (node: AccountNode): TableRow | undefined => values.get(node.row.def.code);
    const ownRow = (node: AccountNode, level: number, column: number): TableRow | undefined => {
        const row = value(node);
        if (!row) return undefined;
        return { ...row, depth: level, bands: [], nameColumn: column,
            ...(node.children.length ? { open: !closed.has(row.code), fold: foldKey(row.code) } : {}) };
    };
    const flatten = (node: AccountNode, level: number, column: number, indent = 0): TableRow[] => {
        const own = ownRow(node, indent, column);
        if (!own) return [...node.children, ...node.following].flatMap(n => flatten(n, level, column, indent));
        const opt = options(node, level);
        if (node.children.length && opt.total !== "none") own.aggregatePosition = opt.total;
        const after = node.following.flatMap(n => flatten(n, level, column, indent + (n.row.def.type === "breakdown" ? 1 : 0)));
        if (!node.children.length || closed.has(own.code)) {
            if (node.children.length && opt.style === "box") own.bands = [{ level: indent, head: true, first: true, last: true, afterHead: false }];
            return [own, ...after];
        }
        const box = opt.style === "box";
        const horizontal = opt.style === "columns";
        const body = node.children.flatMap(n => flatten(n, level + 1, column + (horizontal ? 1 : 0), box ? indent + 1 : indent));
        if (!body.length) return [own, ...after];
        const header = { ...own, cells: own.cells.map(() => ({ text: "", tone: null as null })), type: "heading" as const };
        let result = opt.total === "top" ? [own, ...body] : opt.total === "bottom" ? [...body, own] : horizontal ? body : [header, ...body];
        if (horizontal) {
            body[0].hierarchyCells = [{ row: header, column, rowSpan: body.length }, ...(body[0].hierarchyCells ?? [])];
        }
        if (box) result = result.map((r, i) => ({ ...r, bands: [{ level: indent, head: r.code === own.code,
            first: i === 0, last: i === result.length - 1, afterHead: i > 0 && result[i - 1].code === own.code }, ...r.bands] }));
        return [...result, ...after];
    };
    const hasSplit = (node: AccountNode, level: number): boolean => !!value(node) && !closed.has(node.row.def.code) && node.children.length > 0 &&
        (options(node, level).style === "split" || node.children.some(n => hasSplit(n, level + 1)));
    const compile = (nodes: AccountNode[], level: number, at: string): LayoutPlan => {
        const pieces: LayoutPlan[] = [];
        let pending: TableRow[] = [];
        const flush = () => { if (pending.length) pieces.push(tablePlan(`${at}:${pieces.length}`, pending)); pending = []; };
        for (const node of nodes) {
            if (!hasSplit(node, level)) { pending.push(...flatten(node, level, 0)); continue; }
            flush();
            const own = ownRow(node, level, 0)!;
            const opt = options(node, level);
            if (opt.total !== "none") own.aggregatePosition = opt.total;
            const children = opt.style === "split"
                ? stack(`${own.key}:children`, node.children.map(n => compile([n], level + 1, `${own.key}:${n.row.def.code}`)), opt.direction)
                : compile(node.children, level + 1, `${own.key}:children`);
            const total = tablePlan(`${own.key}:total`, [{ ...own, open: undefined }]);
            const following = node.following.flatMap(n => flatten(n, level, 0, n.row.def.type === "breakdown" ? 1 : 0));
            const content = stack(`${own.key}:content`, [...(opt.total === "none" ? [children] : opt.total === "top" ? [total, children] : [children, total]),
                ...(following.length ? [tablePlan(`${own.key}:following`, following)] : [])]);
            pieces.push({ kind: "frame", key: own.key, style: opt.style,
                header: { key: `frame:${own.key}`, label: own.name, kind: "account", open: true, fold: own.fold, row: own }, content });
        }
        flush();
        return pieces.length === 1 ? pieces[0] : stack(at, pieces);
    };
    return roots.length > 1 && rootDirection !== "joined"
        ? { ...stack(key, roots.map(n => compile([n], 0, `${key}:${n.row.def.code}`)), rootDirection === "horizontal" ? "horizontal" : "vertical"), preserveTables: true } as LayoutPlan : compile(roots, 0, key);
}
