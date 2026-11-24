import { maxOf, minOf } from "./shared/minMax";
import { INDICATOR_KEY, isIndicatorLevel } from "./data";
import { UNDER_KEY } from "./picks";
import { TOTAL_KEY, type DisplayRow, type RowModel } from "./rows";
import type { TableRow } from "./viewModel";
import type { HierarchyOptions, HierarchyStyle, SplitDirection } from "./hierarchySettings";
import type { OrgCell, OrgParent } from "./orgs";

/** 計算式に手を加えず、表示に使う親子だけを組み立てる。 */
export interface AccountNode { row: DisplayRow; children: AccountNode[]; following: AccountNode[] }
export interface AccountTree { roots: AccountNode[]; parents: string[]; warnings: string[] }

/**
 * stepParents：計算行（売上総利益・営業利益…）を、足す範囲の区分の親にするか。する（横積みで親の名前を左に並べる形）と、損益計算書は段階利益の入れ子になる。
 * しないと、計算行は区分と並ぶ 1 行（1.x と同じ見え方。閉じても区分が残る）。根の合計行は、いつも区分の親
 */
export function accountTree(model: RowModel, stepParents = true): AccountTree {
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
        // 科目・指標の行の子はうち（rows.ts）なので、いつも親の下に続ける（集計の行として下に回さない）
        // 小計をうちにした行（UNDER_KEY）の子は、元の小計の中身なので子にする（開き閉じできる）
        // 指標の行の下の「行」の階層（rows.ts の emitLevels）は子にする（開き閉じできる）
        const levels = def.type === "measure" ? def.children.filter(isIndicatorLevel) : [];
        attach(levels, node.children);
        attach(def.children.filter(code => !levels.includes(code)), def.type === "detail" || (def.type === "breakdown" && !def.code.startsWith(UNDER_KEY)) || def.type === "measure" ? node.following : node.children);
        // 合計行のうち（見る人が区分をうちにした行）は子にする（合計行の囲みの設定が、うちの行に効く）
        // 見る人がうちにした行（科目・小計）は子と同じ扱い（親の囲みと展開時の合計の位置が効く）。
        // 書式ペインで指標を「行のうち」に置いた行は、いつも親の下に続ける
        const attached = model.attached.get(def.code) ?? [];
        attach(attached.filter(code => !code.startsWith(INDICATOR_KEY)), node.children);
        attach(attached.filter(code => code.startsWith(INDICATOR_KEY)), node.following);
        attach(model.following.get(def.code) ?? [], node.following);
    }
    // 合計の直後に置いた率は、その合計と一緒に移動する。参照する分子だけでは置き場所を決めない。
    let anchor: AccountNode | undefined;
    for (const node of nodes.values()) {
        if (node.row.def.type === "calc") {
            if (anchor && !owned.has(node.row.def.code)) { anchor.following.push(node); owned.add(node.row.def.code); }
        } else anchor = node.row.def.type === "step" || node.row.def.type === "subtotal" ? node : undefined;
    }
    let roots = Array.from(nodes.values()).filter(n => !owned.has(n.row.def.code) && (n.row.def.type !== "step" || (!stepParents && n.row.def.code !== TOTAL_KEY)));
    const order = new Map(model.display.map((r, i) => [r.def.code, i]));
    const first = (n: AccountNode): number => Math.min(order.get(n.row.def.code)!, minOf(n.children.map(first)));
    // 見る人がうち・その他にして表から外した行は数えない（数えると、範囲が交差したと誤って知らせた）
    // その他・うちにした小計の中身は、開くと見えるよう表に残すが、その他・うちの行の子なので数えない
    const swallowed = new Set(Array.from(nodes.values()).filter(n => n.row.def.type === "others" || n.row.def.code.startsWith(UNDER_KEY)).flatMap(n => n.row.def.children));
    const coverage = (n: AccountNode): Set<string> => n.row.def.type === "step"
        ? new Set(n.row.def.summands.map(s => s.code).filter(code => nodes.has(code) && !swallowed.has(code))) : new Set([n.row.def.code]);
    const warnings: string[] = [];
    // 範囲の小さい計算行から親にする（合計行を上に置くと表の先頭に来るので、並びの順だと合計行が先に区分を取り、計算行がその下に入った）
    const steps = Array.from(nodes.values()).filter(n => n.row.def.type === "step" && (stepParents || n.row.def.code === TOTAL_KEY));
    steps.sort((a, b) => coverage(a).size - coverage(b).size || (a.row.def.code === TOTAL_KEY ? 1 : 0) - (b.row.def.code === TOTAL_KEY ? 1 : 0));
    for (const step of steps) {
        // 自分のうちにした行（合計行のうち）は、もう子なので範囲に数えない
        const own = new Set([...step.children, ...step.following].map(c => c.row.def.code));
        const wanted = new Set(Array.from(coverage(step)).filter(code => !own.has(code)));
        // 計算行を親にしないときは、合計行も計算行を子にしない（閉じても計算行は区分と並んで残る）
        const selected = roots.filter(n => (stepParents || n.row.def.type !== "step") && Array.from(coverage(n)).every(code => wanted.has(code)));
        const found = new Set(selected.flatMap(n => Array.from(coverage(n))));
        // 交差する計算範囲は木にできない。計算行を独立したまま残し、明細は複製しない。
        if (wanted.size > 0 && found.size === wanted.size) {
            step.children = [...step.children, ...selected];
            roots = roots.filter(n => !selected.includes(n));
        } else if (wanted.size > 0) {
            warnings.push(`「${step.row.def.name}」の計算範囲がほかの合計と交差するため、独立した計算行として表示しました`);
        }
        roots.push(step);
        roots.sort((a, b) => first(a) - first(b));
    }
    // 子は表の並び順（見る人がまとめた「その他」を後から足しても、並びの位置に置く）
    for (const node of nodes.values()) node.children.sort((a, b) => first(a) - first(b));
    // 合計行は、うちだけでも配置（囲み）の対象にする（見る人が区分をうちにすると子が無くなる）
    const parents = Array.from(nodes.values()).filter(n => n.children.length > 0 || (n.row.def.code === TOTAL_KEY && n.following.length > 0)).map(n => n.row.def.code);
    return { roots, parents, warnings };
}

export interface HierarchyCell { row: TableRow; column: number; rowSpan: number }
export interface LayoutHeader {
    above?: boolean;
    /** 名前を出さず、開き閉じのしるしだけにする（分割で合計なしを開いているとき。合計の行があるように読めるので） */
    hideLabel?: boolean;
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
    const nameColumns = Math.max(1, maxOf(rows.map(r => (r.nameColumn ?? 0) + 1)));
    return { kind: "table", key, nameColumns, rows: rows.map(r => ({ ...r, nameSpan: nameColumns - (r.nameColumn ?? 0) })) };
}

/** 囲みのセグメントは連続した一つの表にまとめ、列見出しを繰り返さない。横積みと囲みが段ごとに混ざっていても 1 つの表にする（compactOrgMixed） */
export function compactOrgBoxes(plan: LayoutPlan): LayoutPlan {
    const pure = compactOrgPure(plan);
    return pure === plan ? compactOrgMixed(plan) : pure;
}

/**
 * 横積みだけのセグメントは段ごとの列に並べる。囲み・囲みなしは compactOrgMixed に任せる（以前は一番上の段の形だけを見てまとめ、途中の段の
 * 「囲みなし」を囲みとして描いていた。2026-10-29）
 */
function compactOrgPure(plan: LayoutPlan): LayoutPlan {
    return plan.kind === "frame" && plan.style === "columns" ? compactOrgColumns(plan) : plan;
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
        // 全社の合計を出さない根（名前を隠した根）は列を持たない。子のセグメントが一番左の列になる
        if (p.header.hideLabel && p.header.org === null) { visit(p.content, column); return; }
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
    const orgColumns = Math.max(0, maxOf(cells.map(c => c.column + 1)));
    const nameColumns = Math.max(1, maxOf(rows.map(r => (r.nameColumn ?? 0) + (r.nameSpan ?? 1))));
    // 子のないセルだけを残りの列まで伸ばす。途中の行から数値列がずれないようにする。
    for (const row of rows) {
        for (const cell of row.orgCells ?? []) {
            const deeper = (row.orgCells ?? []).some(other => other.column > cell.column);
            if (!deeper) cell.colSpan = orgColumns - cell.column;
        }
    }
    return { kind: "table", key: plan.key, rows: rows.map(r => ({ ...r, nameSpan: nameColumns - (r.nameColumn ?? 0) })), nameColumns, orgColumns };
}

/**
 * 横積みと囲み（囲みなし）が段ごとに混ざったセグメント（「事業部だけ横積み、事業は囲み」など）を 1 つの表にする。
 * 横積みの段は自分の列に名前を置き、子を次の列に並べる。囲みの段は今の列の中で字下げと帯で入れ子にする（compactOrgBoxes と同じ）。
 * 以前はどちらのまとめ方にも合わず、子のセグメントごとに見出し付きの別の表に分かれた（2026-10-21 ユーザー「セグメントを 2 段重ねて、一段目だけ横並びにしたときに崩れた」）。
 * 分割・左右の並びが入っていれば、まとめずにそのまま返す
 */
function compactOrgMixed(plan: LayoutPlan): LayoutPlan {
    const compatible = (p: LayoutPlan): boolean => p.kind === "table" || (p.kind === "stack" ? !p.preserveTables && p.direction === "vertical" && p.children.every(compatible)
        : p.header.kind === "org" && (p.style === "box" || p.style === "plain" || p.style === "columns" || p.key.endsWith(":total")) && compatible(p.content));
    if (plan.kind !== "frame" || plan.header.kind !== "org" || !compatible(plan)) return plan;
    const rows: TableRow[] = [];
    const cells = new Map<string, OrgCell>();
    const all: OrgCell[] = [];
    let orgHeader: LayoutHeader | undefined;
    const hasOrgFrame = (p: LayoutPlan): boolean => p.kind === "frame" || (p.kind === "stack" && p.children.some(hasOrgFrame));
    const toggleOf = (p: Extract<LayoutPlan, { kind: "frame" }>) => p.header.fold !== undefined ? { path: p.header.fold, open: p.header.open! } : null;
    const place = (cell: OrgCell, start: number) => {
        all.push(cell);
        rows[start] = { ...rows[start], orgCells: [cell, ...(rows[start].orgCells ?? [])] };
    };
    // names：同じ列の中で囲む親（rails の段）の名前（compactOrgPure と同じ）。横積みの段は自分の列に名前があるので、次の列では空から
    // plains：rails のうち囲みなしの段（自分の塗りを持たず、外の親の塗りのまま。左の辺と字下げは囲みと同じ）
    // edgeless：囲みの左の辺を引かない段（囲みなしの親のすぐ下の子の段）
    const visit = (p: LayoutPlan, column: number, depth: number, rails: number[], names: OrgParent[], plains: number[] = [], edgeless: number[] = []): void => {
        if (p.kind === "table") { rows.push(...p.rows.map(r => ({ ...r }))); return; }
        if (p.kind === "stack") { p.children.forEach(c => visit(c, column, depth, rails, names, plains, edgeless)); return; }
        const start = rows.length;
        // 計の表の枠（:total）は囲みなしの形で届くが、塗りと線は親の形で決める（下の Object.assign）。ここで囲みなしにすると、合計を下に置いたとき
        // 合計のセルだけ箱が切れた（2026-10-21 ユーザー「本部合計を下にしたときに囲みが区切られちゃってる」）
        const plain = p.style === "plain" && !p.key.endsWith(":total");
        if (p.style === "columns" && !p.key.endsWith(":total")) {
            // 全社の合計を出さない根（名前を隠した根）は列を持たない
            if (p.header.hideLabel && p.header.org === null) { visit(p.content, column, depth, rails, names); return; }
            visit(p.content, column + 1, 0, [], [], []);
            if (rows.length === start) return;
            place({ identity: `${p.header.key}:${column}`, label: p.header.label, labelKey: p.header.key, path: p.header.org ?? null, column, colSpan: 1,
                rowSpan: rows.length - start, toggle: toggleOf(p), total: false, depth, rails, parents: names }, start);
            return;
        }
        const parent = hasOrgFrame(p.content);
        const root = p.header.org === null && parent;
        const hasTotal = (q: LayoutPlan): boolean => q.kind === "stack" && q.children.some(c => c.kind === "frame" && c.header.key === `${p.header.key}:total`);
        // 囲みの親の中で囲みなしにしたノード（東日本本部だけ囲みなし）も、子を一段内側に入れ、左の辺を引く（塗りだけ外の親のまま）。囲まないと、子（第一事業部）が
        // 親と同じ字下げで並び、合計の囲みの左の辺も東日本本部の所で途切れた（2026-11-12 ユーザー）。外に囲みの親が無い囲みなし（全体を囲みなし）は今までどおり囲まない
        const insideBox = rails.some(r => !plains.includes(r));
        const enclosing = parent && (!plain || insideBox) && (!root || hasTotal(p.content));
        const nextPlains = enclosing && plain ? [...plains, depth] : plains;
        // 囲みなしの親の下の子は、子の数にかかわらず囲みの左の辺を引かない（子どうしの境の横線と字下げは残す）。子が 1 つなら親と同じ行をなぞる
        // だけで、2 つ以上でも境は横線と字下げで分かる（2026-11-12 ユーザー「これ縦線必要？第一事業部の左」「ここの縦線が必要なのかどうなのか」）
        const nextEdgeless = enclosing && plain ? [...edgeless, depth + 1] : edgeless;
        visit(p.content, column, enclosing ? depth + 1 : depth, enclosing ? [...rails, depth] : rails, enclosing ? [...names, { key: p.header.key, label: p.header.label }] : names, nextPlains, nextEdgeless);
        if (rows.length === start) return;
        if (!parent) {
            // 囲みなしの段（ノードの上書きで 1 つだけ囲みなしにしたものも）は、自分の箱を引かないだけで、外の親の帯と字下げは残す。段 0・帯なしに
            // すると、外が囲みのとき（東日本本部だけ囲みなし）、そのセルだけ合計の帯が切れて境の線が左にずれ、▾ が囲みの外に出た（2026-11-12 ユーザー）
            const cell: OrgCell = { label: p.header.label, labelKey: p.header.key, path: p.header.org ?? null, column, colSpan: 1, rowSpan: rows.length - start, toggle: toggleOf(p),
                total: p.key.endsWith(":total"), depth, rails, parents: names, ...(plain ? { plain: true } : {}), ...(plains.length ? { plainRails: plains } : {}), ...(edgeless.length ? { edgeless } : {}) };
            place(cell, start);
            cells.set(p.header.key, cell);
            return;
        }
        const total = cells.get(`${p.header.key}:total`);
        if (total) {
            Object.assign(total, { toggle: toggleOf(p), label: p.header.label, labelKey: p.header.key, depth, rails, parents: names,
                ownAbove: !rows[start].orgCells?.includes(total), ...(plains.length ? { plainRails: plains } : {}), ...(edgeless.length ? { edgeless } : {}) });
            if (plain) total.plain = true;
        } else if (root && column === 0) orgHeader = p.header;
        else {
            const first = rows[start].orgCells?.find(c => c.column === column);
            if (first) first.heads = [{ label: p.header.label, key: p.header.key, path: p.header.org ?? null, depth, toggle: toggleOf(p) }, ...(first.heads ?? [])];
        }
    };
    visit(plan, 0, 0, [], []);
    markBoxStarts(rows);
    const orgColumns = Math.max(1, maxOf(all.map(c => c.column + 1)));
    // 右に子の列が無いセル（囲みの葉・計）は残りの列まで伸ばす。数値の列がずれないように
    for (const row of rows) for (const cell of row.orgCells ?? []) {
        if (!(row.orgCells ?? []).some(other => other.column > cell.column)) cell.colSpan = orgColumns - cell.column;
    }
    const nameColumns = Math.max(1, maxOf(rows.map(r => (r.nameColumn ?? 0) + (r.nameSpan ?? 1))));
    return { kind: "table", key: plan.key, rows: rows.map(r => ({ ...r, nameSpan: nameColumns - (r.nameColumn ?? 0) })), nameColumns, orgColumns, orgHeader };
}

/**
 * 囲む親の段のうち、箱がこのブロックで始まる段（boxStart）。帯の上の辺を引く（App.tsx の NestedBox）。列ごとに上のブロックと比べ、その段で囲む
 * 箱（段が自分の段なら自分。計のブロックは親そのもの）が違えば始まる。付けていなかったので、帯の上の辺がどこにも引かれず、第一事業部の計の下に
 * 第二事業部の最初の子が続く所で、事業部の段の線が欠けて 2 つの事業部の箱がつながって見えた（2026-10-29 ユーザー）
 */
function markBoxStarts(rows: TableRow[]): void {
    const boxAt = (cell: OrgCell, d: number) => {
        const depth = cell.depth ?? 0;
        return d < depth ? cell.parents?.[d]?.key : d === depth ? cell.labelKey : undefined;
    };
    const above = new Map<number, OrgCell>();
    for (const row of rows) for (const cell of row.orgCells ?? []) {
        const previous = above.get(cell.column);
        cell.boxStart = (cell.rails ?? []).filter(d => !previous || boxAt(previous, d) !== boxAt(cell, d));
        above.set(cell.column, cell);
    }
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
        /**
         * うちを持つ行は、うちと一緒に 1 つの箱にする（1.x の囲みと同じ。うちの行だけがずれた箱にならない）。箱に入れるのは自分の直接のうちだけで、
         * 後ろに置いた率・指標は箱の外に続ける（後ろの指標が持つうちまで数えると、営業利益が率・指標ごと箱になった）
         */
        const withBreakdowns = (head: TableRow[]): TableRow[] => {
            const breakdowns = node.following.filter(n => n.row.def.type === "breakdown");
            if (opt.style !== "box" || breakdowns.length === 0) return [...head, ...after];
            const inBox = breakdowns.flatMap(n => flatten(n, level, column, indent + 1));
            const rest = node.following.filter(n => n.row.def.type !== "breakdown").flatMap(n => flatten(n, level, column, indent));
            const block = [...head, ...inBox];
            return [...block.map((r, i) => ({ ...r, bands: [{ level: indent, head: r.code === own.code, first: i === 0, last: i === block.length - 1,
                afterHead: i > 0 && block[i - 1].code === own.code }, ...r.bands] })), ...rest];
        };
        if (!node.children.length || closed.has(own.code)) {
            if (node.children.length && opt.style === "box") {
                own.bands = [{ level: indent, head: true, first: true, last: true, afterHead: false }];
                return [own, ...after];
            }
            return withBreakdowns([own]);
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
    // 最上位を分けた表には、セグメントの分割と同じく、上に名前の帯を置く（どの表が台数・利益かが分かるように。単位も帯の右に出る）。
    // 分割の囲み（frame）になった区分は、もう帯があるのでそのまま
    const titled = (n: AccountNode): LayoutPlan => {
        const plan = compile([n], 0, `${key}:${n.row.def.code}`);
        if (plan.kind === "frame") return plan;
        const row = value(n);
        return { kind: "frame", key: `root:${n.row.def.code}`, style: "split", content: plan,
            header: { key: `frame:root:${n.row.def.code}`, label: n.row.def.name, kind: "account", ...(row ? { row } : {}) } };
    };
    return roots.length > 1 && rootDirection !== "joined"
        ? { ...stack(key, roots.map(titled), rootDirection === "horizontal" ? "horizontal" : "vertical"), preserveTables: true } as LayoutPlan : compile(roots, 0, key);
}
