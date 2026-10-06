/**
 * 組織のブロックと、行の折りたたみ。
 *
 * - 組織の欄（最大 3 段）の道筋から木を組む。一番上の段が 2 つ以上なら、それを束ねる「全社」を根に置く（名前は書式ペイン）
 * - 開いた組織は、子の組織のブロックと、自分の計のブロック（子を足した表）を並べる。計は子の後ろか前（書式ペイン、区分の小計の位置とは別）。
 *   閉じた組織は子を隠して計の表だけ。子の無い組織は自分の表で、閉じない
 * - 組織の名前は左の列に、段ごとに 1 列で出し、その組織の行にまとめて縦にかける（旧作の M2）。計の名前は残りの段の列をまたぐ
 * - 既定は根だけ開く：一番上の段（事業部）の計の表が並び、その中（事業）は隠れる。
 *   見る人が開いた組織を保存する（visualState.ts）。根は閉じない
 * - 空の段は「（組織なし）」。子が空の段 1 つだけで、その下に名前のある段が無ければ、その組織の表にする（事業に部門を入れない行だけの事業）
 * - ブロックの値は、その組織より下の道筋のファクトだけで計算する（viewModel.ts）。最新見込みは一番下の段の組織ごとに採るので、足した値と同じ
 *
 * 計の置き場所は区分と独立に選べる・左の列に縦にかける・一番上の段まで開く・閉じた組織は計の表
 */
import { ORG_SEPARATOR } from "./data";
import { UNDER_KEY, isAccountRow } from "./picks";
import { DisplayRow, RowDef } from "./rows";

export const NO_ORG_LABEL = "（セグメントなし）";
/** 一番上の段を束ねた計の名前の既定 */
export const DEFAULT_ROOT_NAME = "合計";

export interface OrgNode {
    /** 道筋（Fact.org の頭）。束ねる根は null */
    path: string | null;
    /** 表に出す名前（空の段は「（組織なし）」） */
    label: string;
    /** 左の何列目か（組織の段。束ねる根は -1） */
    column: number;
    children: OrgNode[];
}

/** 左の列のセル。column から colSpan 列、rowSpan 行にかかる */
export interface OrgCell {
    /** 同じ組織の見出しを別の階層列にも置くときの表示セルの識別子。 */
    identity?: string;
    label: string;
    /** 押したときに絞る組織の道筋（束ねる根の全社は null で、絞らない） */
    path: string | null;
    column: number;
    colSpan: number;
    rowSpan: number;
    /** 開き閉じのしるし。根・子の無い組織には付けない（1 列に畳むと、開いた組織のしるしは計の名前に付ける。singleColumn） */
    toggle: { path: string; open: boolean } | null;
    /** 計（子を足したブロック）の名前 */
    total: boolean;
    /** 1 列に畳んだとき（singleColumn）：名前の字下げの段、縦線を引く親の段、上に重ねる親の名前 */
    depth?: number;
    rails?: number[];
    /** ブロックの上の横線を引き始める縦線の段（上のブロックと共通する一番近い親の段） */
    lineFrom?: number;
    /** 囲む親の段のうち、箱がこのブロックで始まる段・終わる段（入れ子の箱の上と下の辺） */
    boxStart?: number[];
    boxEnd?: number[];
    /** 次のブロックが深い（自分の名前の所の下の辺を引く） */
    ownBottom?: boolean;
    /** 自分の箱の帯が上のブロックから続いてくる（小計が下の親の小計）・下のブロックへ続く（小計が上）。帯の幅には横線を引かない */
    ownAbove?: boolean;
    ownBelow?: boolean;
    /** 囲みなし：箱の塗りと線を引かない（名前と字下げだけ） */
    plain?: boolean;
    /** 囲む親の段のうち、すぐ上のブロックがその親の小計（小計が上）の段。帯の上の辺を引く */
    afterHead?: number[];
    heads?: OrgHeadLine[];
}

/** 表の行のかたまり：1 つの組織の表（どれも同じ行の数）。子のある組織を開いていれば計の表、閉じていれば子を隠した計の表 */
export interface OrgSegment {
    node: OrgNode;
    /** このかたまりの最初の行から始まる左の列のセル（上の段の名前のセルも、ここから始まるものは入る） */
    cells: OrgCell[];
}

export type OrgTotalPosition = "after" | "before";

/**
 * 道筋から組織の木を組む。orgs は届いた順（Power BI の並べ替えの順）。組織を入れていなければ（段が 0）null
 */
export function buildOrgTree(orgs: string[], levels: number, rootName: string): { root: OrgNode; columns: number } | null {
    if (levels === 0 || orgs.length === 0) return null;
    const top: OrgNode[] = [];
    const byPath = new Map<string, OrgNode>();
    for (const org of orgs) {
        const segments = org.split(ORG_SEPARATOR);
        let siblings = top;
        for (let i = 0; i < segments.length; i++) {
            const path = segments.slice(0, i + 1).join(ORG_SEPARATOR);
            let node = byPath.get(path);
            if (!node) {
                node = {
                    path,
                    label: segments[i] === "" ? NO_ORG_LABEL : segments[i],
                    column: i,
                    children: [],
                };
                byPath.set(path, node);
                siblings.push(node);
            }
            siblings = node.children;
        }
    }
    // 子が空の段 1 つだけで、その下に名前のある段も無ければ、その組織の表にする（下から詰める。空の段が続けば続く限り）。
    // 下に名前のある段がある空の段（会社 › （空）› 事業A）は「（組織なし）」の組織として残す（詰めると事業A が 1 段上がり、左の列に
    // 穴が開いて値の列がずれた）
    const prune = (node: OrgNode) => {
        node.children.forEach(prune);
        const only = node.children.length === 1 ? node.children[0] : null;
        if (only && only.children.length === 0 && only.path!.endsWith(ORG_SEPARATOR)) node.children = [];
    };
    top.forEach(prune);
    let columns = 0;
    const measure = (node: OrgNode) => {
        columns = Math.max(columns, node.column + 1);
        node.children.forEach(measure);
    };
    top.forEach(measure);
    const root: OrgNode = top.length === 1 ? top[0] : { path: null, label: rootName, column: -1, children: top };
    return { root, columns };
}

/** 開き閉じできる組織の道筋（根のほかの、子のある組織。すべて開くに使う） */
export function toggleablePaths(root: OrgNode): string[] {
    const paths: string[] = [];
    const walk = (node: OrgNode) => {
        if (node !== root && node.path !== null && node.children.length > 0) paths.push(node.path);
        node.children.forEach(walk);
    };
    walk(root);
    return paths;
}

/** ファクトの組織（道筋）がその組織の下か */
export function containsOrg(node: OrgNode, org: string): boolean {
    return node.path === null || org === node.path || org.startsWith(node.path + ORG_SEPARATOR);
}

/**
 * 表の行のかたまりの並びと、左の列のセル。open は開いた組織の道筋（根はいつも開く）。tableRows は組織の表の行の数（その組織にデータの無い
 * 科目を隠すと組織ごとに違う）
 */
export function layoutOrgs(
    root: OrgNode,
    columns: number,
    open: ReadonlySet<string>,
    position: OrgTotalPosition,
    tableRows: number | ((node: OrgNode) => number)
): OrgSegment[] {
    const segments: OrgSegment[] = [];
    const rowsPerTable = (node: OrgNode) => (typeof tableRows === "number" ? tableRows : tableRows(node));
    // 子の無い組織は開き閉じしない（いつも自分の表）
    const isOpen = (node: OrgNode) => node.children.length > 0 && (node === root ? !open.has(ROOT_FOLD) : node.path !== null && open.has(node.path));
    // 開いた組織はいつも計の表を出す。1 列に畳むと、開いた親の名前・開き閉じ・選択は計の表のセルが持つので、子が 1 つだけでも出す
    // 。根だけは、子が 1 つなら出さない（根は閉じないので、
    // 名前を子の上に重ねれば足りる。表全体が 2 度並ぶのを避ける）。束ねる根はいつも子が 2 つ以上
    const hasTotal = (node: OrgNode) => node.children.length > 1 || node.path === null || (node !== root && node.children.length > 0);
    const rowsOf = (node: OrgNode): number =>
        isOpen(node) ? node.children.reduce((sum, child) => sum + rowsOf(child), 0) + (hasTotal(node) ? rowsPerTable(node) : 0) : rowsPerTable(node);
    const toggleOf = (node: OrgNode, openNow: boolean) =>
        node.children.length === 0 ? null : node === root ? { path: ROOT_FOLD, open: openNow } : node.path === null ? null : { path: node.path, open: openNow };
    const start = (node: OrgNode) => Math.max(0, node.column);

    /** carry：上の段から持ち越した、このかたまりの最初の行から始まるセル */
    const emit = (node: OrgNode, carry: OrgCell[]) => {
        const col = start(node);
        if (!isOpen(node)) {
            // 子の無い組織の表、閉じた組織の計の表（子を隠す）。名前は残りの段の列をまたぐ
            const cell: OrgCell = {
                label: node.label,
                path: node.path,
                column: col,
                colSpan: columns - col,
                rowSpan: rowsPerTable(node),
                toggle: toggleOf(node, false),
                total: false,
            };
            segments.push({ node, cells: [...carry, cell] });
            return;
        }
        const childRows = node.children.reduce((sum, child) => sum + rowsOf(child), 0);
        // 束ねる根は名前の列を持たない（全社の列が表の端から端までかかるだけになる）
        const nameCell: OrgCell | null =
            node.path === null ? null : { label: node.label, path: node.path, column: col, colSpan: 1, rowSpan: childRows, toggle: toggleOf(node, true), total: false };
        const totalCell: OrgCell = {
            label: node.path === null ? node.label : `${node.label}計`,
            path: node.path,
            column: col,
            colSpan: columns - col,
            rowSpan: rowsPerTable(node),
            toggle: null,
            total: true,
        };
        const children = (first: OrgCell[]) => node.children.forEach((child, i) => emit(child, i === 0 ? first : []));
        // 子が 1 つだけなら計の表は子の表と同じなので出さない（第一事業部の下が機器事業だけなど）
        if (!hasTotal(node) && nameCell) {
            children([...carry, nameCell]);
            return;
        }
        if (position === "before") {
            segments.push({ node, cells: [...carry, totalCell] });
            children(nameCell ? [nameCell] : []);
        } else {
            children(nameCell ? [...carry, nameCell] : carry);
            segments.push({ node, cells: [totalCell] });
        }
    };
    emit(root, []);
    return segments;
}

/**
 * 根（合計）の開き閉じの名前。見る人の開いた組織の一覧（openOrgs）に入っていれば根を閉じる（根だけ既定が開き。ほかの組織は入っていれば開く）。
 * 閉じた根は、子を隠した合計の表だけになる
 */
export const ROOT_FOLD = "\u0000root";

/** 組織のブロックの名前（ツールチップ）。1 列に畳んだ名前のセルと同じく、開いた組織の計も名前だけ（「計」は付けない） */
export function blockName(segment: OrgSegment): string {
    return segment.node.label;
}

/** 組織ごとの、データ（ファクト）のある科目のコード。束ねる根（null）はすべての組織 */
export function codesByOrg(root: OrgNode, facts: ReadonlyArray<{ code: string; org: string }>): Map<string | null, Set<string>> {
    const result = new Map<string | null, Set<string>>();
    const add = (key: string | null, code: string) => {
        let set = result.get(key);
        if (!set) result.set(key, (set = new Set()));
        set.add(code);
    };
    const paths = new Set<string>();
    const walk = (node: OrgNode) => {
        if (node.path !== null) paths.add(node.path);
        node.children.forEach(walk);
    };
    walk(root);
    for (const fact of facts) {
        add(null, fact.code);
        // 道筋の頭（上の段から）のうち、木にある組織すべて
        const segments = fact.org.split(ORG_SEPARATOR);
        for (let i = 1; i <= segments.length; i++) {
            const path = segments.slice(0, i).join(ORG_SEPARATOR);
            if (paths.has(path)) add(path, fact.code);
        }
    }
    return result;
}

/**
 * その組織にデータの無い科目の行。値が 0 かではなく、その組織に
 * その科目のファクトがあるかで決める（データがあって値が 0 の科目は残す）。科目が全部隠れた区分・中分類の行、隠れた科目のうち・子も隠す。
 * 計算行・指標・後ろに置いた指標は残す
 */
export function emptyAccountRows(
    display: DisplayRow[],
    attached: ReadonlyMap<string, string[]>,
    has: (code: string) => boolean,
    defs?: ReadonlyMap<string, RowDef>
): Set<string> {
    const shown = new Map(display.map((d) => [d.def.code, d.def]));
    // 並びに無い行（小計をうちにして中身を出さない行など）も、行の定義から引く
    const byCode = { get: (code: string) => shown.get(code) ?? defs?.get(code), has: (code: string) => shown.has(code) || (defs?.has(code) ?? false) };
    const memo = new Map<string, boolean>();
    const empty = (code: string): boolean => {
        const known = memo.get(code);
        if (known !== undefined) return known;
        const def = byCode.get(code);
        let result = false;
        if (isAccountRow(def)) result = !has(code);
        // 見る人が小計をうちにした行：元の小計が空なら空
        else if (def?.type === "breakdown" && code.startsWith(UNDER_KEY)) result = def.summands.length > 0 && def.summands.every((s) => empty(s.code));
        // 見る人が選ばなかった科目をまとめた「その他」：まとめた科目がどれも空
        // まとめた行が小計（中分類）なら、その小計が空か
        else if (def?.type === "others") result = def.summands.every((s) => (byCode.has(s.code) ? empty(s.code) : !has(s.code)));
        else if (def && (def.type === "subtotal" || def.type === "heading")) {
            // 区分・中分類：子の科目・中分類と、足す行（見る人がその他・うちにして並びから外した科目も）がどれも空。指標の子・うちは数えない。
            // 足す行を見ないと、うちにした科目だけが空の区分が、ほかの科目にデータがあっても消えた
            const parts = Array.from(new Set([...def.children, ...(attached.get(code) ?? []), ...def.summands.map((s) => s.code)]));
            const accounts = parts.filter((c) => {
                const child = byCode.get(c);
                // 並びに無い足す行は、見る人が隠した科目
                return child === undefined || isAccountRow(child) || child.type === "others" || child.type === "subtotal" || child.type === "heading";
            });
            result = accounts.length > 0 && accounts.every((c) => (byCode.has(c) ? empty(c) : !has(c)));
        }
        memo.set(code, result);
        return result;
    };
    const hidden = new Set<string>();
    const hideBelow = (code: string) => {
        for (const child of [...(byCode.get(code)?.children ?? []), ...(attached.get(code) ?? [])]) {
            if (hidden.has(child)) continue;
            hidden.add(child);
            hideBelow(child);
        }
    };
    for (const d of display) {
        if (!empty(d.def.code)) continue;
        hidden.add(d.def.code);
        hideBelow(d.def.code);
    }
    return hidden;
}

/** 折りたためる行：子のある小計・見出しの行（囲みの帯を引く行と同じ） */
export function foldable(row: DisplayRow): boolean {
    return row.def.children.length > 0 && (row.def.type === "subtotal" || row.def.type === "heading");
}

/**
 * 閉じた行の下に隠す行のコード。子・孫と、そのうち・後ろに置いた指標の並びの端から端まで（あいだに置いた指標も）。
 * 閉じた行そのもののうち（売上高のうち新製品）と、閉じた行の後ろに置いた指標は隠さない（小計の行の中身・行の外なので）。
 * 小計の位置が上でも下でも同じ行を隠す（上のとき、最後の子の後ろに置いた指標が範囲の外に残った）
 */
export function hiddenRows(
    display: DisplayRow[],
    attached: ReadonlyMap<string, string[]>,
    closed: ReadonlySet<string>,
    following: ReadonlyMap<string, string[]> = new Map()
): Set<string> {
    const hidden = new Set<string>();
    if (closed.size === 0) return hidden;
    const position = new Map(display.map((d, i): [string, number] => [d.def.code, i]));
    const byCode = new Map(display.map((d) => [d.def.code, d.def]));
    display.forEach((row, index) => {
        if (!closed.has(row.def.code) || !foldable(row)) return;
        let from = index;
        let to = index;
        // 閉じた行そのもののうち・後ろに置いた指標は、その下（うち・後ろ・子）ごと、範囲の中にあっても隠さない（小計の位置が上のとき、
        // 範囲の中に置かれる。その指標のうちが隠れた）
        const keep = new Set<string>();
        const keepWalk = (code: string) => {
            for (const next of [...(byCode.get(code)?.children ?? []), ...(attached.get(code) ?? []), ...(following.get(code) ?? [])]) {
                if (keep.has(next)) continue;
                keep.add(next);
                keepWalk(next);
            }
        };
        for (const own of [...(attached.get(row.def.code) ?? []), ...(following.get(row.def.code) ?? [])]) {
            keep.add(own);
            keepWalk(own);
        }
        const walk = (code: string, own: boolean) => {
            // 閉じた行そのもののうち・後ろは入れない。子・孫のうち・後ろは入れる
            const below = [...(byCode.get(code)?.children ?? []), ...(own ? [] : [...(attached.get(code) ?? []), ...(following.get(code) ?? [])])];
            for (const child of below) {
                const at = position.get(child);
                if (at === undefined) continue;
                from = Math.min(from, at);
                to = Math.max(to, at);
                walk(child, false);
            }
        };
        walk(row.def.code, true);
        for (let i = from; i <= to; i++) if (i !== index && !keep.has(display[i].def.code)) hidden.add(display[i].def.code);
    });
    return hidden;
}

/** 1 列に畳んだセグメントの名前のセルに、上に重ねて出す親の名前（計の表の無い親。子が 1 つだけの根） */
export interface OrgHeadLine {
    label: string;
    path: string | null;
    toggle: { path: string; open: boolean } | null;
    /** 字下げの段（束ねる根が 0） */
    depth: number;
}

/**
 * セグメントを 1 列に畳む。段ごとの列（layoutOrgs）の
 * かたまりごとに、そのかたまりのセル 1 つにする：名前は段の深さで字下げし、開いた親の名前と開き閉じのしるしは、その親の計の表に出す（計は
 * 付けない）。計の表の無い親（子が 1 つだけの根）は、子の名前の上に重ねる。rails は、このかたまりを囲む親の段（縦線を引く）。
 * 縦線は、小計の位置が上でも下でも子のブロックにだけ引く
 */
export function singleColumn(segments: OrgSegment[], root: OrgNode): OrgCell[][] {
    const depthOf = (column: number) => column - root.column;
    // 計の表のある親の道筋（その親の名前は計の表に出す）
    const totals = new Set(segments.filter((s) => s.cells.at(-1)?.total).map((s) => s.node.path));
    // ブロックの切れ目の横線は、上下のブロックに共通する一番近い親の縦線から引く。子の箱の底（最後の子と親の小計のあいだ）は親の縦線から、
    // 兄弟の箱をまたぐ所（子会社の小計と全社の子のあいだ）は外の親の縦線から
    const parent = new Map<OrgNode, OrgNode | null>();
    const walk = (node: OrgNode, up: OrgNode | null) => {
        parent.set(node, up);
        node.children.forEach((child) => walk(child, node));
    };
    walk(root, null);
    const chain = (node: OrgNode) => {
        const result: OrgNode[] = [];
        for (let at: OrgNode | null = node; at; at = parent.get(at) ?? null) result.push(at);
        return result;
    };
    // 段 d の親（このブロックを囲む箱）。ブロックの節は、計のブロックなら親そのもの
    const ancestorAt = (node: OrgNode, d: number) => chain(node).find((n) => depthOf(n.column) === d) ?? null;
    const commonDepth = (a: OrgNode, b: OrgNode) => {
        const above = new Set(chain(a));
        const shared = chain(b).find((n) => above.has(n));
        return shared ? depthOf(shared.column) : 0;
    };
    return segments.map((segment, index) => {
        const own = segment.cells.at(-1)!;
        const { node } = segment;
        const depth = depthOf(node.column);
        const heads: OrgHeadLine[] = segment.cells
            .slice(0, -1)
            .filter((c) => !c.total && !totals.has(c.path))
            .map((c) => ({ label: c.label, path: c.path, toggle: c.toggle, depth: depthOf(c.column) }));
        // 開いた親の計の表の名前に開き閉じを付ける（根は ROOT_FOLD）
        const toggle = own.total ? (node === root ? { path: ROOT_FOLD, open: true } : node.path === null ? null : { path: node.path, open: true }) : own.toggle;
        const cell: OrgCell = {
            label: node.label,
            path: own.path,
            column: 0,
            colSpan: 1,
            rowSpan: own.rowSpan,
            toggle,
            total: own.total,
            depth,
            rails: Array.from({ length: depth }, (_, i) => i),
            lineFrom: index === 0 ? 0 : commonDepth(segments[index - 1].node, node),
            boxStart: Array.from({ length: depth }, (_, d) => d).filter((d) => index === 0 || ancestorAt(segments[index - 1].node, d) !== ancestorAt(node, d)),
            boxEnd: Array.from({ length: depth }, (_, d) => d).filter(
                (d) => index === segments.length - 1 || ancestorAt(segments[index + 1].node, d) !== ancestorAt(node, d)
            ),
            ownBottom: index < segments.length - 1 && depthOf(segments[index + 1].node.column) > depth,
            ownAbove: index > 0 && segments[index - 1].node !== node && chain(segments[index - 1].node).includes(node),
            afterHead: Array.from({ length: depth }, (_, d) => d).filter((d) => index > 0 && segments[index - 1].node === ancestorAt(node, d)),
            ownBelow: index < segments.length - 1 && segments[index + 1].node !== node && chain(segments[index + 1].node).includes(node),
            heads,
        };
        return [cell];
    });
}
