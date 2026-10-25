import powerbi from "powerbi-visuals-api";
import { formattingSettings as f } from "powerbi-visuals-utils-formattingmodel";

export type HierarchyStyle = "plain" | "box" | "split" | "columns";
export type TotalPosition = "none" | "top" | "bottom";
export type SplitDirection = "vertical" | "horizontal";
export interface HierarchyOptions { style: HierarchyStyle; total: TotalPosition; direction: SplitDirection }
export interface HierarchyOverride { target: string; style: string; total: string; direction: string }
export interface HierarchyFormat { targets: powerbi.IEnumMember[]; saved: HierarchyOverride[]; scopes?: Partial<Record<"account" | "org", string>> }
export const HIERARCHY_SLOTS = 20;
export const STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: "plain", displayName: "囲みなし" }, { value: "box", displayName: "囲み" },
    { value: "split", displayName: "分割" }, { value: "columns", displayName: "横積み" },
];
export const TOTAL_ITEMS: powerbi.IEnumMember[] = [
    { value: "none", displayName: "なし（畳むと小計を表示）" }, { value: "top", displayName: "上" }, { value: "bottom", displayName: "下" },
];
export const DIRECTION_ITEMS: powerbi.IEnumMember[] = [{ value: "vertical", displayName: "縦" }, { value: "horizontal", displayName: "左右" }];
/** 枠の「すべての行と同じ」（セグメントの枠は「すべてのセグメントと同じ」）。表示名は枠の種類で変える（apply） */
const INHERIT = "inherit";
const inheritItem = (kind: HierarchyKind): powerbi.IEnumMember => ({ value: INHERIT, displayName: kind === "org" ? "すべてのセグメントと同じ" : "すべての行と同じ" });
const none = { value: "", displayName: "指定なし" };
export const dropdown = (name: string, displayName: string, items: powerbi.IEnumMember[], value = String(items[0].value), description?: string) =>
    new f.ItemDropdown({ name, displayName, description, items, value: items.find(i => i.value === value) ?? items[0] });
export const ROOT_ITEMS: powerbi.IEnumMember[] = [{ value: "joined", displayName: "連続した表" }, ...DIRECTION_ITEMS.map(i => ({ ...i, displayName: `${i.displayName}分割` }))];

/** 科目・セグメントの表示の既定（「科目の行」「セグメントの行」のカードの項目）。段・行ごとの上書きは HierarchyOverrides */
export interface HierarchyDefaults {
    accountStyle: f.ItemDropdown;
    accountTotal: f.ItemDropdown;
    accountDirection: f.ItemDropdown;
    accountRoot: f.ItemDropdown;
    orgStyle: f.ItemDropdown;
    orgTotal: f.ItemDropdown;
    orgDirection: f.ItemDropdown;
}

/** 枠を見せるカード：行（account）かセグメント（org） */
export type HierarchyKind = "account" | "org";

/** 対象を選んでいない枠を、どちらのカードに見せるか（前の半分は行、後ろの半分はセグメント） */
export const HIERARCHY_ACCOUNT_SLOTS = HIERARCHY_SLOTS / 2;

/** 枠の種類：対象を選んでいればその種類、選んでいなければ番号で分ける */
export function hierarchyKindOf(target: string, slot: number): HierarchyKind {
    if (target.startsWith("account:")) return "account";
    if (target.startsWith("org:")) return "org";
    return slot <= HIERARCHY_ACCOUNT_SLOTS ? "account" : "org";
}

class HierarchyItem extends f.SimpleCard {
    kind: HierarchyKind = "account";
    target: f.ItemDropdown;
    style: f.ItemDropdown;
    total: f.ItemDropdown;
    direction: f.ItemDropdown;
    constructor(slot: number) {
        super();
        this.name = `hierarchyItem${slot}`;
        this.displayName = `設定 ${slot}`;
        this.target = dropdown(`target${slot}`, "対象", [none]);
        const inherit = inheritItem(this.kind);
        this.style = dropdown(`style${slot}`, "レイアウト", [inherit, ...STYLE_ITEMS], undefined, "選んだレベル・行の子の見せ方。すべての行と同じなら、上の「並べ方」の設定のまま");
        this.total = dropdown(`total${slot}`, "小計の位置", [inherit, ...TOTAL_ITEMS], undefined, "選んだレベル・行の小計を、子の下・上に置くか置かないか");
        this.direction = dropdown(`direction${slot}`, "分割の向き", [inherit, ...DIRECTION_ITEMS], undefined, "レイアウトが「分割」のとき、子の表を縦に並べるか左右に並べるか");
        this.slices = [this.target, this.style, this.total, this.direction];
    }
}
export class HierarchyOverrides extends f.CompositeCard {
    name = "hierarchyOverrides";
    displayName = "レベル・行ごとの配置";
    description = "「行」「セグメント」のレイアウト・小計の位置・分割の向きを、レベルか行ごとに上書きする。同じ対象を複数指定した場合は後の設定を優先。分割の向きは対象の子の並び方。書式ペインでは行・セグメントのカードの「配置」に入れて見せる（arrangeCards）";
    items = Array.from({ length: HIERARCHY_SLOTS }, (_, i) => new HierarchyItem(i + 1));
    groups = [new f.Group({ name: "entries", displayName: "個別設定", slices: [],
        container: new f.Container({ displayName: "編集対象", containerItems: this.items }) })];
    apply(data?: HierarchyFormat): void {
        this.items.forEach((item, i) => {
            const saved = data?.saved[i];
            const target = saved?.target ?? "";
            // 行とセグメントのカードに分けて見せる（arrangeCards）。対象の選択肢はその種類だけにし、頭の「行：」「セグメント：」は外す
            item.kind = hierarchyKindOf(target, i + 1);
            const own = (data?.targets ?? []).filter((t) => String(t.value).startsWith(`${item.kind}:`));
            item.target.items = [none, ...own.map((t) => ({ ...t, displayName: String(t.displayName).replace(/^(行|セグメント)：/, "") }))];
            // 「すべての行と同じ」「すべてのセグメントと同じ」は枠の種類で名前を変える。説明もセグメントの枠はセグメントの言い方にする
            const inherit = inheritItem(item.kind);
            for (const select of [item.style, item.total, item.direction]) select.items = [inherit, ...select.items.filter((i) => i.value !== INHERIT)];
            item.style.description = item.kind === "org" ? "選んだレベル・セグメントの子の見せ方。すべてのセグメントと同じなら、上の「並べ方」の設定のまま" : "選んだレベル・行の子の見せ方。すべての行と同じなら、上の「並べ方」の設定のまま";
            item.total.description = item.kind === "org" ? "選んだレベル・セグメントの小計を、子の下・上に置くか置かないか" : "選んだレベル・行の小計を、子の下・上に置くか置かないか";
            if (target && !item.target.items.some(v => v.value === target)) item.target.items.push({ value: target, displayName: `${target}（表に無い）` });
            for (const prop of ["target", "style", "total", "direction"] as const) {
                const select = item[prop];
                select.value = select.items.find(v => v.value === saved?.[prop]) ?? select.items[0];
            }
        });
    }
}

export function readHierarchyOverrides(objects: powerbi.DataViewObjects | undefined): HierarchyOverride[] {
    const object = objects?.hierarchyOverrides;
    return Array.from({ length: HIERARCHY_SLOTS }, (_, i) => {
        const read = (prop: string) => typeof object?.[`${prop}${i + 1}`] === "string" ? String(object[`${prop}${i + 1}`]) : "";
        return { target: read("target"), style: read("style"), total: read("total"), direction: read("direction") };
    });
}

export function hierarchyOptions(settings: HierarchyDefaults, overrides: HierarchyOverride[], kind: "account" | "org", id: string, level: number): HierarchyOptions {
    const read = (select: f.ItemDropdown, items: powerbi.IEnumMember[], fallback: string) => {
        const value = String(select.value.value);
        return items.some(i => i.value === value) ? value : fallback;
    };
    const result: HierarchyOptions = {
        style: read(settings[`${kind}Style`], STYLE_ITEMS, kind === "account" ? "box" : "split") as HierarchyStyle,
        total: read(settings[`${kind}Total`], TOTAL_ITEMS, "bottom") as TotalPosition,
        direction: read(settings[`${kind}Direction`], DIRECTION_ITEMS, "vertical") as SplitDirection,
    };
    for (const target of [`${kind}:level:${level}`, `${kind}:node:${id}`]) {
        for (const entry of overrides.filter(o => o.target === target)) {
            if (STYLE_ITEMS.some(i => i.value === entry.style)) result.style = entry.style as HierarchyStyle;
            if (TOTAL_ITEMS.some(i => i.value === entry.total)) result.total = entry.total as TotalPosition;
            if (DIRECTION_ITEMS.some(i => i.value === entry.direction)) result.direction = entry.direction as SplitDirection;
        }
    }
    return result;
}
