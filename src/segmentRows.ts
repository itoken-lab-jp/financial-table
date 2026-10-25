import powerbi from "powerbi-visuals-api";
import { formattingSettings as f } from "powerbi-visuals-utils-formattingmodel";

/**
 * セグメントごとに隠す行（書式ペインの「セグメントごとに隠す行」カード）。1 つの枠でセグメントと行（科目・指標・計算行）を選ぶと、
 * そのセグメントのブロックでその行と下の行を出さない。見た目だけで、親の小計・計算行の値は変えない（事業の合計と全社の合計が食い違わない）。
 * 下の行を持つ行（区分・中分類の小計、後ろに指標を置いた行）は「隠し方」で、その行だけを隠して下の行を残すこともできる（既定は下の行も隠す。
 * 2026-10-20 ユーザー「小計の行だけ隠すのか、その区分丸ごと隠すのか選べるように」）。
 * セグメントには全社（合計）のブロックも選べる。保存はセグメントの道筋（ROOT は合計）と行のコードと隠し方
 */
export const SEGMENT_ROW_SLOTS = 20;

/** 隠し方：all は選んだ行と下の行（区分ごと）、row は選んだ行だけ（その行のうちも一緒に隠す。親の無いうちが残らないように） */
export const HIDE_MODES = { all: "all", row: "row" } as const;
export type HideMode = (typeof HIDE_MODES)[keyof typeof HIDE_MODES];
const MODE_ITEMS: powerbi.IEnumMember[] = [
    { value: HIDE_MODES.all, displayName: "下の行も隠す（区分ごと）" },
    { value: HIDE_MODES.row, displayName: "この行だけ隠す" },
];

export interface SegmentRowSaved {
    org: string;
    row: string;
    mode: HideMode;
}

export interface SegmentRowsFormat {
    orgs: powerbi.IEnumMember[];
    rows: powerbi.IEnumMember[];
    saved: SegmentRowSaved[];
    /** 下の行を持つ行のコード（「隠し方」を出す行） */
    parents?: string[];
}

const NONE: powerbi.IEnumMember = { value: "", displayName: "指定なし" };

class SegmentRowItem extends f.SimpleCard {
    org: f.ItemDropdown;
    row: f.ItemDropdown;
    mode: f.ItemDropdown;
    constructor(slot: number) {
        super();
        this.name = `segmentRow${slot}`;
        // 表示名は固定（名前を変えると Desktop で編集する枠の選択が 1 本目に戻った。計算行と同じ）
        this.displayName = `設定 ${slot}`;
        this.org = new f.ItemDropdown({ name: `org${slot}`, displayName: "セグメント", items: [NONE], value: NONE });
        this.row = new f.ItemDropdown({ name: `row${slot}`, displayName: "隠す行", description: "選んだ行と、その下の行を出さない（隠し方で、その行だけにもできる）", items: [NONE], value: NONE });
        this.mode = new f.ItemDropdown({
            name: `mode${slot}`,
            displayName: "隠し方",
            description: "下の行も隠すか（区分ごと）、この行だけを隠して下の行は残すか",
            items: MODE_ITEMS,
            value: MODE_ITEMS[0],
        });
        this.slices = [this.org, this.row, this.mode];
    }
}

export class SegmentRowsCard extends f.CompositeCard {
    name = "segmentRows";
    displayName = "セグメントごとに隠す行";
    description = "セグメントを選び、そのセグメントで出さない行（科目・指標・計算行）を選ぶ。見た目だけで、小計・計算行の値は変えない";
    visible = false;
    items = Array.from({ length: SEGMENT_ROW_SLOTS }, (_, i) => new SegmentRowItem(i + 1));
    groups = [
        new f.Group({
            name: "segmentRowsGroup",
            displayName: "隠す行",
            slices: [],
            container: new f.Container({ displayName: "編集対象", containerItems: this.items }),
        }),
    ];

    /** 選択肢と保存値を流し込む。セグメントの欄が無ければ出さない。選択肢に無い保存値は「（表に無い）」で残す */
    apply(data?: SegmentRowsFormat): void {
        this.visible = !!data && data.orgs.length > 0;
        this.items.forEach((item, i) => {
            const saved = data?.saved[i];
            const fill = (select: f.ItemDropdown, choices: powerbi.IEnumMember[], value: string) => {
                select.items = [NONE, ...choices];
                if (value && !select.items.some((c) => c.value === value)) select.items.push({ value, displayName: `${value}（表に無い）` });
                select.value = select.items.find((c) => c.value === value) ?? NONE;
            };
            fill(item.org, data?.orgs ?? [], saved?.org ?? "");
            fill(item.row, data?.rows ?? [], saved?.row ?? "");
            item.mode.value = MODE_ITEMS.find((c) => c.value === saved?.mode) ?? MODE_ITEMS[0];
            // 下の行を持つ行を選んだ枠だけ（科目・下に何も無い計算行では、どちらも同じ）
            item.mode.visible = !!saved?.row && (data?.parents ?? []).includes(saved.row);
        });
    }
}

export function readSegmentRows(objects: powerbi.DataViewObjects | undefined): SegmentRowSaved[] {
    const object = objects?.segmentRows;
    const read = (name: string) => (typeof object?.[name] === "string" ? String(object[name]) : "");
    return Array.from({ length: SEGMENT_ROW_SLOTS }, (_, i) => ({
        org: read(`org${i + 1}`),
        row: read(`row${i + 1}`),
        mode: read(`mode${i + 1}`) === HIDE_MODES.row ? HIDE_MODES.row : HIDE_MODES.all,
    }));
}

/**
 * セグメントのブロック（org はブロックの道筋、合計は rootKey）で隠す行のコード。隠し方が all なら選んだ行の下の行（子・うち・後ろに置いた行）も、
 * row なら選んだ行とそのうちだけを入れる。childrenOf は行のコード → 下の行のコード、isBreakdown はうちの行か
 */
export function hiddenRowsIn(saved: SegmentRowSaved[], org: string, childrenOf: (code: string) => string[], isBreakdown: (code: string) => boolean = () => false): Set<string> {
    const hidden = new Set<string>();
    const add = (code: string) => {
        if (hidden.has(code)) return;
        hidden.add(code);
        childrenOf(code).forEach(add);
    };
    for (const slot of saved) {
        if (slot.org !== org || !slot.row) continue;
        if (slot.mode === HIDE_MODES.row) {
            hidden.add(slot.row);
            childrenOf(slot.row).filter(isBreakdown).forEach(add);
        } else add(slot.row);
    }
    return hidden;
}
