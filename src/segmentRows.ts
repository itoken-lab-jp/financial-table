import powerbi from "powerbi-visuals-api";
import { formattingSettings as f } from "powerbi-visuals-utils-formattingmodel";

/**
 * セグメントごとに隠す行（書式ペインの「セグメントごとに隠す行」カード）。1 つの枠でセグメントと行（科目・指標・計算行）を選ぶと、
 * そのセグメントのブロックでその行と下の行を出さない。見た目だけで、親の小計・計算行の値は変えない（事業の合計と全社の合計が食い違わない）。
 * セグメントには全社（合計）のブロックも選べる。保存はセグメントの道筋（ROOT は合計）と行のコード
 */
export const SEGMENT_ROW_SLOTS = 20;

export interface SegmentRowSaved {
    org: string;
    row: string;
}

export interface SegmentRowsFormat {
    orgs: powerbi.IEnumMember[];
    rows: powerbi.IEnumMember[];
    saved: SegmentRowSaved[];
}

const NONE: powerbi.IEnumMember = { value: "", displayName: "指定なし" };

class SegmentRowItem extends f.SimpleCard {
    org: f.ItemDropdown;
    row: f.ItemDropdown;
    constructor(slot: number) {
        super();
        this.name = `segmentRow${slot}`;
        // 表示名は固定（名前を変えると Desktop で編集する枠の選択が 1 本目に戻った。計算行と同じ）
        this.displayName = `設定 ${slot}`;
        this.org = new f.ItemDropdown({ name: `org${slot}`, displayName: "セグメント", items: [NONE], value: NONE });
        this.row = new f.ItemDropdown({ name: `row${slot}`, displayName: "隠す行", description: "選んだ行と、その下の行を出さない", items: [NONE], value: NONE });
        this.slices = [this.org, this.row];
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
        });
    }
}

export function readSegmentRows(objects: powerbi.DataViewObjects | undefined): SegmentRowSaved[] {
    const object = objects?.segmentRows;
    const read = (name: string) => (typeof object?.[name] === "string" ? String(object[name]) : "");
    return Array.from({ length: SEGMENT_ROW_SLOTS }, (_, i) => ({ org: read(`org${i + 1}`), row: read(`row${i + 1}`) }));
}

/**
 * セグメントのブロック（org はブロックの道筋、合計は rootKey）で隠す行のコード。選んだ行の下の行（子・うち）も入れる。
 * childrenOf は行のコード → 下の行のコード
 */
export function hiddenRowsIn(saved: SegmentRowSaved[], org: string, childrenOf: (code: string) => string[]): Set<string> {
    const hidden = new Set<string>();
    const add = (code: string) => {
        if (hidden.has(code)) return;
        hidden.add(code);
        childrenOf(code).forEach(add);
    };
    for (const slot of saved) if (slot.org === org && slot.row) add(slot.row);
    return hidden;
}
