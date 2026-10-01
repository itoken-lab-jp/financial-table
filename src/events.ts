/**
 * イベントの比較順。
 *
 * - イベントの性質は比較順 1 つ。数か日付（ファイルの更新日など）で、大きい・新しいほど確か。種類（実績・見通し・計画）は持たない
 * - 最新見込み：組織 × 月ごとに、数字のあるイベントのうち比較順の一番大きいものを採る（compute.ts）
 * - 比較順が空のイベント（前年実績など）は、主に採らず比較にだけ使う
 * - 縦持ちは比較順の列、横持ちは書式ペインのメジャーごとの比較順（初期値は欄の並びで、後ろほど確か）。
 *   イベントの名前は読まない（決まった言葉で決めると汎用にならない）。縦持ちで比較順の列が無ければ、最新見込みは出さない
 * - イベントの列と金額のメジャー 2 本以上を一緒に入れると、イベントは「実績・数量」のようにメジャーと組になる（系列）。
 *   既定・比較の落とし先・比較順が同じかどうかは、同じ系列（メジャー）の中だけで見る
 */
import { EventInfo, InputData } from "./data";

/** 書式ペインのメジャーごとの比較順の「比較だけ（比較順なし）」 */
export const ORDER_NONE = "none";
/** 書式ペインのメジャーごとの金額の持ち方の「表全体に合わせる」 */
export const SIGN_TABLE = "table";
export const SIGN_VALUES = ["auto", "debit", "credit", "positive"];

/** 書式ペインに出す、横持ちのメジャーごとの設定 */
export interface MeasureSetting {
    name: string;
    queryName: string;
    /** 比較順（ORDER_NONE か "1"〜）。保存が無ければ欄の並び */
    order: string;
    /** 比較順の選択肢（"1"〜メジャーの数。それより大きい値が保存されていれば、その値も） */
    orderChoices: string[];
    /** 金額の持ち方（SIGN_TABLE か auto・debit・credit・positive）。保存が無ければ表全体に合わせる */
    sign: string;
}

export interface EventModel {
    /** 選ぶときの並び：比較順の小さい（弱い）順。比較順の無いイベントは後ろ */
    events: EventInfo[];
    /** 最新見込みで採る順（比較順の大きい順）。比較順の無いイベントは入れない */
    latest: string[];
    /** 主に選べるイベント（比較順のあるもの。1 つも無ければすべて） */
    mainEvents: string[];
    /** 最新見込みを出せないときの主の既定：最初の系列（金額の欄の最初のメジャー）で比較順の一番大きいイベント */
    strongest: string | null;
    warnings: string[];
    /** 横持ちのメジャーごとの設定（書式ペイン）。メジャーが 2 本以上の横持ちのときだけ */
    measureSettings: MeasureSetting[];
    /** メジャーごとに選んだ金額の持ち方（表全体に合わせるものは入れない） */
    measureSigns: Map<string, string>;
}

/** 名前の並び（比較順が同じ・無いときの並べ方だけに使う）。数字は数として比べる（第2回 < 第10回） */
const byName = (a: string, b: string) => a.localeCompare(b, "ja", { numeric: true });

/**
 * 選ぶときの並び：比較順の小さい順、比較順の無いイベントは後ろ。同じ比較順は、縦持ちは名前、横持ちは欄の並び（後ろほど強い）。
 * 届いた順（縦持ち）には意味を持たせない
 */
export function sortEvents(events: EventInfo[], source: InputData["eventSource"]): EventInfo[] {
    const tie = (a: EventInfo, b: EventInfo) => (source === "column" ? byName(a.name, b.name) : a.seen - b.seen);
    return [...events].sort((a, b) => {
        if (a.order === null || b.order === null) return a.order === b.order ? tie(a, b) : a.order === null ? 1 : -1;
        return a.order - b.order || tie(a, b);
    });
}

/**
 * 比較の個々のイベントの採り方：組織 × 月にそのイベントが無ければ、同じ系列で比較順が下のイベントへ落とす。
 * 比較順の無いイベントは落とさない
 */
export function fallbackOf(event: string, model: Pick<EventModel, "events" | "latest">): string[] {
    const at = model.latest.indexOf(event);
    if (at < 0) return [event];
    const seriesOf = (name: string) => model.events.find((e) => e.name === name)?.series?.name;
    const series = seriesOf(event);
    return model.latest.slice(at).filter((name) => seriesOf(name) === series);
}

/** 比較順を決める。縦持ちは比較順の列、横持ちは書式ペインの値（無ければ欄の並び） */
export function resolveEvents(input: InputData): EventModel {
    const warnings: string[] = [];
    const events = input.events.map((e) => ({ ...e }));
    const multiple = events.length >= 2;

    if (input.eventSource !== "column") {
        // 横持ち：書式ペインのメジャーごとの比較順。保存が無ければ欄の並び（1 から。後ろほど確か＝新しいイベントを後ろに入れる）
        events.forEach((e) => {
            const saved = e.measure?.savedOrder ?? null;
            const order = saved === ORDER_NONE ? null : saved !== null && /^\d+$/.test(saved) ? Number(saved) : e.seen + 1;
            e.order = order;
            e.orderKind = order === null ? null : "number";
        });
    }

    const sorted = sortEvents(events, input.eventSource);
    const ordered = sorted.filter((e) => e.order !== null);
    const latest = [...ordered].reverse().map((e) => e.name);
    // 最初の系列（系列が無ければすべて）で一番確かなイベント
    const firstSeries = Math.min(...ordered.map((e) => e.series?.index ?? 0));
    const strongest = [...ordered].reverse().find((e) => (e.series?.index ?? 0) === firstSeries)?.name ?? null;

    for (const name of input.orderConflicts.slice(0, 3)) warnings.push(`シナリオ「${name}」の比較順が行によって違う。最初に読めた値を使った`);
    if (input.unreadableOrders.length > 0) {
        warnings.push(`比較順が数か日付として読めないシナリオがある（${input.unreadableOrders.slice(0, 3).join("・")}）。比較にだけ使った。比較順は数か日付にする`);
    }
    if (new Set(ordered.map((e) => e.orderKind)).size > 1) warnings.push("比較順に数と日付が混ざっている。日付を数より大きい（確か）として並べた。どちらかにそろえる");
    // 同じ比較順は、同じ系列の中だけで見る（金額と数量の予算は、どちらもイベントの表の同じ比較順を持つ）
    const tied = ordered.filter((e) => ordered.some((o) => o !== e && o.order === e.order && o.series?.name === e.series?.name)).map((e) => e.name);
    if (tied.length > 0) {
        const by = input.eventSource === "column" ? "名前" : "欄の並び";
        warnings.push(`比較順が同じシナリオがある（${tied.slice(0, 3).join("・")}）。${by}の後ろを確かとした。比較順を別の値にする`);
    }
    if (input.eventSource === "column" && input.has.eventOrder && multiple && ordered.length === 0) {
        warnings.push("比較順のあるシナリオが無い（どれも空）。最新見込みは出さない。主に採るシナリオに比較順（数か日付。大きいほど確か）を入れる");
    }

    const count = events.length;
    const signOf = (e: EventInfo) => {
        const saved = e.measure?.savedSign ?? null;
        return saved !== null && SIGN_VALUES.includes(saved) ? saved : SIGN_TABLE;
    };
    const measureSettings: MeasureSetting[] =
        input.eventSource === "measures" && multiple
            ? events.map((e) => {
                  const orderChoices = Array.from({ length: count }, (_, i) => String(i + 1));
                  if (e.order !== null && e.order > count) orderChoices.push(String(e.order));
                  return {
                      name: e.name,
                      queryName: e.measure?.queryName ?? e.name,
                      order: e.order === null ? ORDER_NONE : String(e.order),
                      orderChoices,
                      sign: signOf(e),
                  };
              })
            : [];
    // 書式ペインに出しているとき（メジャーが 2 本以上の横持ち）だけ効かせる。1 本に減らしたあとに残った保存値は、見えないので使わない
    const measureSigns = new Map(measureSettings.filter((m) => m.sign !== SIGN_TABLE).map((m) => [m.name, m.sign]));

    return {
        events: sorted,
        latest,
        mainEvents: (ordered.length > 0 ? ordered : sorted).map((e) => e.name),
        strongest,
        warnings,
        measureSettings,
        measureSigns,
    };
}
