/**
 * 見る人が表の上のメニューで選んだ主と比較の列（列 1〜3 のシナリオと見せ方。期間の見出しの ▾ で選んだ、その期間だけの列も）を、レポートに保存する。
 *
 * ページを移るとビジュアルは作り直されるので、React の state だけでは戻ったときに消える。
 * `host.persistProperties` で書き戻し、`dataView.metadata.objects` から読み直す。
 *
 * 主は、選んだときの書式ペインの値（作り手が決めた主。保存値のまま、決めていなければ空）も一緒に持つ。作り手が書式ペインで決め直したら
 * これとずれるので、そのとき見る人の選択は捨てる。これが無いと、一度でも選ばれたレポートは書式ペインから変えられなくなる。データで決まる
 * 既定（最新見込み・比較順の一番小さいイベント）は持たない。データが変わっただけで捨てないため。
 * 比較の列は書式ペインに持たない（見る人が出し入れする）。作り手が Desktop で選んで保存した形が、開いたときの形になる
 *
 * 開き閉じも見る人の選択として持つ。組織は開いた所だけ（既定は根だけ開く。閉じた所を持つと「保存が無い = 全部開く」になり、
 * 組織の多いデータで最初の表が長くなる）、行は既定から切り替えた所だけ（組織のブロックがあれば既定は閉じるので開いた行、無ければ既定は
 * 開くので閉じた行。viewModel）。2 つの並びは分けて持つ（1 つにすると、作り手が組織の欄を出し入れしたときに同じ保存値の向きが逆に
 * 読まれた）。書式ペインの値とは関わらないので、作り手が決め直しても捨てない
 *
 * 見せる科目の選択（区分・中分類ごとに選んだ科目と見せ方）も持つ。作り手の既定は書式ペインに持たない（科目の数だけ並べられない）ので、
 * 作り手が Desktop で選んで保存したものが、開いたときの形になる
 *
 * 見る人が選んだ期間の列（年度をまたげる）と合計の列も持つ。選んだときの書式ペインの期間の設定と組にし、作り手が変えたら効かせない（viewModel）
 */
import powerbi from "powerbi-visuals-api";

import { PeriodPick } from "./periods";
import { PickMode, RowPick } from "./picks";

import DataView = powerbi.DataView;

/** metadata.objects に保存する内部状態のオブジェクト名。書式ペインには出さない */
export const VISUAL_STATE_OBJECT = "visualState";

/** 比較の列の数の上限 */
export const COMPARE_SLOTS = 3;

/** 比較の列 1 つ：シナリオ（null は「なし」で、その列は出ない）と見せ方。シナリオを「なし」にしても見せ方は残す（選び直したとき同じ見せ方） */
export interface CompareSlot {
    compare: string | null;
    view: string;
}

/**
 * 比較の列の選択。列 1〜3 の順が表の並び。
 * 同じシナリオを 2 つの列に選んでよい（差と比を別の列にするなど）
 */
export interface CompareChoice {
    /** 全期間の列 1〜3（いつも COMPARE_SLOTS 個）。null はまだ選んでいない（既定の 1 列。既定のシナリオは主で変わるので決め打ちで持たない） */
    all: CompareSlot[] | null;
    /** この期間だけの列 1〜3（期間の名前 → 列）。無い期間は全期間と同じ。期間の名前は年月つき（periodName。年度を替えると別の期間） */
    periods: Record<string, CompareSlot[]>;
}

export interface VisualState {
    /** 見る人が選んだ主。null = まだ選んでいない */
    main: string | null;
    /** 選んだときに作り手が書式ペインで決めていた主（決めていなければ null） */
    mainBase: string | null;
    /** 比較の列の選択。null はまだ選んでいない（既定の 1 本）。全部外したのは all が空（null と分ける） */
    compares: CompareChoice | null;
    /** 開いた組織の道筋（根はいつも開く） */
    openOrgs: string[];
    /** 開いた行のコード（区分・中分類の小計の行・見出し）。行を既定で閉じる表（組織のブロックがある）で使う */
    openRows: string[];
    /** 閉じた行のコード。行を既定で開く表（組織の欄が無い）で使う */
    closedRows: string[];
    /** 見せる科目の選択（選んだ区分・中分類だけ）。無い親はすべての科目を出す */
    picks: RowPick[];
    /** 見る人が選んだ期間の列。null は書式ペインのまま */
    periods: PeriodPick | null;
}

export interface PersistedVisualState {
    /** 保存直後の古い update を見分けるためのキー */
    raw: string;
    state: VisualState;
}

export const EMPTY_VISUAL_STATE: VisualState = { main: null, mainBase: null, compares: null, openOrgs: [], openRows: [], closedRows: [], picks: [], periods: null };

const KEYS = ["main", "mainBase"] as const;
const LIST_KEYS = ["openOrgs", "openRows", "closedRows"] as const;

export function serializeVisualState(state: VisualState): string {
    return JSON.stringify([...KEYS.map((k) => state[k]), state.compares, ...LIST_KEYS.map((k) => state[k]), state.picks, state.periods]);
}

/** persistProperties に渡す text のプロパティ。null は空文字、並びは JSON（空なら空文字） */
export function toPersistedProperties(state: VisualState): Record<string, string> {
    return {
        ...Object.fromEntries(KEYS.map((k) => [k, state[k] ?? ""])),
        compares: state.compares ? JSON.stringify(state.compares) : "",
        ...Object.fromEntries(LIST_KEYS.map((k) => [k, state[k].length > 0 ? JSON.stringify(state[k]) : ""])),
        picks: state.picks.length > 0 ? JSON.stringify(state.picks) : "",
        periods: state.periods ? JSON.stringify(state.periods) : "",
    };
}

/** 開き閉じの 1 つを切り替える（並びに入っていれば抜き、無ければ足す） */
export function toggled(list: string[], value: string): string[] {
    return list.includes(value) ? list.filter((v) => v !== value) : [...list, value];
}

/**
 * 保存値の比較の列 1〜3 を読む。並びでなければ null。 の形（比較対象の名前の並び）と空の並びも null（今の形はいつも 3 列を書く）。
 * 形の合わない列の中身は「なし」、足りない列は「なし」で埋め、多い列は捨てる。シナリオの空の文字は「なし」
 */
function readSlots(value: unknown): CompareSlot[] | null {
    if (!Array.isArray(value) || value.length === 0 || value.some((v) => typeof v === "string")) return null;
    return Array.from({ length: COMPARE_SLOTS }, (_, i): CompareSlot => {
        const v: unknown = value[i];
        const o = v !== null && typeof v === "object" && !Array.isArray(v) ? (v as Record<string, unknown>) : undefined;
        const compare = o && typeof o.compare === "string" && o.compare !== "" ? o.compare : null;
        return { compare, view: o && typeof o.view === "string" ? o.view : "" };
    });
}

/** 保存値の比較の列の選択を読む。無い・形の合わないものは null（既定の 1 列）。 の形（比較対象の名前の並び）は読み替えない（公開前） */
function readCompares(value: unknown): CompareChoice | null {
    if (typeof value !== "string" || value === "") return null;
    try {
        const o = JSON.parse(value) as Record<string, unknown>;
        if (!o || typeof o !== "object" || !(o.all === null || Array.isArray(o.all))) return null;
        const all = o.all === null ? null : readSlots(o.all);
        // 前の形は丸ごと読まない（全部外した空の並びだけ「なし」として残ることのないように）
        if (o.all !== null && all === null) return null;
        const periods: Record<string, CompareSlot[]> = {};
        if (o.periods && typeof o.periods === "object") {
            for (const [k, v] of Object.entries(o.periods as Record<string, unknown>)) {
                // 手で書いた __proto__ の期間で、入れ物の prototype を差し替えない
                if (k === "__proto__") continue;
                const slots = readSlots(v);
                if (slots) periods[k] = slots;
            }
        }
        return { all, periods };
    } catch {
        return null;
    }
}

/** 保存値の期間の選択を読む。無い・形の合わないものは null（書式ペインのまま） */
function readPeriods(value: unknown): PeriodPick | null {
    if (typeof value !== "string" || value === "") return null;
    try {
        const o = JSON.parse(value) as Record<string, unknown>;
        if (!o || typeof o !== "object" || typeof o.base !== "string" || !Array.isArray(o.columns)) return null;
        return { base: o.base, columns: o.columns.filter((c): c is string => typeof c === "string"), total: o.total === true };
    } catch {
        return null;
    }
}

/** 保存値の見せる科目の選択を読む。形の合わないものは捨てる。同じ親が 2 度あれば最初のもの */
function readPicks(value: unknown): RowPick[] {
    if (!Array.isArray(value)) return [];
    const seen = new Set<string>();
    return value.flatMap((item): RowPick[] => {
        if (!item || typeof item !== "object") return [];
        const o = item as Record<string, unknown>;
        if (typeof o.parent !== "string" || (o.mode !== "others" && o.mode !== "under") || !Array.isArray(o.codes) || seen.has(o.parent)) return [];
        seen.add(o.parent);
        return [{ parent: o.parent, mode: o.mode, codes: o.codes.filter((c): c is string => typeof c === "string") }];
    });
}

/** 保存済みの状態を読む。無い・主が文字でなければ null。壊れた並び・比較の列は空にして読む（主の選択は捨てない） */
export function readVisualState(dataView: DataView | undefined): PersistedVisualState | null {
    const raw = dataView?.metadata?.objects?.[VISUAL_STATE_OBJECT];
    if (!raw || typeof raw !== "object") return null;
    const object = raw as unknown as Record<string, unknown>;
    const state: VisualState = { ...EMPTY_VISUAL_STATE, compares: null, openOrgs: [], openRows: [], closedRows: [], picks: [], periods: null };
    for (const key of KEYS) {
        const value = object[key];
        if (value === undefined || value === "") continue;
        if (typeof value !== "string") return null;
        state[key] = value;
    }
    // 壊れた並びは空にする（主の選択は捨てない）
    const parse = (value: unknown): unknown => {
        if (typeof value !== "string" || value === "") return [];
        try {
            return JSON.parse(value);
        } catch {
            return [];
        }
    };
    state.compares = readCompares(object.compares);
    state.picks = readPicks(parse(object.picks));
    state.periods = readPeriods(object.periods);
    for (const key of LIST_KEYS) {
        const list = parse(object[key]);
        state[key] = Array.isArray(list) ? list.filter((v): v is string => typeof v === "string") : [];
    }
    return { raw: serializeVisualState(state), state };
}

/** 作り手が書式ペインで決めた値が選んだときと変わっていれば、見る人の選択を捨てる。変わらなければ同じオブジェクトを返す */
export function withoutStale(state: VisualState, authorMain: string | null, periodBase?: string): VisualState {
    const staleMain = state.main !== null && state.mainBase !== authorMain;
    // 期間の選択も、作り手が期間の設定を変えたら捨てる（効かせないだけにすると、作り手が元に戻したときに古い選択が生き返った）
    const stalePeriods = periodBase !== undefined && state.periods !== null && state.periods.base !== periodBase;
    if (!staleMain && !stalePeriods) return state;
    return {
        ...state,
        main: staleMain ? null : state.main,
        mainBase: staleMain ? null : state.mainBase,
        periods: stalePeriods ? null : state.periods,
    };
}

/*
 * 比較の列の選択の決まり
 *
 * - 保存するのは見る人が選んだ列 1〜3 そのもの（シナリオと見せ方）。表には、シナリオが今選べる（candidates。主と同じもの・表の年度に
 *   データの無いものは入らない）列だけを、列 1〜3 の順に出す（shownSlots）。今は使えないシナリオの列は、選択を消さずに出さない
 *   （主や年度を戻せば出る。既定に差し替えない）。ダイアログではその列に「今は使えない」と出す
 * - 選ぶ操作は、押した列の箱だけを書き換える。ほかの列（今は使えない列も）は触らない
 * - まだ選んでいない（all が null）ときは既定の 1 列（defaultSlots。シナリオは主で変わる既定の比較対象）。最初に選んだとき、その既定を書き出して
 *   から押した所を変える
 * - 期間の見出しの ▾ は「全期間と同じ」のチェックで、その期間だけの列を持つかを見る人が決める。外すと今の全期間の列を写し、入れるとその期間の
 *   列を捨てる。その期間だけの列が全期間と同じになっても勝手に全期間に戻さない（見る人のチェックのまま）
 */

/** まだ選んでいないときの列 1〜3：既定のシナリオ（fallback）を 1 列目に、見せ方は既定 */
export function defaultSlots(fallback: string, view: string): CompareSlot[] {
    return Array.from({ length: COMPARE_SLOTS }, (_, i): CompareSlot => ({ compare: i === 0 ? fallback : null, view }));
}

/** 全期間（period が null）か期間の列 1〜3 と、その期間だけの列を持っているか。期間の列が無ければ全期間の列 */
export function slotsOf(choice: CompareChoice | null, period: string | null, fallback: string, view: string): { slots: CompareSlot[]; own: boolean } {
    const own = period !== null ? choice?.periods[period] : undefined;
    if (own) return { slots: own, own: true };
    return { slots: choice?.all ?? defaultSlots(fallback, view), own: false };
}

/** 表に出す列（列の番号つき、列 1〜3 の順）：シナリオが今選べる列だけ */
export function shownSlots(slots: CompareSlot[], candidates: string[]): Array<{ compare: string; view: string; index: number }> {
    return slots.flatMap((slot, index) => (slot.compare !== null && candidates.includes(slot.compare) ? [{ compare: slot.compare, view: slot.view, index }] : []));
}

/** 見る人が比較の列を選んだこと（period が null なら表の上のバーの「比較」で全期間、名前なら期間の見出しの ▾ でその期間） */
export type ComparePatch =
    /** 列 index のシナリオ（null は「なし」）か見せ方を変える */
    | { kind: "slot"; period: string | null; index: number; compare?: string | null; view?: string }
    /** この期間だけの列を持つか（「全期間と同じ」のチェックを外すと true） */
    | { kind: "own"; period: string; own: boolean };

/** 比較の列の選択を書く（上の決まり）。fallback と view は表に出したときと同じ既定（EventMenu） */
export function chooseCompare(state: VisualState, fallback: string, view: string, patch: ComparePatch): VisualState {
    const base: CompareChoice = state.compares ?? { all: null, periods: {} };
    const all = base.all ?? defaultSlots(fallback, view);
    let next: CompareChoice;
    switch (patch.kind) {
        case "slot": {
            if (patch.index < 0 || patch.index >= COMPARE_SLOTS) return state;
            const edit = (slots: CompareSlot[]) =>
                slots.map((slot, i): CompareSlot =>
                    i === patch.index ? { compare: patch.compare !== undefined ? patch.compare : slot.compare, view: patch.view ?? slot.view } : slot
                );
            if (patch.period === null) {
                next = { ...base, all: edit(all) };
                break;
            }
            // その期間だけの列を持っていなければ（ダイアログでは押せないが）、全期間の列を写してから変える
            next = { ...base, periods: { ...base.periods, [patch.period]: edit(base.periods[patch.period] ?? all) } };
            break;
        }
        case "own": {
            const periods = { ...base.periods };
            if (patch.own) periods[patch.period] = periods[patch.period] ?? all.map((slot) => ({ ...slot }));
            else delete periods[patch.period];
            next = { ...base, periods };
            break;
        }
    }
    return { ...state, compares: next };
}

/** 見る人が見せる科目のダイアログで選んだこと */
export type PickPatch =
    | { kind: "toggle"; code: string }
    | { kind: "mode"; mode: PickMode }
    | { kind: "all" }
    | { kind: "none" }
    | { kind: "reset" };

/**
 * 見せる科目の選択を書く。まだ選んでいない親は、すべての科目を選んだ「その他」から始める。すべてを選んでも、選択は残す
 * （あとから出てきた科目は選ばなかった側に入れるため）。「もとに戻す」はその親の選択を捨てる（すべての科目を出す）。
 * accounts はその親の選べる科目（表の並び）
 */
export function choosePick(state: VisualState, parent: string, accounts: string[], patch: PickPatch): VisualState {
    const current = state.picks.find((p) => p.parent === parent) ?? { parent, mode: "others" as PickMode, codes: accounts };
    let next: RowPick | null;
    switch (patch.kind) {
        case "toggle": {
            const codes = current.codes.includes(patch.code) ? current.codes.filter((c) => c !== patch.code) : [...current.codes, patch.code];
            // 表の並びにそろえる（今は無い科目の選択は後ろに残す）
            next = { ...current, codes: [...accounts.filter((c) => codes.includes(c)), ...codes.filter((c) => !accounts.includes(c))] };
            break;
        }
        case "mode":
            next = { ...current, mode: patch.mode };
            break;
        // すべて選ぶ・すべて外すは、今は無い科目の選択も捨てる（フィルターを戻したときに、外したはずの科目が出ないように）
        case "all":
            next = { ...current, codes: accounts };
            break;
        case "none":
            next = { ...current, codes: [] };
            break;
        case "reset":
            next = null;
            break;
    }
    const picks = state.picks.filter((p) => p.parent !== parent);
    return { ...state, picks: next ? [...picks, next] : picks };
}

/** 見る人が期間のダイアログで選んだこと */
export type PeriodPatch = { kind: "toggle"; id: string } | { kind: "total" } | { kind: "set"; ids: string[]; on: boolean } | { kind: "reset" };

/**
 * 期間の選択を書く。まだ選んでいない（か、作り手が期間の設定を変えた）ときは、今出している列から始める。「もとに戻す」は選択を捨てる
 * （書式ペインのまま）。base は今の書式ペインの期間の設定、current は今出している列の名前と合計の列、表が見る人の選択で組めているか
 */
export function choosePeriods(
    state: VisualState,
    base: string,
    current: { columns: string[]; total: boolean; picked: boolean },
    patch: PeriodPatch
): VisualState {
    if (patch.kind === "reset") return { ...state, periods: null };
    // 表が見る人の選択で組めていないとき（選んだ列が 1 つも無く、書式ペインのままに戻した）は、保存値ではなく今出している列から始める
    // （ダイアログは今の列を示すので、保存値から始めるとチェックが逆に効いた）
    const from: PeriodPick =
        current.picked && state.periods && state.periods.base === base ? state.periods : { base, columns: current.columns, total: current.total };
    let next: PeriodPick;
    switch (patch.kind) {
        case "toggle":
            next = { ...from, columns: from.columns.includes(patch.id) ? from.columns.filter((c) => c !== patch.id) : [...from.columns, patch.id] };
            break;
        case "total":
            next = { ...from, total: !from.total };
            break;
        case "set":
            next = {
                ...from,
                columns: patch.on ? Array.from(new Set([...from.columns, ...patch.ids])) : from.columns.filter((c) => !patch.ids.includes(c)),
            };
            break;
    }
    return { ...state, periods: next };
}
