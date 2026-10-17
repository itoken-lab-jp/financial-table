import powerbi from "powerbi-visuals-api";

/**
 * 書式ペインの「適用先」（「数値」「行」「セグメント」のカードの一番上）。一覧には最初から全部の対象（段・行・セグメント）を並べ、選ぶと
 * 下の欄がその対象の値になる（2026-10-17 ユーザー「最初から対象を持っておけないか」「適用先の分け方を統一」）。
 *
 * 保存は今の枠のまま（rowNumbers・hierarchyOverrides の target n など）。まだ枠を持たない対象を選んだら、空いた枠にその対象を書き込み
 * （ビジュアルが persistProperties。Reservation）、下の欄をその枠につなぐ。空いた枠は、対象の無い枠か、何も変えていない枠
 * （全部「表全体と同じ」「既定に合わせる」に戻った枠。選んだだけの対象も 1 枠使うので使い回す）
 */
export const SCOPE_ALL = "all";

/** 枠 1 本の今の対象と、何か変えているか */
export interface ScopeSlot {
    target: string;
    used: boolean;
}

/** 空いた枠に書き込む対象（枠は 1 から） */
export interface Reservation {
    objectName: string;
    properties: Record<string, string>;
}

export interface ScopeChoice {
    items: powerbi.IEnumMember[];
    /** 選んだ適用先の値（SCOPE_ALL か対象） */
    value: string;
    /** 下の欄をつなぐ枠（1 から）。すべてなら null。枠が足りなければ null */
    slot: number | null;
    /** 対象を書き込む枠（1 から）。もう枠を持っていれば null */
    reserve: number | null;
}

/** 変えた対象の印（一覧で見分ける） */
export const USED_MARK = "（変更あり）";

/**
 * 適用先の一覧と、選んだ対象の枠を決める。選択肢に無い保存値の枠（表から消えた行）は、何か変えていれば「（表に無い）」で一覧の最後に残す
 * （選んで「表全体と同じ」に戻せば消せる）
 */
export function resolveScope(
    allName: string,
    choices: powerbi.IEnumMember[],
    slots: ScopeSlot[],
    saved: string,
    options: { own?: (target: string) => boolean; keep?: Set<string> } = {},
): ScopeChoice {
    // own：枠をほかの種類と分け合うとき（行とセグメントの配置）、この一覧の種類の対象か。keep：ほかの一覧が今選んでいる対象（その枠を使い回すと取り合いになる）
    const own = options.own ?? (() => true);
    const keep = options.keep ?? new Set<string>();
    const usedTargets = new Set(slots.filter((s) => s.used && s.target !== "" && own(s.target)).map((s) => s.target));
    const known = new Set(choices.map((c) => String(c.value)));
    const stale = Array.from(usedTargets).filter((t) => !known.has(t));
    const items: powerbi.IEnumMember[] = [
        { value: SCOPE_ALL, displayName: allName },
        ...choices.map((c) => (usedTargets.has(String(c.value)) ? { ...c, displayName: `${c.displayName}${USED_MARK}` } : c)),
        ...stale.map((t) => ({ value: t, displayName: `${t.replace(/^(account|org):(node|level):/, "")}（表に無い）` })),
    ];
    const value = items.some((i) => i.value === saved) ? saved : SCOPE_ALL;
    if (value === SCOPE_ALL) return { items, value, slot: null, reserve: null };
    const mine = slots.findIndex((s) => s.target === value);
    if (mine >= 0) return { items, value, slot: mine + 1, reserve: null };
    const empty = slots.findIndex((s) => s.target === "");
    const free = empty >= 0 ? empty : slots.findIndex((s) => !s.used && !keep.has(s.target));
    return free >= 0 ? { items, value, slot: free + 1, reserve: free + 1 } : { items, value, slot: null, reserve: null };
}
