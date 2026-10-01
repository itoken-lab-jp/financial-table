/**
 * 比較の列を選ぶダイアログ。表の上のバーの「比較」（全期間）と、期間の見出しの ▾（その期間）で同じ形。
 *
 * 列 1〜3 ごとに、シナリオ（「なし」でその列は出ない）と見せ方の 2 つの箱。期間のダイアログは上に「全期間と同じ」のチェックを置き、入っている
 * あいだは全期間の列を押せない形で見せる。外すと今の全期間の列を写して、その期間だけ変えられる。選ぶたびに表に効かせて保存する（visual.ts）。
 * 今は使えないシナリオ（主と同じ・表の年度にデータが無い）を選んだ列は「今は使えない」と出す（選択は消さない。visualState.ts の決まり）。
 * 閉じる・外を押す・Esc・Tab（Shift+Tab）でダイアログの外へ出る・ビジュアルの大きさが変わると閉じる。ビジュアルの枠の中に収める
 */
import * as React from "react";

import { COMPARE_VIEW_ITEMS } from "./settings";
import { SlotView } from "./viewModel";
import { ComparePatch } from "./visualState";

export interface CompareDialogProps {
    /** 見出し（「比較の列（全期間）」「比較の列（4月）」） */
    title: string;
    /** シナリオの選択肢（今選べるもの） */
    items: Array<{ value: string; label: string }>;
    /** 列 1〜3。期間のダイアログで「全期間と同じ」なら全期間の列 */
    slots: SlotView[];
    /** 期間の名前（期間の見出しの ▾）。表の上のバーなら null */
    period: string | null;
    /** 期間のダイアログで、その期間だけの列を持っていないか（「全期間と同じ」のチェック） */
    same?: boolean;
    /** 開いたときにフォーカスを置く列（比較の列の見出しから開いたとき）。null なら最初の箱 */
    focusSlot?: number | null;
    left: number;
    top: number;
    /** 開くボタン（外を押したときに閉じない。ボタンが自分で開き閉じする） */
    opener: string;
    onChoose: (patch: ComparePatch) => void;
    onClose: (returnFocus: boolean) => void;
}

export function CompareDialog({ title, items, slots, period, same, focusSlot, left, top, opener, onChoose, onClose }: CompareDialogProps): React.JSX.Element {
    const ref = React.useRef<HTMLDivElement>(null);
    // ビジュアルの枠の中に収める位置（描いてから大きさを測る）。はみ出す分は左と上へ寄せ、高さは枠に合わせてスクロールさせる
    const [place, setPlace] = React.useState<{ left: number; top: number; maxHeight?: number }>({ left, top });
    React.useLayoutEffect(() => {
        const pop = ref.current;
        const root = pop?.offsetParent as HTMLElement | null;
        if (!pop || !root) return;
        const x = Math.max(0, Math.min(left, root.clientWidth - pop.offsetWidth - 2));
        const below = root.clientHeight - top - 2;
        const y = pop.offsetHeight <= below ? top : Math.max(0, root.clientHeight - pop.offsetHeight - 2);
        setPlace({ left: x, top: y, maxHeight: Math.max(80, root.clientHeight - y - 2) });
    }, [left, top]);
    // 開いたら、比較の列の見出しから開いたならその列のシナリオの箱へ、ほかは最初の箱へ
    React.useEffect(() => {
        const own = focusSlot !== undefined && focusSlot !== null ? ref.current?.querySelector<HTMLElement>(`select[data-slot="${focusSlot}"]:not(:disabled)`) : null;
        (own ?? ref.current?.querySelector<HTMLElement>("input, select:not(:disabled)"))?.focus({ preventScroll: true });
    }, [focusSlot, period]);
    // 外を押す・Tab でダイアログの外へフォーカスが移ったら閉じる（開くボタンは自分で開き閉じする）
    React.useEffect(() => {
        const outside = (target: HTMLElement) => !ref.current?.contains(target) && !target.closest?.(opener);
        const down = (e: MouseEvent) => {
            if (outside(e.target as HTMLElement)) onClose(false);
        };
        const focus = (e: FocusEvent) => {
            if (outside(e.target as HTMLElement)) onClose(false);
        };
        document.addEventListener("mousedown", down, true);
        document.addEventListener("focusin", focus, true);
        return () => {
            document.removeEventListener("mousedown", down, true);
            document.removeEventListener("focusin", focus, true);
        };
    }, [onClose, opener]);
    const locked = period !== null && same === true;
    const unavailable = slots.some((slot) => slot.compare !== null && !slot.available);
    return (
        <div
            ref={ref}
            className="ft-pick-pop ft-compare-pop"
            role="dialog"
            aria-label={title}
            style={{ left: place.left, top: place.top, maxHeight: place.maxHeight }}
            onKeyDown={(e) => {
                if (e.key === "Escape") {
                    e.preventDefault();
                    onClose(true);
                }
                // 表のキーボード（矢印で動くなど）に渡さない
                e.stopPropagation();
            }}
            onClick={(e) => e.stopPropagation()}
            onContextMenu={(e) => {
                e.preventDefault();
                e.stopPropagation();
            }}
        >
            <div className="ft-pick-title">
                <span>{title}</span>
                <button type="button" className="ft-pick-close" onClick={() => onClose(true)}>
                    閉じる
                </button>
            </div>
            <div className="ft-pick-body">
                {period !== null && (
                    <label className="ft-pick-account ft-compare-same">
                        <input type="checkbox" checked={same === true} onChange={() => onChoose({ kind: "own", period, own: same === true })} />
                        全期間と同じ
                    </label>
                )}
                {slots.map((slot, index) => {
                    const name = `列${index + 1}`;
                    return (
                        <div key={index} className={`ft-compare-slot${locked ? " ft-compare-locked" : ""}`}>
                            <span className="ft-compare-slot-name">{name}</span>
                            <select
                                className="ft-select"
                                aria-label={`${name}のシナリオ`}
                                data-slot={index}
                                disabled={locked}
                                value={slot.compare ?? ""}
                                onChange={(e) => onChoose({ kind: "slot", period, index, compare: e.target.value === "" ? null : e.target.value })}
                            >
                                <option value="">（なし）</option>
                                {items.map((item) => (
                                    <option key={item.value} value={item.value}>
                                        {item.label}
                                    </option>
                                ))}
                                {slot.compare !== null && !slot.available && <option value={slot.compare}>{slot.label}（今は使えない）</option>}
                            </select>
                            <select
                                className="ft-select"
                                aria-label={`${name}の見せ方`}
                                disabled={locked}
                                // シナリオが「なし」の列は見せ方を隠す（場所は空けたまま。箱が左右にずれない）
                                style={slot.compare === null ? { visibility: "hidden" } : undefined}
                                value={slot.view}
                                onChange={(e) => onChoose({ kind: "slot", period, index, view: e.target.value })}
                            >
                                {COMPARE_VIEW_ITEMS.map((view) => (
                                    <option key={view.value} value={view.value}>
                                        {view.displayName}
                                    </option>
                                ))}
                            </select>
                        </div>
                    );
                })}
                {unavailable && <div className="ft-compare-hint">「今は使えない」は、主と同じか、表の年度にデータが無いシナリオ。主や年度を戻すと出る</div>}
                <div className="ft-compare-hint">
                    {period === null
                        ? "期間ごとに変えるときは、期間の見出しの ▾ で選ぶ"
                        : locked
                          ? "全期間の列（表の上の「比較」）に付いていく。外すとこの期間だけ変えられる"
                          : "この期間だけの列。「全期間と同じ」を入れると全期間の列に戻る"}
                </div>
            </div>
        </div>
    );
}
