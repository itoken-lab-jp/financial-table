/**
 * 出す期間を選ぶダイアログ。表の上のバーの「期間を選ぶ」のボタンのすぐ下に出す（見せる科目のダイアログと同じ形）。
 *
 * データのある年度（新しい順）ごとに、通期・累計、上期・下期、四半期とその月をチェックで並べる。どの列も、チェックを入れると表に出る
 * （四半期のチェックは四半期の列で、その月の列とは別）。年度をまたいで選べる。出している月を足した合計の列も選べる。
 * 選ぶたびに表に効かせて保存する（visual.ts）。年度の名前の行を押すと開き閉じ（はじめは、列を出している年度だけ開く）。
 * 閉じる・外を押す・Esc・Tab（Shift+Tab）でダイアログの外へ出る・ビジュアルの大きさが変わると閉じる
 */
import * as React from "react";

import { PeriodChoiceYear } from "./periods";
import { PeriodPatch } from "./visualState";

export interface PeriodDialogProps {
    years: PeriodChoiceYear[];
    /** 今出している列の名前と、合計の列を出しているか、見る人が選んでいるか */
    current: string[];
    total: boolean;
    picked: boolean;
    left: number;
    top: number;
    onPick: (patch: PeriodPatch) => void;
    onClose: (returnFocus: boolean) => void;
}

export function PeriodDialog({ years, current, total, picked, left, top, onPick, onClose }: PeriodDialogProps): React.JSX.Element {
    const ref = React.useRef<HTMLDivElement>(null);
    const idBase = React.useId();
    const shown = new Set(current);
    const [open, setOpen] = React.useState<Set<string>>(() => new Set(years.filter((y) => y.columns.some((c) => shown.has(c.id))).map((y) => y.key)));
    // ボタンの下に出し、右がはみ出す分は左へ寄せる（ビジュアルの枠の中に収める。描いてから幅を測る）
    const [x, setX] = React.useState(left);
    React.useLayoutEffect(() => {
        const pop = ref.current;
        const root = pop?.offsetParent as HTMLElement | null;
        if (pop && root) setX(Math.max(0, Math.min(left, root.clientWidth - pop.offsetWidth - 2)));
    }, [left]);
    React.useEffect(() => {
        ref.current?.querySelector<HTMLElement>("button, input")?.focus();
    }, []);
    // 外を押す・Tab でダイアログの外へフォーカスが移ったら閉じる（開くボタンは自分で開き閉じする）
    React.useEffect(() => {
        const outside = (target: HTMLElement) => !ref.current?.contains(target) && !target.closest?.(".ft-period-open");
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
    }, [onClose]);
    const check = (id: string, label: string) => (
        <label key={id} className="ft-pick-account">
            <input type="checkbox" checked={shown.has(id)} onChange={() => onPick({ kind: "toggle", id })} />
            {label}
        </label>
    );
    return (
        <div
            ref={ref}
            className="ft-pick-pop"
            role="dialog"
            aria-label="出す期間"
            style={{ left: x, top, maxHeight: `calc(100% - ${top}px - 0.5em)` }}
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
                <span>出す期間</span>
                <button type="button" className="ft-pick-close" onClick={() => onClose(true)}>
                    閉じる
                </button>
            </div>
            <div className="ft-pick-body ft-period-top">
                <label className="ft-pick-account">
                    <input type="checkbox" checked={total} onChange={() => onPick({ kind: "total" })} />
                    出している月の合計の列
                </label>
                {picked && (
                    <div className="ft-pick-actions">
                        <button
                            type="button"
                            onClick={(e) => {
                                // 押すとこのボタンが消えるので、フォーカスをダイアログの中に残す
                                e.currentTarget.closest(".ft-pick-pop")?.querySelector<HTMLElement>(".ft-pick-parent")?.focus();
                                onPick({ kind: "reset" });
                            }}
                        >
                            もとに戻す
                        </button>
                    </div>
                )}
            </div>
            {years.map((year, index) => {
                const expanded = open.has(year.key);
                const listId = `${idBase}-${index}`;
                const count = year.columns.filter((c) => shown.has(c.id)).length;
                const of = (kind: string) => year.columns.filter((c) => c.kind === kind);
                const months = of("month");
                return (
                    <div key={year.key} className="ft-pick-group">
                        <button
                            type="button"
                            className="ft-pick-parent"
                            aria-expanded={expanded}
                            aria-controls={expanded ? listId : undefined}
                            onClick={() =>
                                setOpen((before) => {
                                    const next = new Set(before);
                                    if (next.has(year.key)) next.delete(year.key);
                                    else next.add(year.key);
                                    return next;
                                })
                            }
                        >
                            <span className="ft-pick-mark" aria-hidden="true">
                                {expanded ? "▾" : "▸"}
                            </span>
                            {year.name}
                            <span className="ft-pick-count">{count > 0 ? `${count} 列` : "出していない"}</span>
                        </button>
                        {expanded && (
                            <div id={listId} className="ft-pick-body" role="group" aria-label={year.name}>
                                <div className="ft-pick-actions">
                                    <button type="button" onClick={() => onPick({ kind: "set", ids: year.columns.map((c) => c.id), on: true })}>
                                        すべて選ぶ
                                    </button>
                                    <button type="button" onClick={() => onPick({ kind: "set", ids: year.columns.map((c) => c.id), on: false })}>
                                        すべて外す
                                    </button>
                                </div>
                                <div className="ft-period-line">{[...of("year"), ...of("ytd")].map((c) => check(c.id, c.label))}</div>
                                <div className="ft-period-line">{of("half").map((c) => check(c.id, c.label))}</div>
                                {of("quarter").map((q) => (
                                    <div key={q.id} className="ft-period-line">
                                        {check(q.id, q.label)}
                                        {months.filter((m) => q.months.includes(m.months[0])).map((m) => check(m.id, m.label))}
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                );
            })}
        </div>
    );
}
