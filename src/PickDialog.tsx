/**
 * 見せる科目を選ぶダイアログ。表の上のバーの「科目を選ぶ」で開く。区分・中分類ごとに、出す科目のチェックと、
 * 選ばなかった科目の見せ方（その他にまとめる・選んだ科目をうちにする）を選ぶ。選ぶたびに表に効かせて保存する（visual.ts）。
 *
 * 「科目を選ぶ」のボタンのすぐ下に出す（右がはみ出す分は左へ寄せる）。ボタンはバーの中ほどにあるので、左の行の名前を覆わず、
 * 選んだ結果をダイアログを開いたまま見られる。
 * 区分・中分類は名前の行だけを並べ、押すと科目が開く（区分の多い表で長くならないように）。はじめは選んでいる親だけを開く（親が 1 つなら開く）。
 * 閉じる・外を押す・Esc・Tab（Shift+Tab）でダイアログの外へ出る・ビジュアルの大きさが変わると閉じ、キーボードで閉じたら開いたボタンにフォーカスを戻す（Desktop では Esc がビジュアルに届かないので、閉じるのボタン）
 */
import * as React from "react";

import { PickMode, PickParent, RowPick } from "./picks";
import { PickPatch } from "./visualState";

export interface PickDialogProps {
    parents: PickParent[];
    current: RowPick[];
    /** 開いたボタンの左端と、ボタンのすぐ下の位置（ビジュアルの枠からの px） */
    left: number;
    top: number;
    onPick: (parent: string, accounts: string[], patch: PickPatch) => void;
    onClose: (returnFocus: boolean) => void;
}

const MODES: Array<{ mode: PickMode; label: string }> = [
    { mode: "others", label: "残りをその他にまとめる" },
    { mode: "under", label: "選んだ科目をうちにする" },
];

export function PickDialog({ parents, current, left, top, onPick, onClose }: PickDialogProps): React.JSX.Element {
    const ref = React.useRef<HTMLDivElement>(null);
    // 開いた科目の欄の id（区分の名前は空白を含みうるので、id にしない）
    const idBase = React.useId();
    const [open, setOpen] = React.useState<Set<string>>(() =>
        parents.length === 1 ? new Set([parents[0].code]) : new Set(current.filter((p) => parents.some((q) => q.code === p.parent)).map((p) => p.parent)),
    );
    // ボタンの下に出し、右がはみ出す分は左へ寄せる（ビジュアルの枠の中に収める。描いてから幅を測る）
    const [x, setX] = React.useState(left);
    React.useLayoutEffect(() => {
        const pop = ref.current;
        const root = pop?.offsetParent as HTMLElement | null;
        if (pop && root) setX(Math.max(0, Math.min(left, root.clientWidth - pop.offsetWidth - 2)));
    }, [left]);
    // 開いたら最初の操作できる所へ
    React.useEffect(() => {
        ref.current?.querySelector<HTMLElement>("button, input")?.focus();
    }, []);
    // 外を押したら閉じる（開くボタンは自分で開き閉じする）
    React.useEffect(() => {
        const down = (e: MouseEvent) => {
            const target = e.target as HTMLElement;
            if (ref.current?.contains(target) || target.closest?.(".ft-pick-open")) return;
            onClose(false);
        };
        // Tab・Shift+Tab でダイアログの外へフォーカスが移ったら閉じる（開くボタンを通って、その先へ移ったときも）。
        // Power BI がフォーカスを取ったとき（ビジュアルの外）は、この文書の中に移らないので閉じない
        const focus = (e: FocusEvent) => {
            const target = e.target as HTMLElement;
            if (ref.current?.contains(target) || target.closest?.(".ft-pick-open")) return;
            onClose(false);
        };
        document.addEventListener("mousedown", down, true);
        document.addEventListener("focusin", focus, true);
        return () => {
            document.removeEventListener("mousedown", down, true);
            document.removeEventListener("focusin", focus, true);
        };
    }, [onClose]);
    const toggleOpen = (code: string) =>
        setOpen((before) => {
            const next = new Set(before);
            if (next.has(code)) next.delete(code);
            else next.add(code);
            return next;
        });
    return (
        <div
            ref={ref}
            className="ft-pick-pop"
            role="dialog"
            aria-label="見せる科目"
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
                <span>見せる科目</span>
                <button type="button" className="ft-pick-close" onClick={() => onClose(true)}>
                    閉じる
                </button>
            </div>
            {parents.map((parent, index) => {
                const pick = current.find((p) => p.parent === parent.code) ?? null;
                const accounts = parent.accounts.map((a) => a.code);
                const chosen = new Set(pick ? pick.codes : accounts);
                const count = accounts.filter((c) => chosen.has(c)).length;
                const expanded = open.has(parent.code);
                const change = (patch: PickPatch) => onPick(parent.code, accounts, patch);
                const listId = `${idBase}-${index}`;
                return (
                    <div key={parent.code} className="ft-pick-group" style={{ paddingLeft: `${parent.depth * 1}em` }}>
                        <button
                            type="button"
                            className="ft-pick-parent"
                            aria-expanded={expanded}
                            aria-controls={expanded ? listId : undefined}
                            onClick={() => toggleOpen(parent.code)}
                        >
                            <span className="ft-pick-mark" aria-hidden="true">
                                {expanded ? "▾" : "▸"}
                            </span>
                            {parent.name}
                            <span className="ft-pick-count">
                                {pick ? `${count}/${accounts.length}${pick.mode === "under" ? "・うち" : "・その他"}` : "すべて"}
                            </span>
                        </button>
                        {expanded && (
                            <div id={listId} className="ft-pick-body">
                                <div className="ft-pick-modes" role="radiogroup" aria-label={`${parent.name}の選ばなかった科目`}>
                                    {MODES.map(({ mode, label }) => (
                                        <label key={mode} className="ft-pick-mode">
                                            <input
                                                type="radio"
                                                name={`${listId}-mode`}
                                                checked={(pick?.mode ?? "others") === mode}
                                                onChange={() => change({ kind: "mode", mode })}
                                            />
                                            {label}
                                        </label>
                                    ))}
                                </div>
                                <div className="ft-pick-actions">
                                    <button type="button" onClick={() => change({ kind: "all" })}>
                                        すべて選ぶ
                                    </button>
                                    <button type="button" onClick={() => change({ kind: "none" })}>
                                        すべて外す
                                    </button>
                                    {pick && (
                                        <button
                                            type="button"
                                            onClick={(e) => {
                                                // 押すとこのボタンが消えるので、フォーカスを区分の名前の行へ移す（ダイアログの外へ落とさない）
                                                e.currentTarget.closest(".ft-pick-group")?.querySelector<HTMLElement>(".ft-pick-parent")?.focus();
                                                change({ kind: "reset" });
                                            }}
                                        >
                                            もとに戻す
                                        </button>
                                    )}
                                </div>
                                {parent.accounts.map((account) => (
                                    <label key={account.code} className="ft-pick-account">
                                        <input
                                            type="checkbox"
                                            checked={chosen.has(account.code)}
                                            onChange={() => change({ kind: "toggle", code: account.code })}
                                        />
                                        {account.name}
                                    </label>
                                ))}
                            </div>
                        )}
                    </div>
                );
            })}
        </div>
    );
}
