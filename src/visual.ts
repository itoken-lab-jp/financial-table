/*
 *  Power BI Visual
 *  Licensed under the MIT License.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import { hasBrowserMenu } from "./shared/copyImage";
import { themeOf } from "./theme";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import "./../style/visual.less";
import { toRootCoordinates } from "./shared/tooltip";
import { App, CellTooltip } from "./App";
import { SelectionNodes } from "./data";
import { InputLoader, SUBTOTAL_WAIT_MS, subtotalAction } from "./loading";
import { SelectTarget, selectableView, nextSelection, reconcileSelection, sameSelection, targetKey } from "./selection";
import { DataDrivenFormat, VisualFormattingSettingsModel } from "./settings";
import { INDICATOR_NOTICE, INDICATOR_SUBTOTAL_WARNING, LOADING_NOTICE, TRUNCATED_NOTICE, ViewModel, transform } from "./viewModel";
import {
    EMPTY_VISUAL_STATE,
    VISUAL_STATE_OBJECT,
    ComparePatch,
    PeriodPatch,
    PickPatch,
    VisualState,
    chooseCompare,
    choosePeriods,
    choosePick,
    readVisualState,
    serializeVisualState,
    toPersistedProperties,
    toggled,
    withoutStale,
} from "./visualState";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;

/** 警告のアイコンに出す件数の上限。全件を表示すると警告欄が長くなるため制限する */
const MAX_WARNINGS = 5;
/** 保存の応答をこの時間待っても来なければ、保存に失敗したものとして待つのをやめる */
const PENDING_TIMEOUT_MS = 5000;

export class Visual implements IVisual {
    private root: Root;
    private element: HTMLElement;
    private host: IVisualHost;
    /** セルのツールチップ。表は単位で丸めるので、正確な値を Power BI のツールチップで出す */
    private tooltip: CellTooltip;
    private events: IVisualEventService;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    /** 最後の update（見る人がメニューで選んだとき、同じ内容で作り直す） */
    private lastOptions: VisualUpdateOptions | null = null;
    /** 見る人が選んだ主と比較。ページを移ってもレポートから読み直す */
    private visualState: VisualState = EMPTY_VISUAL_STATE;
    /** persistProperties 直後の値。保存が返る前の古い dataView で選択を巻き戻さないための印 */
    private pendingVisualState: string | null = null;
    private pendingAt = 0;
    /** 同じ状態を何度も当て直さないための印 */
    private lastRestoredVisualState: string | null = null;
    /** 受け取り（続きの読み込みを 1 回分ずつ足し合わせ、読み込みのあとの update でも持ち続ける。loading.ts） */
    private loader: InputLoader;
    /** 最後に組んだ表の書式ペインの選択肢（続きの回の途中で表を組み直さないときも、書式ペインに入れる） */
    private lastFormat: DataDrivenFormat | null = null;
    /**
     * 指標の小計（科目の最初の段の小計）を入れる設定を保存した段の欄と時刻。保存はこのセッションで段の欄ごとに 1 回だけ
     *
     */
    private subtotalsPersisted = new Map<string, number>();
    /** 小計の設定を保存したあと、待つ時間がたったら見直す時計（update が来なくても、届かなければ知らせる） */
    private subtotalTimer: number | null = null;
    /** 選択。押したものの並びと、最後に描いた表（選択を変えたら描き直すだけで、組み直さない） */
    private selectionManager: ISelectionManager;
    private selection: SelectTarget[] = [];
    private lastRender: { viewModel: ViewModel; viewport: powerbi.IViewport } | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.element = options.element;
        this.events = options.host.eventService;
        const tooltipService = options.host.tooltipService;
        this.tooltip = {
            show: (items, clientX, clientY, move) => {
                if (!tooltipService.enabled()) return;
                const args = { coordinates: toRootCoordinates(clientX, clientY, this.element), isTouchEvent: false, dataItems: items, identities: [] as powerbi.visuals.ISelectionId[] };
                if (move) tooltipService.move(args);
                else tooltipService.show(args);
            },
            hide: () => tooltipService.hide({ isTouchEvent: false, immediately: false }),
        };
        this.formattingSettingsService = new FormattingSettingsService();
        this.root = createRoot(options.element);
        this.loader = new InputLoader(() => this.host.fetchMoreData(false));
        this.selectionManager = options.host.createSelectionManager();
        // ブックマークやほかのビジュアルで選択が変わったら、表の見た目を合わせる：空なら解き、表の選択と違う選択 ID が来たら（ブックマークが
        // 別の選択を戻した）見た目を解く。何を選んでいたかは、表の見た目には戻さない
        this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => {
            if (this.selection.length === 0) return;
            const nodes = this.lastRender?.viewModel.selectionNodes;
            const mine = nodes ? this.selection.flatMap((t) => this.selectionIds(t, nodes)) : [];
            // 件数や字の並びではなく、equals で両方向に含まれるかで比べる（区分と科目を重ねて選ぶと同じ選択 ID が 2 度入る）
            const same = mine.every((m) => ids.some((id) => id.equals(m))) && ids.every((id) => mine.some((m) => m.equals(id)));
            if (ids.length === 0 || !same) {
                this.selection = [];
                this.renderApp();
            }
        });
    }

    public update(options: VisualUpdateOptions): void {
        // レンダリングイベントは認定要件。必ず started / finished(failed) を対で呼ぶ。
        this.events.renderingStarted(options);
        try {
            this.lastOptions = options;
            this.restoreVisualState(options.dataViews?.[0]);
            this.build(options, true);
            this.events.renderingFinished(options);
        } catch (error) {
            console.error("update failed", error);
            this.events.renderingFailed(options, String(error));
        }
    }

    /** showWarnings：警告のアイコンを出し直すか（既定は update のとき。見る人のメニューの操作では出し直さない） */
    private build(options: VisualUpdateOptions, fromUpdate: boolean, showWarnings = fromUpdate): void {
        const dataView = options.dataViews?.[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        // 色の既定はレポートのテーマから（書式ペインで決めた色はそのまま）
        const theme = themeOf(this.host.colorPalette as ISandboxExtendedColorPalette);
        const objects = dataView?.metadata?.objects;
        this.formattingSettings.applyTheme(theme, (object, property) => objects?.[object]?.[property] !== undefined);

        // 行が 30,000 を超えると、Power BI は続きを残して届ける（metadata.segment）。1 回分ずつ取りにいき（fetchMoreData(false)）、
        // 手元で足し合わせる。
        // 読み込みのあとの update（大きさの変更・書式の変更・保存の応答）では読み直さない。続きの回の途中は表を組み直さない（2-13）
        const state = fromUpdate ? this.loader.receive(dataView, options.type, options.operationKind) : this.loader.current();
        if (state && !state.rebuild) {
            if (this.lastFormat) this.formattingSettings.applyData(this.lastFormat);
            return;
        }
        const viewModel: ViewModel = transform(dataView, this.formattingSettings, this.visualState, state?.input, theme);
        // 主と比較・年度の選択肢はデータ次第なので、populate のあとに流し込む
        this.formattingSettings.applyData(viewModel.format);
        this.lastFormat = viewModel.format;
        const loading = state?.loading ?? false;

        // 作り手が書式ペインで主・比較の列を決め直して古くなった見る人の選択は捨て、捨てた状態を保存し直す
        // （保存が返す update では、もう捨てるものが無いので輪にならない）。データで決まる既定が変わっただけでは捨てない。
        // 読み込みの途中は保存し直さない（続きの読み込みの update と、保存が呼ぶ update を重ねない）
        if (viewModel.menu && !viewModel.truncated) {
            const current = withoutStale(this.visualState, viewModel.menu.mainBase, viewModel.periodPicker?.base);
            if (current !== this.visualState) this.persistVisualState(current);
        }

        // 指標の値は組織 × イベント × 月の小計で読む。小計は宣言の既定で切ってあるので、指標があって小計が届かなければ、科目の
        // 最初の段の小計を入れる設定を保存する。読み込みの途中は保存しない
        const missing = viewModel.indicatorSubtotals;
        if (missing && !loading) {
            const action = subtotalAction(missing, this.subtotalsPersisted.get(missing.queryName), Date.now());
            if (action === "persist") {
                this.persistSubtotals(missing.queryName);
                this.scheduleSubtotalCheck();
            }
            if (action === "warn") viewModel.warnings.unshift(INDICATOR_SUBTOTAL_WARNING);
            else viewModel.notice = INDICATOR_NOTICE;
        }

        // 取り切れない（上限）ときは、届いた分で組んで知らせる。読み込みの途中の警告は途中までの分のものなので、読み切ってから出す
        if (loading) viewModel.notice = LOADING_NOTICE;
        if (state?.cutOff) viewModel.warnings.unshift(TRUNCATED_NOTICE);
        if (showWarnings && !loading && viewModel.warnings.length > 0) {
            const shown = viewModel.warnings.slice(0, MAX_WARNINGS);
            const rest = viewModel.warnings.length - shown.length;
            this.host.displayWarningIcon("表に入れられなかったものがあります", [...shown, ...(rest > 0 ? [`ほか ${rest} 件`] : [])].join("\n"));
        }

        this.lastRender = { viewModel, viewport: options.viewport };
        if (this.selection.length > 0 && !loading) this.reconcile(viewModel, fromUpdate);
        this.renderApp();
    }

    /**
     * 表が変わったら（行・組織を閉じた、年度を変えた、データが変わった）選択を合わせ直す。無くなったもの・月の変わった期間は外し、
     * 変わったら Power BI にも渡し直す。update のときにホストの選択が空なら（ほかのビジュアルを押して解かれた）表の選択も解く
     */
    private reconcile(viewModel: ViewModel, fromUpdate: boolean): void {
        const nodes = viewModel.selectionNodes;
        if (fromUpdate && !this.selectionManager.hasSelection()) {
            this.selection = [];
            return;
        }
        // データが空（節点が無い）なら、表の選択を捨て、Power BI の選択も解く（表だけ捨てると、ほかのビジュアルが絞られたまま残った）
        if (!nodes) {
            this.selection = [];
            void Promise.resolve(this.selectionManager.clear()).catch((error) => console.error("選択を解けませんでした", error));
            return;
        }
        const view = selectableView(viewModel);
        // 選択 ID の無くなったもの（フィルターでその月・科目のデータが消えた）も外す（残すと押しても解けなかった）
        const next = reconcileSelection(this.selection, view).filter((t) => this.selectionIds(t, nodes).length > 0);
        if (sameSelection(next, this.selection)) return;
        this.selection = next;
        this.sendSelection(nodes);
    }

    /** 表の選択を Power BI に渡す（選択 ID が無ければ解く） */
    private sendSelection(nodes: SelectionNodes): void {
        const ids = this.selection.flatMap((t) => this.selectionIds(t, nodes));
        const done = ids.length > 0 ? this.selectionManager.select(ids, false) : this.selectionManager.clear();
        void Promise.resolve(done).catch((error) => console.error("選択できませんでした", error));
    }

    /** 最後に組んだ表を描く（選択を変えたときは、これだけを呼ぶ） */
    private renderApp(): void {
        if (!this.lastRender) return;
        const { viewModel, viewport } = this.lastRender;
        // 操作できない場所（ダッシュボードのタイルなど）では、メニューを出さず名前だけ出す
        const allowInteractions = this.host.hostCapabilities?.allowInteractions !== false;
        const menu = viewModel.menu;
        const nodes = viewModel.selectionNodes;
        this.root.render(
            React.createElement(App, {
                viewModel,
                viewport,
                tooltip: this.tooltip,
                selection: this.selection,
                onSelect: allowInteractions && nodes ? (target: SelectTarget, multi: boolean) => this.select(target, multi, nodes) : undefined,
                onContextMenu:
                    allowInteractions && nodes
                        ? (target: SelectTarget | null, x: number, y: number) => {
                              // 押した所に結びついたメニュー（含める・含めない・ドリルスルー）。1 つの選択 ID にならない所（区分・四半期）は既定のメニュー
                              const ids = target ? this.selectionIds(target, nodes) : [];
                              this.selectionManager.showContextMenu(ids.length === 1 ? ids[0] : ({} as ISelectionId), { x, y });
                          }
                        : undefined,
                onClearSelection: allowInteractions ? () => this.clearSelection() : undefined,
                // 「画像としてコピー」（操作できる所だけ）。ブラウザーのメニューが出る所（サービス）では、ボタンの右クリックで「画像をコピー」も
                copyImage: allowInteractions ? { browserMenu: hasBrowserMenu(this.host.hostEnv) } : undefined,
                onChooseMain:
                    allowInteractions && menu ? (event: string) => this.changeVisualState({ ...this.visualState, main: event, mainBase: menu.mainBase }) : undefined,
                // 比較の列（上のバーの「比較」は全期間、期間の見出しの ▾ はその期間だけ）。まだ選んでいなければ既定の 1 列から書く
                onChooseCompare:
                    allowInteractions && menu
                        ? (patch: ComparePatch) => this.changeVisualState(chooseCompare(this.visualState, menu.compares.fallback, menu.compares.view, patch))
                        : undefined,
                // 開き閉じも見る人の選択として保存する（ページを移って戻っても同じ形）
                onToggleOrg: allowInteractions ? (path: string) => this.changeVisualState({ ...this.visualState, openOrgs: toggled(this.visualState.openOrgs, path) }) : undefined,
                onToggleRow: allowInteractions
                    ? (code: string) =>
                          this.changeVisualState({ ...this.visualState, [viewModel.rowStateKey]: toggled(this.visualState[viewModel.rowStateKey], code) })
                    : undefined,
                onSetFolds: allowInteractions ? (patch: Partial<VisualState>) => this.changeVisualState({ ...this.visualState, ...patch }) : undefined,
                // 見せる科目：区分・中分類ごとに選んだ科目と見せ方
                // 出す期間：書式ペインのまま（か、作り手が期間の設定を変えた）なら、今出している列から始める
                onPickPeriods:
                    allowInteractions && viewModel.periodPicker
                        ? (patch: PeriodPatch) => {
                              const picker = viewModel.periodPicker!;
                              this.changeVisualState(
                                  choosePeriods(this.visualState, picker.base, { columns: picker.current, total: picker.total, picked: picker.picked }, patch)
                              );
                          }
                        : undefined,
                onPick: allowInteractions
                    ? (parent: string, accounts: string[], patch: PickPatch) => this.changeVisualState(choosePick(this.visualState, parent, accounts, patch))
                    : undefined,
            })
        );
    }

    /** 押したとき：選択を足し引きして、Power BI に渡す（ほかのビジュアルを絞る）。同じものだけをもう一度押すと解く */
    private select(target: SelectTarget, multi: boolean, nodes: SelectionNodes): void {
        // 選択 ID の無い所（データの無い月の列、値の届かなかった科目）は選ばない（見た目だけ選んでホストの選択が解けた）
        // ただし、解く操作（選んでいるものを Ctrl で押す、1 つだけ選んでいるものをもう一度押す）は通す。解かない操作まで通すと、読み込みの
        // 途中に選択 ID の無いものだけを選んだ見た目になり、Power BI の選択が解けた
        const selected = this.selection.some((t) => targetKey(t) === targetKey(target));
        const releasing = selected && (multi || this.selection.length === 1);
        if (this.selectionIds(target, nodes).length === 0 && !releasing) return;
        this.selection = nextSelection(this.selection, target, multi);
        this.sendSelection(nodes);
        this.renderApp();
    }

    /** 空いた所を押したとき。表に選択が無くても、ホストに選択があれば（ブックマークで戻ったなど）解く */
    private clearSelection(): void {
        if (this.selection.length === 0 && !this.selectionManager.hasSelection()) return;
        this.selection = [];
        void Promise.resolve(this.selectionManager.clear()).catch((error) => console.error("選択を解けませんでした", error));
        this.renderApp();
    }

    /**
     * 選んだものの選択 ID（matrix の節点から）。科目は科目コードの段の節点だけで作る（組織・イベントの段を入れると、その値でも絞る）。
     * 組織のブロックの中なら組織の段の節点も重ねる。セルは科目 × 月（四半期・通期はその月すべて）、列の見出しは月、組織の名前は組織
     */
    private selectionIds(target: SelectTarget, nodes: SelectionNodes): ISelectionId[] {
        const builder = () => this.host.createSelectionIdBuilder();
        const orgNodes = (path: string | null) => (path === null ? [] : (nodes.orgs.get(path) ?? []));
        const monthNodes = (months: number[]) => months.flatMap((m) => nodes.months.get(m) ?? []);
        const withRow = (code: string, org: string | null) => {
            const account = nodes.accounts.get(code);
            if (!account) return null;
            let b = builder().withMatrixNode(account, nodes.rowLevels);
            for (const node of orgNodes(org)) b = b.withMatrixNode(node, nodes.rowLevels);
            return b;
        };
        switch (target.kind) {
            case "row":
                return target.codes.flatMap((code) => {
                    const b = withRow(code, target.org);
                    return b ? [b.createSelectionId()] : [];
                });
            case "cell":
                return target.codes.flatMap((code) =>
                    monthNodes(target.months).flatMap((month) => {
                        const b = withRow(code, target.org);
                        return b ? [b.withMatrixNode(month, nodes.columnLevels).createSelectionId()] : [];
                    })
                );
            case "period":
                return monthNodes(target.months).map((month) => builder().withMatrixNode(month, nodes.columnLevels).createSelectionId());
            case "org": {
                const list = orgNodes(target.path);
                if (list.length === 0) return [];
                let b = builder();
                for (const node of list) b = b.withMatrixNode(node, nodes.rowLevels);
                return [b.createSelectionId()];
            }
        }
    }

    /** 見る人の選択を保存して作り直す。update() からは呼ばない（保存が update を呼び、また保存する輪になる） */
    private changeVisualState(next: VisualState): void {
        if (!this.lastOptions) return;
        this.persistVisualState(next);
        this.build(this.lastOptions, false);
    }

    private persistVisualState(next: VisualState): void {
        this.visualState = next;
        this.pendingVisualState = serializeVisualState(next);
        this.pendingAt = Date.now();
        try {
            this.host.persistProperties({
                merge: [{ objectName: VISUAL_STATE_OBJECT, properties: toPersistedProperties(next), selector: null }],
            });
        } catch (error) {
            // 保存に失敗しても、このセッションの見た目は保つ
            this.pendingVisualState = null;
            console.error("見る人の選択（主と比較・開き閉じ）を保存できませんでした", error);
        }
    }

    /**
     * 指標の小計を入れる設定を保存する：行の小計と段ごとの小計を入れ、科目の最初の段（段の欄の selector）だけを入れる。
     * ほかの段の既定は切ったまま
     */
    private persistSubtotals(queryName: string): void {
        this.subtotalsPersisted.set(queryName, Date.now());
        try {
            this.host.persistProperties({
                merge: [
                    { objectName: "subTotals", properties: { rowSubtotals: true, perRowLevel: true }, selector: null },
                    { objectName: "subTotals", properties: { levelSubtotalEnabled: true }, selector: { metadata: queryName } },
                ],
            });
        } catch (error) {
            console.error("指標の小計の設定を保存できませんでした", error);
        }
    }

    /**
     * 小計の設定を保存してから待つ時間がたったら、表を組み直して見直す（読み直しは呼ばない）。保存が効かず update も来ないと、
     * 「指標の値を読み込んでいます…」が出たままになるので、届いていなければ知らせる
     */
    private scheduleSubtotalCheck(): void {
        if (this.subtotalTimer !== null) window.clearTimeout(this.subtotalTimer);
        this.subtotalTimer = window.setTimeout(() => {
            this.subtotalTimer = null;
            if (this.lastOptions) this.build(this.lastOptions, false, true);
        }, SUBTOTAL_WAIT_MS + 500);
    }

    /** 保存済みの見る人の選択を読み戻す */
    private restoreVisualState(dataView: powerbi.DataView | undefined): void {
        if (this.pendingVisualState !== null && Date.now() - this.pendingAt > PENDING_TIMEOUT_MS) this.pendingVisualState = null;
        const persisted = readVisualState(dataView);
        if (!persisted) {
            // 保存が消えた（保存の無いブックマークに切り替えたなど）なら、選択の無い状態に戻す
            if (this.pendingVisualState === null && this.lastRestoredVisualState !== null) {
                this.visualState = EMPTY_VISUAL_STATE;
                this.lastRestoredVisualState = null;
            }
            return;
        }
        // persistProperties の直後に古い dataView が来ても、選んだばかりの選択を戻さない
        if (this.pendingVisualState !== null && persisted.raw !== this.pendingVisualState) return;
        if (this.pendingVisualState === persisted.raw) this.pendingVisualState = null;
        if (this.lastRestoredVisualState === persisted.raw) return;
        this.visualState = persisted.state;
        this.lastRestoredVisualState = persisted.raw;
    }

    /** 書式設定ペインを開くたび / 値変更のたびに呼ばれる */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public destroy(): void {
        if (this.subtotalTimer !== null) window.clearTimeout(this.subtotalTimer);
        this.root?.unmount();
    }
}
