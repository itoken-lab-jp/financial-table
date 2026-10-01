/**
 * update ごとの受け取り。続きの読み込みを 1 回分ずつ受けて足し合わせ、表を組む受け取りを持ち続ける。
 *
 * 1 回分ずつ受ける形（fetchMoreData(false)）では、ホストの手元に残るのは最後の回だけ。読み込みのあとに来る update
 * （大きさの変更・表示の切り替え・書式の変更・見る人の選択の保存の応答）は最後の回の DataView を持って来るので、
 * それを全体として読み直すと、表が最後の回だけに縮む。そこで
 * - データの変わらない update（type に Data が無い）は読み直さない
 * - データの update でも、届いた中身が前に読んだ回と同じ（指紋が同じ）なら、同じ回がもう一度届いたとみて読み直さない
 *   （書式の保存値はここで持たず、表を組むときに DataView から読むので、書式の変更は効く）。金額のメジャーの名前と
 *   メジャーごとの保存値（横持ちの比較順・金額の持ち方）だけは受け取りに入るので、変わっていれば当て直す（1 回で読み切った
 *   受け取りなら届いた回をそのまま使い、足し合わせた受け取りなら名前と保存値だけを当て直す）
 * - 見分けられないもの：フィルターを変えた結果が、前の読み込みの最後の回と行も値も同じになると、同じ回がもう一度届いたとみる
 *   （最後の回の頭には前の回の最後の葉が重なって届くので、組織や科目で区切りよく絞った結果とはまず一致しない。docs）
 * - 続きの回（Segment）は足し合わせる。前の回までに読んだ葉は読まない（続きの回の重なり）
 * - 読み込みの途中にほかの update が来ても、続きを二重に取りに行かない。読み込み中でないときに届いた続きの回は読まない
 * ホストに依らない（続きの取り方を渡す）ので、流れをテストで確かめられる（test/loading.test.ts）
 */
import powerbi from "powerbi-visuals-api";

import { InputData, mergeInput, readInput, refreshMeasures } from "./data";

/** update の種類のうちデータ（VisualUpdateType.Data。const enum は実行時に無いので数で持つ） */
export const DATA_UPDATE = 2;
/** 続きの読み込みの 1 回分が届いた update（VisualDataChangeOperationKind.Segment） */
export const SEGMENT = 2;

export interface LoadResult {
    /** 表を組む受け取り（読み込みの途中は、それまでに届いた分） */
    input: InputData;
    /** 続きを読み込んでいる */
    loading: boolean;
    /** 取り切れない（上限）。届いた分で組んで知らせる */
    cutOff: boolean;
    /** 表を組み直すか。続きの回が届いて、まだ残りがあるときは組み直さない */
    rebuild: boolean;
}

export class InputLoader {
    /** 表を組む受け取り（読み込みの途中は、それまでに届いた分） */
    private input: InputData | null = null;
    /** 続きを読み込んでいる間の足し合わせ */
    private loadingInput: InputData | null = null;
    private cutOff = false;
    /** 最後に読んだ回の指紋（同じ回がもう一度届いたら読み直さない）と、メジャーの名前と保存値の印 */
    private lastFingerprint: string | null = null;
    private lastMeasureKey: string | null = null;
    /** 表を組む受け取りが、何回かに分かれて届いたのを足し合わせたものか（読み直すと最後の回だけに縮む） */
    private accumulated = false;

    /** fetchMore：続きを取りに行く（ホストの fetchMoreData(false)）。取り切れなければ false */
    constructor(private readonly fetchMore: () => boolean) {}

    /** 今の受け取り（見る人がメニューで選んだときの組み直し）。まだ何も届いていなければ null */
    public current(): LoadResult | null {
        return this.input ? { input: this.input, loading: this.loadingInput !== null, cutOff: this.cutOff, rebuild: true } : null;
    }

    /** update が届いた。type は VisualUpdateOptions.type、operationKind は VisualUpdateOptions.operationKind */
    public receive(dataView: powerbi.DataView | undefined, type: number | undefined, operationKind: number | undefined): LoadResult {
        const dataUpdate = type === undefined || (type & DATA_UPDATE) !== 0;
        const current = this.current();
        if (current && !dataUpdate) return current;
        // 読み込み中でないときに届いた続きの回（前の問い合わせの残りなど）は読まない
        if (current && operationKind === SEGMENT && this.loadingInput === null) return current;
        const continuing = operationKind === SEGMENT && this.loadingInput !== null;
        const part = readInput(dataView, continuing ? this.loadingInput!.leafKeys : undefined);
        if (current && !continuing && part.fingerprint === this.lastFingerprint) {
            // 同じ行と値がもう一度届いた。メジャーの名前か保存値だけが変わっていれば当て直す
            if (part.measureKey !== this.lastMeasureKey) {
                this.lastMeasureKey = part.measureKey;
                if (this.accumulated) refreshMeasures(this.input!, part);
                else {
                    if (this.loadingInput === this.input) this.loadingInput = part;
                    this.input = part;
                }
            }
            return this.current()!;
        }
        this.lastFingerprint = part.fingerprint;
        this.lastMeasureKey = part.measureKey;
        const merged = continuing ? mergeInput(this.loadingInput!, part) : part;
        this.accumulated = continuing;
        let loading = false;
        this.cutOff = false;
        if (merged.truncated) {
            if (this.fetchMore()) loading = true;
            else this.cutOff = true;
        }
        this.loadingInput = loading ? merged : null;
        this.input = merged;
        return { input: merged, loading, cutOff: this.cutOff, rebuild: !(loading && continuing) };
    }
}

/** 指標の小計の設定を保存してから、この時間は読み直しを待つ（知らせを出し、警告は出さない） */
export const SUBTOTAL_WAIT_MS = 15000;

/**
 * 指標の値（組織 × イベント × 月の小計）が届いていないときにどうするか（visual.ts）。
 * - persist：小計を入れる設定を保存する（このセッションでその段の欄にまだ保存しておらず、保存済みでもない）
 * - wait：保存したばかりなので、読み直しを待つ（知らせを出す）
 * - warn：保存したのに届かない・保存済みなのに届かない（知らせる。保存と読み直しの輪にしない）
 */
export function subtotalAction(missing: { queryName: string; enabled: boolean }, persistedAt: number | undefined, now: number): "persist" | "wait" | "warn" {
    if (persistedAt === undefined) return missing.enabled ? "warn" : "persist";
    return now - persistedAt < SUBTOTAL_WAIT_MS ? "wait" : "warn";
}
