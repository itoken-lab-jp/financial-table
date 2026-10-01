/**
 * 計算の行の式。科目の表の「式」の列に書く。
 *
 *   [営業利益] / [売上高]
 *   [4000] + [4300]
 *   ([A] - [B]) * 100
 *
 * [ ] の中は行のコード（無ければ科目名）。括弧の外の英字で始まる語（OP・SALES など）もコードとして読む。
 * 数字だけの語は数として読む（コードが数字なら [ ] で囲む）。先頭の = は無視する（Excel の書き方のまま貼れるように）。
 *
 * 空欄（値が無い）の扱いは DAX に合わせる：足す・引くは空欄を 0 として計算し、両方空欄なら空欄。
 * 掛け算は片方が空欄なら空欄。割り算は、割られる数が空欄なら空欄、割る数が空欄か 0 なら空欄（Infinity を表に出さない）。
 */

export type Expr =
    | { kind: "num"; value: number }
    | { kind: "ref"; name: string }
    | { kind: "neg"; operand: Expr }
    | { kind: "bin"; op: "+" | "-" | "*" | "/"; left: Expr; right: Expr };

export class FormulaError extends Error {}

type Token = { t: "num"; v: number } | { t: "ref"; v: string } | { t: "op"; v: string };

function tokenize(text: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;
    const src = text.trim().replace(/^=/, "");
    while (i < src.length) {
        const c = src[i];
        if (/\s/.test(c)) {
            i++;
        } else if (c === "[") {
            const end = src.indexOf("]", i + 1);
            if (end < 0) throw new FormulaError("[ が閉じていない");
            const name = src.slice(i + 1, end).trim();
            if (!name) throw new FormulaError("[ ] の中が空");
            tokens.push({ t: "ref", v: name });
            i = end + 1;
        } else if ("+-*/()".includes(c)) {
            tokens.push({ t: "op", v: c });
            i++;
        } else if (/[0-9.]/.test(c)) {
            let j = i;
            while (j < src.length && /[0-9.]/.test(src[j])) j++;
            const value = Number(src.slice(i, j));
            if (!Number.isFinite(value)) throw new FormulaError(`数が読めない：${src.slice(i, j)}`);
            tokens.push({ t: "num", v: value });
            i = j;
        } else if (/[A-Za-z_]/.test(c)) {
            let j = i;
            while (j < src.length && /[A-Za-z0-9_.]/.test(src[j])) j++;
            tokens.push({ t: "ref", v: src.slice(i, j) });
            i = j;
        } else {
            throw new FormulaError(`読めない字：${c}`);
        }
    }
    return tokens;
}

/** 式を読む。読めなければ FormulaError */
export function parseFormula(text: string): Expr {
    const tokens = tokenize(text);
    let pos = 0;
    const peek = () => tokens[pos];
    const isOp = (v: string) => peek()?.t === "op" && peek().v === v;

    const primary = (): Expr => {
        const token = peek();
        if (!token) throw new FormulaError("式が途中で終わっている");
        if (token.t === "num") {
            pos++;
            return { kind: "num", value: token.v };
        }
        if (token.t === "ref") {
            pos++;
            return { kind: "ref", name: token.v };
        }
        if (token.v === "(") {
            pos++;
            const inner = additive();
            if (!isOp(")")) throw new FormulaError(") が足りない");
            pos++;
            return inner;
        }
        if (token.v === "-" || token.v === "+") {
            pos++;
            const operand = primary();
            return token.v === "-" ? { kind: "neg", operand } : operand;
        }
        throw new FormulaError(`ここに ${token.v} は置けない`);
    };

    const multiplicative = (): Expr => {
        let left = primary();
        while (isOp("*") || isOp("/")) {
            const op = peek().v as "*" | "/";
            pos++;
            left = { kind: "bin", op, left, right: primary() };
        }
        return left;
    };

    const additive = (): Expr => {
        let left = multiplicative();
        while (isOp("+") || isOp("-")) {
            const op = peek().v as "+" | "-";
            pos++;
            left = { kind: "bin", op, left, right: multiplicative() };
        }
        return left;
    };

    if (tokens.length === 0) throw new FormulaError("式が空");
    const expr = additive();
    if (pos < tokens.length) throw new FormulaError(`余分な字：${JSON.stringify(tokens[pos].v)}`);
    return expr;
}

/** 式が引いている行の名前（重複なし、出てきた順） */
export function formulaRefs(expr: Expr): string[] {
    const out: string[] = [];
    const walk = (e: Expr) => {
        if (e.kind === "ref") {
            if (!out.includes(e.name)) out.push(e.name);
        } else if (e.kind === "neg") walk(e.operand);
        else if (e.kind === "bin") {
            walk(e.left);
            walk(e.right);
        }
    };
    walk(expr);
    return out;
}

/** 空欄は null。resolve は参照の値（空欄なら null） */
export function evaluateFormula(expr: Expr, resolve: (name: string) => number | null): number | null {
    switch (expr.kind) {
        case "num":
            return expr.value;
        case "ref":
            return resolve(expr.name);
        case "neg": {
            const v = evaluateFormula(expr.operand, resolve);
            return v === null ? null : -v;
        }
        case "bin": {
            const a = evaluateFormula(expr.left, resolve);
            const b = evaluateFormula(expr.right, resolve);
            if (expr.op === "/") {
                if (a === null || b === null || b === 0) return null;
                return a / b;
            }
            if (expr.op === "*") return a === null || b === null ? null : a * b;
            if (a === null && b === null) return null;
            const x = a ?? 0;
            const y = b ?? 0;
            return expr.op === "+" ? x + y : x - y;
        }
    }
}
