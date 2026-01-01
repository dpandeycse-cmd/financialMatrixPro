type CFIconPlacement = "only" | "left" | "right";
type CFSurfaceSettingsEx = CFSurfaceSettings & { iconPlacement?: CFIconPlacement };
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";
import { buildMatrixTooltipItems } from "./tooltip";

type DisplayUnit = "auto" | "none" | "thousands" | "millions" | "billions";

type LayoutRowType = "data" | "calc" | "blank";

interface RowStyleOverride {
    bold?: boolean;
    color?: string;
    background?: string;
    italic?: boolean;
    underline?: boolean;
}

interface LayoutRow {
    code: string;
    label?: string;
    parent?: string | null;
    type?: LayoutRowType;
    order?: number;
    formula?: string;
    style?: RowStyleOverride;
}

interface LayoutJson {
    rows?: LayoutRow[];
}

interface ColumnKey {
    key: string;
    measure?: string;
}

type ColumnBucket = "values";

interface BucketedColumnKey extends ColumnKey {
    bucket: ColumnBucket;
    levels: string[];
    columnLevels: string[];
    leafLabel: string;
}

interface RowNode {
    code: string;
    label: string;
    parent: string | null;
    type: LayoutRowType;
    orderIndex: number;
    order?: number;
    formula?: string;
    style?: RowStyleOverride;
    fontFamilyOverride?: string;
    fontSizeOverride?: number;
    children: string[];
    depth: number;
    rowLevel?: number; // 0..n-1 for Row fields; -1 for Group; undefined for custom/layout rows
    groupKey?: string | null;
    isTotal?: boolean;
}

interface Model {
    columns: BucketedColumnKey[];
    rows: RowNode[];
    values: Map<string, Map<string, number | null>>;
    valuesRaw: Map<string, Map<string, any>>;
    cfFieldRaw: Map<string, any>;
    rowFieldRaw: Map<string, Map<string, any>>;
    colFieldRaw: Map<string, Map<string, any>>;
    rowHeaderColumns: string[];
    hasGroup: boolean;
    rowFieldCount: number;
    showColumnTotal?: boolean;
    tooltipMeasureNames?: string[];
    rowSelectionIdByCode?: Map<string, powerbi.extensibility.ISelectionId>;
}

type CFRuleTarget = "rowHeader" | "columnHeader" | "cell";
type CFMatchMode = "any" | "contains" | "equals";
type CFConditionType = "value" | "text";
type CFOperator = "gt" | "gte" | "lt" | "lte" | "eq" | "between" | "contains";
type CFApplyScope = "self" | "entireRow" | "entireColumn";
type CFFormatProp = "background" | "fontColor" | "icon";
type CFFormatStyle = "rules" | "gradient" | "fieldValue";

type CFBuiltinIconKey =
    | "ragGreen" | "ragYellow" | "ragRed"
    | "arrowUp" | "arrowRight" | "arrowDown"
    | "triangleUp" | "triangleDown"
    | "check" | "cross"
    | "circle" | "dash";

const CF_BUILTIN_ICONS: Array<{ key: CFBuiltinIconKey; label: string }> = [
    { key: "ragGreen", label: "RAG: Green" },
    { key: "ragYellow", label: "RAG: Yellow" },
    { key: "ragRed", label: "RAG: Red" },
    { key: "arrowUp", label: "Arrow: Up" },
    { key: "arrowRight", label: "Arrow: Right" },
    { key: "arrowDown", label: "Arrow: Down" },
    { key: "triangleUp", label: "Triangle: Up" },
    { key: "triangleDown", label: "Triangle: Down" },
    { key: "check", label: "Check" },
    { key: "cross", label: "Cross" },
    { key: "circle", label: "Circle" },
    { key: "dash", label: "Dash" }
];

const CF_ICON_ALIASES: Record<string, CFBuiltinIconKey> = {
    // Common Power BI-ish names (best-effort)
    traffichighlight: "circle",
    redcircle: "ragRed",
    yellowcircle: "ragYellow",
    greencircle: "ragGreen",
    trianglehigh: "triangleUp",
    trianglelow: "triangleDown",
    arrowup: "arrowUp",
    arrowright: "arrowRight",
    arrowdown: "arrowDown",
    checkmark: "check",
    crossmark: "cross",
    dash: "dash",
    circle: "circle",
    raggreen: "ragGreen",
    ragyellow: "ragYellow",
    ragred: "ragRed"
};

const CF_ICON_SVG: Record<CFBuiltinIconKey, string> = {
    ragGreen: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><circle cx="8" cy="8" r="6" fill="#107c10"/></svg>`,
    ragYellow: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><circle cx="8" cy="8" r="6" fill="#fce100" stroke="#605e5c" stroke-width="0.5"/></svg>`,
    ragRed: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><circle cx="8" cy="8" r="6" fill="#d13438"/></svg>`,
    arrowUp: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M8 2l4 5H9v7H7V7H4l4-5z" fill="#107c10"/></svg>`,
    arrowRight: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M14 8l-5 4V9H2V7h7V4l5 4z" fill="#605e5c"/></svg>`,
    arrowDown: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M8 14l-4-5h3V2h2v7h3l-4 5z" fill="#d13438"/></svg>`,
    triangleUp: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M8 3l6 10H2L8 3z" fill="#107c10"/></svg>`,
    triangleDown: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M8 13L2 3h12L8 13z" fill="#d13438"/></svg>`,
    check: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M6.2 11.2L3.3 8.3l1.4-1.4 1.5 1.5 5.1-5.1 1.4 1.4-6.5 6.5z" fill="#107c10"/></svg>`,
    cross: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M4.2 3l8.8 8.8-1.2 1.2L3 4.2 4.2 3zm8.8 1.2L4.2 13 3 11.8 11.8 3l1.2 1.2z" fill="#d13438"/></svg>`,
    circle: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><circle cx="8" cy="8" r="5" fill="#605e5c"/></svg>`,
    dash: `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="3" y="7" width="10" height="2" rx="1" fill="#605e5c"/></svg>`
};

type CFEmptyValueHandling = "blank" | "zero";
type CFBoundType = "auto" | "number";

interface CFGradientBound {
    type: CFBoundType;
    value?: string;
    color: string;
}

interface CFGradientSettings {
    formatStyle: "gradient";
    basedOnField: string; // leaf label
    emptyValues: CFEmptyValueHandling;
    min: CFGradientBound;
    max: CFGradientBound;
    useMid: boolean;
    midColor: string;
}

interface CFFieldValueSettings {
    formatStyle: "fieldValue";
    basedOnField: string; // leaf label
    excludedRowCodes?: string[];
}

type CFSurfaceSettings =
    | { formatStyle: "rules" }
    | CFGradientSettings
    | CFFieldValueSettings;

interface CFStyle {
    fontColor?: string;
    background?: string;
    iconName?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
}

interface CFRule {
    id: string;
    enabled: boolean;
    target: CFRuleTarget;
    scope?: CFApplyScope;
    formatProp?: CFFormatProp;
    rowMatchMode: CFMatchMode;
    rowMatchText: string;
    colMatchMode: CFMatchMode;
    colMatchText: string;
    conditionType: CFConditionType;
    operator: CFOperator;
    value1: string;
    value2: string;
    style: CFStyle;
}

interface CFConfig {
    version: 1 | 2;
    rules: CFRule[];
    surfaces?: Record<string, CFSurfaceSettings>;
}

type CTAggregation = "none" | "sum" | "avg" | "min" | "max" | "count";

interface CTValueMapping {
    field: string; // display label from bound fields (e.g. "Sales" or a category name)
    aggregation: CTAggregation;
}

interface CTParentRow {
    parentNo: string;
    parentName: string;
    values?: CTValueMapping[];
    format?: {
        fontFamily?: string;
        fontSize?: number;
        color?: string;
        bold?: boolean;
    };
}

interface CTChildRow {
    id: string;
    setParentNo: string;
    childName: string;
    // Optional: generate children automatically from a category field (e.g., City).
    // When set, the visual will create one child row per distinct value in this field.
    // childName is ignored for auto-generated children.
    childNameFromField?: string;
    // Optional: if set, only include values where this field equals setParentNo (e.g., Country == parentNo).
    parentMatchField?: string;
    values: CTValueMapping[];
    format?: {
        fontFamily?: string;
        fontSize?: number;
        color?: string;
        bold?: boolean;
    };
}

interface CustomTableConfig {
    version: 1;
    parents: CTParentRow[];
    children: CTChildRow[];

    // When true, allow multiple values per child row and expand columns by value field name (legacy behavior).
    // When false/undefined, treat each child row as a single selected value and do NOT expand columns by value.
    showValueNamesInColumns?: boolean;

    // Optional list of Column field keys (queryName or displayName) to hide in Custom Table.
    hiddenColumnFieldKeys?: string[];
}

function cryptoId(prefix: string = "r"): string {
    try {
        const bytes = new Uint32Array(2);
        window.crypto.getRandomValues(bytes);
        return `${prefix}_${bytes[0].toString(16)}${bytes[1].toString(16)}`;
    } catch {
        // Fallback: still avoid Math.random for lint; use time only.
        return `${prefix}_${Date.now().toString(16)}`;
    }
}

function addOption(select: HTMLSelectElement, value: string, label: string): void {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    select.appendChild(opt);
}

interface FancySelect {
    host: HTMLDivElement;
    select: HTMLSelectElement;
    setValue: (v: string) => void;
    getValue: () => string;
    syncFromSelect: () => void;
}

function createFancySelect(hostClassName: string = "fm-cf-select"): FancySelect {
    const host = document.createElement("div");
    host.className = `fm-dd ${hostClassName}`;

    const button = document.createElement("button");
    button.type = "button";
    button.className = "fm-dd-btn";
    button.setAttribute("aria-haspopup", "listbox");
    button.setAttribute("aria-expanded", "false");

    const menu = document.createElement("div");
    menu.className = "fm-dd-menu";
    menu.hidden = true;

    const select = document.createElement("select");
    select.className = "fm-dd-native";
    select.tabIndex = -1;
    select.setAttribute("aria-hidden", "true");

    host.appendChild(button);
    host.appendChild(menu);
    host.appendChild(select);

    let isOpen = false;
    let suppressDocHandler = false;

    const close = () => {
        if (!isOpen) return;
        isOpen = false;
        menu.hidden = true;
        host.classList.remove("fm-dd-open");
        button.setAttribute("aria-expanded", "false");
        document.removeEventListener("mousedown", onDocMouseDown, true);
        document.removeEventListener("keydown", onDocKeyDown, true);
    };

    const open = () => {
        if (isOpen) return;
        isOpen = true;
        syncFromSelect();
        menu.hidden = false;
        host.classList.add("fm-dd-open");
        button.setAttribute("aria-expanded", "true");
        suppressDocHandler = true;
        document.addEventListener("mousedown", onDocMouseDown, true);
        document.addEventListener("keydown", onDocKeyDown, true);
        setTimeout(() => { suppressDocHandler = false; }, 0);
    };

    const toggle = () => {
        if (isOpen) close();
        else open();
    };

    const onDocMouseDown = (ev: MouseEvent) => {
        if (suppressDocHandler) return;
        const target = ev.target as any;
        if (!host.contains(target)) close();
    };

    const onDocKeyDown = (ev: KeyboardEvent) => {
        if (!isOpen) return;
        if (ev.key === "Escape") {
            ev.preventDefault();
            close();
        }
    };

    const syncLabel = () => {
        const current = select.value;
        const match = Array.from(select.options).find(o => o.value === current);
        button.textContent = match?.textContent || "";
        button.disabled = !!select.disabled;
        host.classList.toggle("fm-dd-disabled", !!select.disabled);
    };

    const syncFromSelect = () => {
        syncLabel();
        menu.textContent = "";
        const opts = Array.from(select.options);
        for (const opt of opts) {
            const item = document.createElement("button");
            item.type = "button";
            item.className = "fm-dd-item";
            item.textContent = opt.textContent || opt.value;
            item.disabled = opt.disabled;
            if (opt.value === select.value) item.classList.add("fm-dd-item-selected");
            item.addEventListener("click", () => {
                if (opt.disabled) return;
                if (select.value !== opt.value) {
                    select.value = opt.value;
                    select.dispatchEvent(new Event("change"));
                }
                close();
            });
            menu.appendChild(item);
        }
    };

    button.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (select.disabled) return;
        toggle();
    });

    // Keep button label in sync when value changes programmatically.
    select.addEventListener("change", () => {
        syncLabel();
    });

    return {
        host,
        select,
        setValue: (v: string) => {
            select.value = v;
            syncLabel();
        },
        getValue: () => select.value,
        syncFromSelect
    };
}

type TokenType = "number" | "string" | "ident" | "ref" | "op" | "comma" | "lparen" | "rparen" | "eof";

interface Token {
    type: TokenType;
    value?: string;
}

class ExprTokenizer {
    private readonly text: string;
    private index: number = 0;

    constructor(text: string) {
        this.text = text ?? "";
    }

    public next(): Token {
        this.skipWs();
        if (this.index >= this.text.length) {
            return { type: "eof" };
        }

        const ch = this.text[this.index];

        if (ch === ",") {
            this.index++;
            return { type: "comma" };
        }
        if (ch === "(") {
            this.index++;
            return { type: "lparen" };
        }
        if (ch === ")") {
            this.index++;
            return { type: "rparen" };
        }

        if (ch === "[") {
            const end = this.text.indexOf("]", this.index + 1);
            if (end === -1) {
                const rest = this.text.slice(this.index + 1);
                this.index = this.text.length;
                return { type: "ref", value: rest.trim() };
            }
            const inside = this.text.slice(this.index + 1, end);
            this.index = end + 1;
            return { type: "ref", value: inside.trim() };
        }

        if (ch === "\"" || ch === "'") {
            const quote = ch;
            this.index++;
            let out = "";
            while (this.index < this.text.length) {
                const c = this.text[this.index++];
                if (c === quote) {
                    break;
                }
                if (c === "\\" && this.index < this.text.length) {
                    const n = this.text[this.index++];
                    out += n;
                    continue;
                }
                out += c;
            }
            return { type: "string", value: out };
        }

        const two = this.text.slice(this.index, this.index + 2);
        const ops2 = [">=", "<=", "==", "!=", "&&", "||"];
        if (ops2.includes(two)) {
            this.index += 2;
            return { type: "op", value: two };
        }

        const oneOps = ["+", "-", "*", "/", "^", ">", "<", "!"];
        if (oneOps.includes(ch)) {
            this.index++;
            return { type: "op", value: ch };
        }

        if (this.isDigit(ch) || (ch === "." && this.isDigit(this.peek()))) {
            const start = this.index;
            this.index++;
            while (this.index < this.text.length) {
                const c = this.text[this.index];
                if (this.isDigit(c) || c === ".") {
                    this.index++;
                    continue;
                }
                break;
            }
            return { type: "number", value: this.text.slice(start, this.index) };
        }

        if (this.isIdentStart(ch)) {
            const start = this.index;
            this.index++;
            while (this.index < this.text.length) {
                const c = this.text[this.index];
                if (this.isIdentPart(c)) {
                    this.index++;
                    continue;
                }
                break;
            }
            const raw = this.text.slice(start, this.index);
            const upper = raw.toUpperCase();
            if (upper === "AND") return { type: "op", value: "AND" };
            if (upper === "OR") return { type: "op", value: "OR" };
            if (upper === "NOT") return { type: "op", value: "NOT" };
            return { type: "ident", value: raw };
        }

        this.index++;
        return { type: "op", value: ch };
    }

    private skipWs(): void {
        while (this.index < this.text.length && /\s/.test(this.text[this.index])) {
            this.index++;
        }
    }

    private peek(): string {
        if (this.index + 1 >= this.text.length) return "\0";
        return this.text[this.index + 1];
    }

    private isDigit(c: string): boolean {
        return c >= "0" && c <= "9";
    }

    private isIdentStart(c: string): boolean {
        return /[A-Za-z_]/.test(c);
    }

    private isIdentPart(c: string): boolean {
        return /[A-Za-z0-9_]/.test(c);
    }
}

class ExprParser {
    private tokenizer: ExprTokenizer;
    private lookahead: Token;

    constructor(text: string) {
        this.tokenizer = new ExprTokenizer(text);
        this.lookahead = this.tokenizer.next();
    }

    public parse(): any {
        return this.parseOr();
    }

    private eat(type: TokenType, value?: string): Token {
        const t = this.lookahead;
        if (t.type !== type) {
            throw new Error(`Unexpected token: ${t.type}`);
        }
        if (value && (t.value ?? "") !== value) {
            throw new Error(`Unexpected token value: ${t.value}`);
        }
        this.lookahead = this.tokenizer.next();
        return t;
    }

    private match(type: TokenType, value?: string): boolean {
        if (this.lookahead.type !== type) return false;
        if (value && (this.lookahead.value ?? "") !== value) return false;
        return true;
    }

    private parseOr(): any {
        let left = this.parseAnd();
        while (this.match("op", "OR") || this.match("op", "||")) {
            const op = this.eat("op").value as string;
            const right = this.parseAnd();
            left = { type: "bin", op, left, right };
        }
        return left;
    }

    private parseAnd(): any {
        let left = this.parseEquality();
        while (this.match("op", "AND") || this.match("op", "&&")) {
            const op = this.eat("op").value as string;
            const right = this.parseEquality();
            left = { type: "bin", op, left, right };
        }
        return left;
    }

    private parseEquality(): any {
        let left = this.parseComparison();
        while (this.match("op", "==") || this.match("op", "!=")) {
            const op = this.eat("op").value as string;
            const right = this.parseComparison();
            left = { type: "bin", op, left, right };
        }
        return left;
    }

    private parseComparison(): any {
        let left = this.parseTerm();
        while (this.match("op", ">") || this.match("op", "<") || this.match("op", ">=") || this.match("op", "<=")) {
            const op = this.eat("op").value as string;
            const right = this.parseTerm();
            left = { type: "bin", op, left, right };
        }
        return left;
    }

    private parseTerm(): any {
        let left = this.parseFactor();
        while (this.match("op", "+") || this.match("op", "-")) {
            const op = this.eat("op").value as string;
            const right = this.parseFactor();
            left = { type: "bin", op, left, right };
        }
        return left;
    }

    private parseFactor(): any {
        let left = this.parsePower();
        while (this.match("op", "*") || this.match("op", "/")) {
            const op = this.eat("op").value as string;
            const right = this.parsePower();
            left = { type: "bin", op, left, right };
        }
        return left;
    }

    private parsePower(): any {
        let left = this.parseUnary();
        while (this.match("op", "^")) {
            const op = this.eat("op").value as string;
            const right = this.parseUnary();
            left = { type: "bin", op, left, right };
        }
        return left;
    }

    private parseUnary(): any {
        if (this.match("op", "+") || this.match("op", "-") || this.match("op", "NOT") || this.match("op", "!")) {
            const op = this.eat("op").value as string;
            const expr = this.parseUnary();
            return { type: "un", op, expr };
        }
        return this.parsePrimary();
    }

    private parsePrimary(): any {
        if (this.match("number")) {
            const v = parseFloat(this.eat("number").value as string);
            return { type: "num", value: v };
        }
        if (this.match("string")) {
            const s = this.eat("string").value as string;
            return { type: "str", value: s };
        }
        if (this.match("ref")) {
            const r = this.eat("ref").value as string;
            return { type: "ref", value: r };
        }
        if (this.match("ident")) {
            const ident = this.eat("ident").value as string;
            if (this.match("lparen")) {
                this.eat("lparen");
                const args: any[] = [];
                if (!this.match("rparen")) {
                    args.push(this.parseOr());
                    while (this.match("comma")) {
                        this.eat("comma");
                        args.push(this.parseOr());
                    }
                }
                this.eat("rparen");
                return { type: "call", name: ident, args };
            }
            return { type: "ident", value: ident };
        }
        if (this.match("lparen")) {
            this.eat("lparen");
            const expr = this.parseOr();
            this.eat("rparen");
            return expr;
        }
        if (this.match("eof")) {
            return { type: "num", value: 0 };
        }
        throw new Error("Invalid expression");
    }
}

function toNumber(v: any): number {
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (v == null) return NaN;
    const n = Number(v);
    return Number.isFinite(n) ? n : NaN;
}

function toBool(v: any): boolean {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0 && !Number.isNaN(v);
    if (typeof v === "string") return v.length > 0;
    return !!v;
}

export class Visual implements IVisual {
    private host: powerbi.extensibility.visual.IVisualHost;
    private selectionManager: powerbi.extensibility.ISelectionManager;
    private tooltipService: any = null;
    private activeSelectionKeys: Set<string> = new Set();
    private formattingSettingsService: FormattingSettingsService;
    private formattingSettings: VisualFormattingSettingsModel;
    private root: HTMLDivElement;
    private collapsed: Set<string> = new Set();

    private cfRulesJson: string = "";
    private cfEditorDraft: CFConfig | null = null;
    private cfAvailableFields: string[] = [];
    private cfAvailableMeasureFields: string[] = [];

    private customTableJson: string = "";
    private customTableEditorDraft: CustomTableConfig | null = null;
    private customTableAvailableFields: string[] = [];
    private customTableAvailableCategoryFields: string[] = [];
    private customTableFieldIsMeasure: Map<string, boolean> = new Map();
    private customTableBoundColumnFields: Array<{ key: string; label: string }> = [];
    private lastDataView: DataView | null = null;
    private lastViewport: powerbi.IViewport | null = null;
    // Landing overlay (same pattern as advancedPieChart)
    private landingEl: HTMLDivElement | null = null;
    private pendingCustomTableJson: string | null = null;
    private pendingCustomTableSetAt: number = 0;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        try {
            this.tooltipService = (this.host as any)?.tooltipService ?? null;
        } catch {
            this.tooltipService = null;
        }
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();

        this.root = document.createElement("div");
        this.root.className = "fm-container";
        options.element.appendChild(this.root);

        // Keep a local view of selection so we can render selected row styling.
        try {
            const smAny: any = this.selectionManager as any;
            if (smAny && typeof smAny.registerOnSelectCallback === "function") {
                smAny.registerOnSelectCallback((ids: any[]) => {
                    const keys = new Set<string>();
                    if (Array.isArray(ids)) {
                        for (const id of ids) {
                            try {
                                const k = typeof id?.getKey === "function" ? id.getKey() : "";
                                if (k) keys.add(k);
                            } catch {
                                // ignore
                            }
                        }
                    }
                    this.activeSelectionKeys = keys;
                    // Re-render to apply selection highlight.
                    if (this.lastDataView && this.lastViewport) {
                        const model = this.buildModel(this.lastDataView);
                        this.render(model, this.lastViewport);
                    }
                });
            }
        } catch {
            // ignore selection callback wiring errors
        }

        this.ensureLandingOverlay();
    }

    private showTooltip(ev: MouseEvent, items: powerbi.extensibility.VisualTooltipDataItem[], selectionId?: powerbi.extensibility.ISelectionId | null): void {
        const svc: any = this.tooltipService;
        if (!svc || !Array.isArray(items) || items.length === 0) return;
        try {
            svc.show({
                coordinates: [ev.clientX, ev.clientY],
                dataItems: items,
                identities: selectionId ? [selectionId] : undefined,
                isTouchEvent: false
            });
        } catch {
            // ignore tooltip errors
        }
    }

    private hideTooltip(immediately: boolean = true): void {
        const svc: any = this.tooltipService;
        if (!svc) return;
        try {
            svc.hide({ isTouchEvent: false, immediately });
        } catch {
            // ignore tooltip errors
        }
    }

    private applySelectionFromIds(ids: any[]): void {
        const keys = new Set<string>();
        if (Array.isArray(ids)) {
            for (const id of ids) {
                try {
                    const k = typeof id?.getKey === "function" ? id.getKey() : "";
                    if (k) keys.add(k);
                } catch {
                    // ignore
                }
            }
        }
        this.activeSelectionKeys = keys;
    }

    private selectRow(selectionId: powerbi.extensibility.ISelectionId | null, multiSelect: boolean = false): void {
        if (!selectionId) return;
        try {
            const p: any = (this.selectionManager as any).select(selectionId, multiSelect);
            if (p && typeof p.then === "function") {
                p.then((ids: any[]) => {
                    this.applySelectionFromIds(ids);
                    if (this.lastDataView && this.lastViewport) {
                        const model = this.buildModel(this.lastDataView);
                        this.render(model, this.lastViewport);
                    }
                }).catch(() => {
                    // ignore
                });
            }
        } catch {
            // ignore selection errors
        }
    }

    public update(options: VisualUpdateOptions): void {
        const dataView = options.dataViews && options.dataViews[0];
        if (!dataView) {
            this.lastViewport = options.viewport;
            this.hideLanding();
            this.renderEmptyState(options.viewport, { reason: "noDataView", missingRow: true, missingValues: true });
            return;
        }
        this.lastDataView = dataView;
        this.lastViewport = options.viewport;

        this.ensureLandingOverlay();

        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);

        // Persisted rules (not surfaced as a formatting pane slice).
        this.cfRulesJson = this.getObjectValue<string>(dataView, "conditionalFormatting", "rulesJson", "");

        // Persisted custom table definition (not surfaced as a formatting pane slice).
        // NOTE: persistProperties round-trip can lag; do not overwrite locally-applied JSON with stale/empty values.
        const dvCtJson = this.getObjectValue<string>(dataView, "customTable", "tableJson", "");
        if (this.pendingCustomTableJson) {
            if (dvCtJson === this.pendingCustomTableJson) {
                // Round-trip complete.
                this.customTableJson = dvCtJson;
                this.pendingCustomTableJson = null;
                this.pendingCustomTableSetAt = 0;
            } else {
                const ageMs = Date.now() - this.pendingCustomTableSetAt;
                if (ageMs < 10_000) {
                    // Keep using the locally applied JSON until the host catches up.
                    this.customTableJson = this.pendingCustomTableJson;
                } else {
                    // Give up after a while; host didn't reflect it.
                    this.pendingCustomTableJson = null;
                    this.pendingCustomTableSetAt = 0;
                    this.customTableJson = dvCtJson;
                }
            }
        } else {
            // Prefer non-empty DataView value; otherwise keep existing local value.
            const dvTrim = String(dvCtJson || "").trim();
            if (dvTrim) this.customTableJson = dvCtJson;
            else if (!String(this.customTableJson || "").trim()) this.customTableJson = dvCtJson;
        }

        // Hint behavior (same approach as advancedPieChart):
        // Show landing message when required roles are missing; hide as soon as ready.
        const ctCard: any = (this.formattingSettings as any).customTable;
        const customTableEnabled = (ctCard?.enabled?.value as boolean | undefined) ?? false;

        // If nothing is bound at all, Power BI may still provide a DataView shell.
        // Use categorical presence (not metadata fallback) to decide the landing state.
        const categorical: any = (dataView as any)?.categorical;
        const catsNow: any[] = (categorical?.categories as any[]) || [];
        const valsNow: any[] = (categorical?.values as any[]) || [];
        const hasAnyCategoricalBound = (catsNow.length > 0) || (valsNow.length > 0);

        if (!hasAnyCategoricalBound) {
            this.hideLanding();
            this.renderEmptyState(options.viewport, { reason: "missingBindings", missingRow: true, missingValues: true });
            return;
        }

        if (customTableEnabled) {
            // Custom Table mode: require at least one measure bound.
            // UX rule: once the user starts binding anything, hide the onboarding (even if not fully ready).
            const hasAnyValue = valsNow.some(v => {
                const roles: any = v?.source?.roles || {};
                return !!roles.values || !!roles.customTableValue || !!roles.formatBy;
            });
            if (!hasAnyValue) {
                this.hideLanding();
                this.clearRoot();
                return;
            }
        } else {
            const hasRow = catsNow.some(c => !!c?.source?.roles?.category);
            const hasAnyValue = valsNow.some(v => {
                const roles: any = v?.source?.roles || {};
                return !!roles.values || !!roles.customTableValue || !!roles.formatBy;
            });

            // UX rule: as soon as *any* field is added (any role), hide the onboarding.
            // If bindings are incomplete, render nothing (avoid blocking the canvas with the guide).
            if (!hasRow || !hasAnyValue) {
                this.hideLanding();
                this.clearRoot();
                return;
            }
        }

        this.hideLanding();

        const model = this.buildModel(dataView);

        // Populate available fields for "Based on field" and Custom Table value pickers from the actual bound metadata.
        // NOTE: Custom visuals can only list fields that are bound to the visual.
        const out = new Set<string>();
        const outMeasures = new Set<string>();

        // Custom Table editor fields:
        // Prefer *dedicated* roles (customTableValue/customTableValueField). If user hasn't bound anything there,
        // fall back to the general bound-field list for convenience.
        const ctRoleOut = new Set<string>();
        const ctRoleIsMeasure = new Map<string, boolean>();
        const ctFallbackOut = new Set<string>();
        const ctFallbackIsMeasure = new Map<string, boolean>();

        // 1) Prefer DataView metadata columns (most reliable across shapes)
        const metaCols: any[] = ((dataView as any)?.metadata?.columns as any[]) || [];
        for (const c of metaCols) {
            const name = String((c?.displayName || c?.queryName || "")).trim();
            if (!name) continue;
            const roles: any = c?.roles || {};
            // Only include measures / value roles to avoid flooding with row/column categories.
            const isMeasure = !!c?.isMeasure;
            const isValueRole = !!roles?.values || !!roles?.formatBy || !!roles?.tooltips;
            if (!isMeasure && !isValueRole) continue;
            out.add(name);

            if (isMeasure || !!roles?.values || !!roles?.formatBy) outMeasures.add(name);

            // Custom table fallback uses the same labels.
            const fallbackLabel = name;
            ctFallbackOut.add(fallbackLabel);
            ctFallbackIsMeasure.set(fallbackLabel, true);

            // Dedicated Custom Table values role
            if (roles.customTableValue) {
                const ctLabel = name;
                ctRoleOut.add(ctLabel);
                ctRoleIsMeasure.set(ctLabel, true);
            }
        }

        // 2) Also collect from categorical values (covers cases where isMeasure isn't set)
        const dvValues: any[] = ((dataView as any)?.categorical?.values as any[]) || [];
        for (const v of dvValues) {
            const src = v?.source;
            const name = String((src?.displayName || src?.queryName || "")).trim();
            if (!name) continue;
            const roles: any = src?.roles || {};
            out.add(name);
            outMeasures.add(name);

            const fallbackLabel = name;
            ctFallbackOut.add(fallbackLabel);
            ctFallbackIsMeasure.set(fallbackLabel, true);

            if (roles.customTableValue) {
                const ctLabel = name;
                ctRoleOut.add(ctLabel);
                ctRoleIsMeasure.set(ctLabel, true);
            }
        }

        // 2b) Include bound categorical fields (Row/Column/etc.) so users can see columns too.
        // Note: Field value for *cell values* is typically driven by measures, but listing bound categories helps parity with Power BI UX.
        const dvCats: any[] = ((dataView as any)?.categorical?.categories as any[]) || [];
        const catFieldOut = new Set<string>();
        // Bound Column fields (for Custom Table editor column-level hide/show).
        const colFields: Array<{ key: string; label: string }> = [];
        const colFieldKeySet = new Set<string>();
        for (const c of dvCats) {
            const src = c?.source;
            const roles: any = src?.roles || {};
            if (!roles.column) continue;
            const key = String((src?.queryName || src?.displayName || "")).trim();
            const label = String((src?.displayName || src?.queryName || "")).trim();
            if (!key || colFieldKeySet.has(key)) continue;
            colFieldKeySet.add(key);
            colFields.push({ key, label: label || key });
        }
        this.customTableBoundColumnFields = colFields;
        for (const c of dvCats) {
            const src = c?.source;
            const name = String((src?.displayName || src?.queryName || "")).trim();
            if (name) out.add(name);

            // Track category field labels for auto-child generation.
            if (name) catFieldOut.add(name);

            // Custom table fallback supports picking bound categories too (with aggregation).
            if (name) {
                ctFallbackOut.add(name);
                if (!ctFallbackIsMeasure.has(name)) ctFallbackIsMeasure.set(name, false);
            }

            // Dedicated Custom Table value fields role
            const roles: any = src?.roles || {};
            if (name && roles.customTableValueField) {
                ctRoleOut.add(name);
                if (!ctRoleIsMeasure.has(name)) ctRoleIsMeasure.set(name, false);
            }
        }

        this.customTableAvailableCategoryFields = Array.from(catFieldOut.values()).sort((a, b) => a.localeCompare(b));

        // 3) Finally, include the actual leaf labels used in the rendered model.
        for (const c of model.columns) {
            const ll = (c.leafLabel || "").trim();
            if (ll) out.add(ll);
        }

        this.cfAvailableFields = Array.from(out.values()).sort((a, b) => a.localeCompare(b));
        this.cfAvailableMeasureFields = Array.from(outMeasures.values()).sort((a, b) => a.localeCompare(b));

        if (ctRoleOut.size > 0) {
            this.customTableAvailableFields = Array.from(ctRoleOut.values()).sort((a, b) => a.localeCompare(b));
            this.customTableFieldIsMeasure = ctRoleIsMeasure;
        } else {
            this.customTableAvailableFields = Array.from(ctFallbackOut.values()).sort((a, b) => a.localeCompare(b));
            this.customTableFieldIsMeasure = ctFallbackIsMeasure;
        }

        this.render(model, options.viewport);
    }

    private ensureLandingOverlay(): void {
        if (this.landingEl && this.landingEl.parentElement === this.root) return;
        const existing = this.root.querySelector<HTMLDivElement>(".fm-landing");
        if (existing) {
            this.landingEl = existing;
            return;
        }
        const el = document.createElement("div");
        el.className = "fm-landing";
        (el as any).dataset.fmPersist = "1";
        el.style.display = "none";
        this.root.appendChild(el);
        this.landingEl = el;
    }

    private showLanding(message: string): void {
        this.ensureLandingOverlay();
        if (!this.landingEl) return;
        this.landingEl.textContent = message;
        this.landingEl.style.display = "block";
    }

    private hideLanding(): void {
        if (!this.landingEl) return;
        this.landingEl.style.display = "none";
        this.landingEl.textContent = "";
    }

    private hasAnyBoundField(dataView: DataView): boolean {
        const cats: any[] = ((dataView as any)?.categorical?.categories as any[]) || [];
        const vals: any[] = ((dataView as any)?.categorical?.values as any[]) || [];

        if (cats.length > 0 || vals.length > 0) return true;

        const metaCols: any[] = ((dataView as any)?.metadata?.columns as any[]) || [];
        return metaCols.some(c => {
            const roles: any = c?.roles || {};
            return !!roles.column || !!roles.category || !!roles.values || !!roles.formatBy || !!roles.formatByField || !!roles.customTableValue || !!roles.customTableValueField;
        });
    }

    private getBindingStatus(dataView: DataView): { hasRow: boolean; hasAnyValue: boolean } {
        const cats: any[] = ((dataView as any)?.categorical?.categories as any[]) || [];
        const vals: any[] = ((dataView as any)?.categorical?.values as any[]) || [];

        const hasRow = cats.some(c => !!c?.source?.roles?.category);
        const hasAnyValue = vals.some(v => {
            const roles: any = v?.source?.roles || {};
            return !!roles.values || !!roles.customTableValue || !!roles.formatBy;
        });

        if (hasRow || hasAnyValue) return { hasRow, hasAnyValue };

        // Fallback: some shapes expose roles only via metadata columns.
        const metaCols: any[] = ((dataView as any)?.metadata?.columns as any[]) || [];
        const metaHasRow = metaCols.some(c => !!c?.roles?.category);
        const metaHasAnyValue = metaCols.some(c => {
            const roles: any = c?.roles || {};
            return !!roles.values || !!roles.customTableValue || !!roles.formatBy;
        });
        return { hasRow: metaHasRow, hasAnyValue: metaHasAnyValue };
    }

    private getObjectValue<T>(dataView: DataView, objectName: string, propertyName: string, defaultValue: T): T {
        const objects: any = (dataView as any)?.metadata?.objects;
        const obj = objects && (objects as any)[objectName];
        const val = obj && (obj as any)[propertyName];
        return (val as T) ?? defaultValue;
    }

    private persistObjectProperties(objectName: string, properties: Record<string, any>): void {
        try {
            (this.host as any).persistProperties({
                merge: [
                    {
                        objectName,
                        selector: null,
                        properties
                    }
                ]
            });
        } catch {
            // ignore
        }
    }

    private safeParseCfConfig(json: string): CFConfig {
        try {
            const parsed = JSON.parse(json || "") as Partial<CFConfig>;
            const v = (parsed as any)?.version;
            const versionOk = v === 1 || v === 2;
            if (parsed && versionOk && Array.isArray((parsed as any).rules)) {
                return {
                    version: (v as any) ?? 1,
                    rules: (parsed as any).rules
                        .filter(r => !!r)
                        .map((r: any) => this.normalizeCfRule(r))
                    ,
                    surfaces: (parsed as any).surfaces && typeof (parsed as any).surfaces === "object"
                        ? (parsed as any).surfaces
                        : undefined
                };
            }
        } catch {
            // ignore
        }
        return { version: 1, rules: [], surfaces: undefined };
    }

    private cfSurfaceKey(target: CFRuleTarget, formatProp: CFFormatProp): string {
        return `${target}:${formatProp}`;
    }

    private getSurfaceSettings(config: CFConfig, target: CFRuleTarget, formatProp: CFFormatProp): CFSurfaceSettings {
        const key = this.cfSurfaceKey(target, formatProp);
        const s = config.surfaces ? (config.surfaces as any)[key] : undefined;
        const fs = (s as any)?.formatStyle as CFFormatStyle | undefined;
        if (fs === "gradient" || fs === "fieldValue" || fs === "rules") {
            // Icons don't support gradient.
            if (formatProp === "icon" && fs === "gradient") return { formatStyle: "rules" };
            // Icons surface is only meaningful for value cells.
            if (fs === "fieldValue" && formatProp === "icon" && target !== "cell") return { formatStyle: "rules" };
            return s as any;
        }
        return { formatStyle: "rules" };
    }

    private setSurfaceSettings(draft: CFConfig, target: CFRuleTarget, formatProp: CFFormatProp, settings: CFSurfaceSettings): void {
        if (!draft.surfaces) draft.surfaces = {};
        const key = this.cfSurfaceKey(target, formatProp);
        (draft.surfaces as any)[key] = settings;
    }

    private parseSafeCssColor(v: any): string | null {
        const s = (v == null ? "" : String(v)).trim();
        if (!s) return null;
        if (/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(s)) return s;
        if (/^(rgb|rgba)\(\s*\d+\s*,\s*\d+\s*,\s*\d+(\s*,\s*(0(\.\d+)?|1(\.0+)?))?\s*\)$/.test(s)) return s;
        if (/^(hsl|hsla)\(\s*\d+\s*,\s*\d+%\s*,\s*\d+%(\s*,\s*(0(\.\d+)?|1(\.0+)?))?\s*\)$/.test(s)) return s;
        if (/^[a-zA-Z]+$/.test(s)) return s;
        return null;
    }

    private hexToRgb(hex: string): { r: number; g: number; b: number } | null {
        const h = (hex || "").trim();
        if (!/^#/.test(h)) return null;
        const raw = h.slice(1);
        if (raw.length === 3) {
            const r = parseInt(raw[0] + raw[0], 16);
            const g = parseInt(raw[1] + raw[1], 16);
            const b = parseInt(raw[2] + raw[2], 16);
            if ([r, g, b].some(x => Number.isNaN(x))) return null;
            return { r, g, b };
        }
        if (raw.length === 6 || raw.length === 8) {
            // If 8 digits are provided, ignore alpha for now.
            const rrggbb = raw.length === 8 ? raw.slice(0, 6) : raw;
            const r = parseInt(rrggbb.slice(0, 2), 16);
            const g = parseInt(rrggbb.slice(2, 4), 16);
            const b = parseInt(rrggbb.slice(4, 6), 16);
            if ([r, g, b].some(x => Number.isNaN(x))) return null;
            return { r, g, b };
        }
        return null;
    }

    private lerp(a: number, b: number, t: number): number {
        return a + (b - a) * t;
    }

    private clamp01(t: number): number {
        if (!Number.isFinite(t)) return 0;
        return Math.max(0, Math.min(1, t));
    }

    private lerpColorHex(a: string, b: string, t: number): string {
        const ca = this.hexToRgb(a);
        const cb = this.hexToRgb(b);
        if (!ca || !cb) return b;
        const tt = this.clamp01(t);
        const r = Math.round(this.lerp(ca.r, cb.r, tt));
        const g = Math.round(this.lerp(ca.g, cb.g, tt));
        const bl = Math.round(this.lerp(ca.b, cb.b, tt));
        const toHex2 = (n: number) => n.toString(16).padStart(2, "0");
        return `#${toHex2(r)}${toHex2(g)}${toHex2(bl)}`;
    }

    private normalizeCfRule(r: any): CFRule {
        const clean = (s: any) => (s == null ? "" : String(s));
        const cleanOptional = (s: any): string | undefined => {
            const v = (s == null ? "" : String(s)).trim();
            return v ? v : undefined;
        };
        const hasOwn = (obj: any, key: string): boolean => {
            try {
                return obj != null && Object.prototype.hasOwnProperty.call(obj, key);
            } catch {
                return false;
            }
        };
        const cleanMode = (m: any): CFMatchMode => (m === "equals" || m === "contains") ? m : "any";
        const cleanTarget = (t: any): CFRuleTarget => (t === "columnHeader" || t === "cell") ? t : "rowHeader";
        const cleanCond = (c: any): CFConditionType => (c === "text") ? "text" : "value";
        const cleanScope = (s: any): CFApplyScope => {
            return (s === "entireRow" || s === "entireColumn") ? s : "self";
        };
        const cleanFormatProp = (p: any): CFFormatProp | undefined => {
            return (p === "background" || p === "fontColor" || p === "icon") ? p : undefined;
        };
        const cleanOp = (o: any): CFOperator => {
            const ops: CFOperator[] = ["gt", "gte", "lt", "lte", "eq", "between", "contains"];
            return ops.includes(o) ? o : "gte";
        };

        const id = clean(r.id) || cryptoId("r");

        // IMPORTANT: Do not inject empty strings/false defaults into style.
        // Absent properties must remain absent so they don't overwrite other rules when merged.
        const style: CFStyle = {};
        const fontColor = cleanOptional(r.style?.fontColor);
        const background = cleanOptional(r.style?.background);
        const iconName = cleanOptional(r.style?.iconName);
        if (fontColor) style.fontColor = fontColor;
        if (background) style.background = background;
        if (iconName) style.iconName = iconName;
        if (hasOwn(r.style, "bold")) style.bold = r.style?.bold === true ? true : false;
        if (hasOwn(r.style, "italic")) style.italic = r.style?.italic === true ? true : false;
        if (hasOwn(r.style, "underline")) style.underline = r.style?.underline === true ? true : false;

        // Decide which format tab this rule belongs to.
        // Existing rules may not have this property; infer from style.
        const inferredFormatProp: CFFormatProp = (
            cleanFormatProp(r.formatProp)
            ?? (iconName ? "icon" : undefined)
            ?? (background && !fontColor ? "background" : undefined)
            ?? (fontColor && !background ? "fontColor" : undefined)
            ?? "background"
        );

        return {
            id,
            enabled: !!r.enabled,
            target: cleanTarget(r.target),
            scope: cleanScope(r.scope),
            formatProp: inferredFormatProp,
            rowMatchMode: cleanMode(r.rowMatchMode),
            rowMatchText: clean(r.rowMatchText),
            colMatchMode: cleanMode(r.colMatchMode),
            colMatchText: clean(r.colMatchText),
            conditionType: cleanCond(r.conditionType),
            operator: cleanOp(r.operator),
            value1: clean(r.value1),
            value2: clean(r.value2),
            style
        };
    }

    private applyCfStyle(el: HTMLElement, style: CFStyle): void {
        if (!style) return;
        if (style.fontColor) el.style.color = style.fontColor;
        if (style.background) el.style.background = style.background;
        if (style.bold != null) el.style.fontWeight = style.bold ? "700" : "";
        if (style.italic != null) el.style.fontStyle = style.italic ? "italic" : "";
        if (style.underline != null) el.style.textDecoration = style.underline ? "underline" : "";
    }

    private createIconElementFromSpec(spec: any): HTMLElement | null {
        const raw = (spec == null ? "" : String(spec)).trim();
        if (!raw) return null;

        const wrap = document.createElement("span");
        wrap.className = "fm-cell-icon";

        // Allow data URIs (SVG/PNG) for field-value icons.
        if (/^data:image\/(svg\+xml|png);/i.test(raw)) {
            // If it's an svg+xml;utf8, help with '#' encoding (Power BI often needs %23)
            const fixed = (/^data:image\/svg\+xml;utf8,/i.test(raw) && raw.includes("#"))
                ? raw.replace(/#/g, "%23")
                : raw;
            const img = document.createElement("img");
            img.alt = "";
            img.src = fixed;
            wrap.appendChild(img);
            return wrap;
        }

        const key = raw.toLowerCase().replace(/\s+/g, "");
        const resolved = (CF_ICON_ALIASES[key] as CFBuiltinIconKey | undefined)
            ?? (CF_BUILTIN_ICONS.some(i => i.key === (key as any)) ? (key as any as CFBuiltinIconKey) : undefined);
        if (!resolved) return null;

        const svgText = CF_ICON_SVG[resolved] || "";
        if (!svgText) return null;
        try {
            const doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
            const svg = doc.documentElement;
            if (svg && svg.tagName && svg.tagName.toLowerCase() === "svg") {
                wrap.appendChild(document.importNode(svg, true));
                return wrap;
            }
        } catch {
            // ignore
        }
        return null;
    }

    private evalCfRules(
        config: CFConfig,
        ctx: {
            target: CFRuleTarget;
            formatProp?: CFFormatProp;
            rowCode?: string;
            rowLabel?: string;
            colKey?: string;
            colText?: string;
            value?: number | null;
            text?: string;
        }
    ): CFStyle | null {
        const norm = (s: string) => (s || "").trim().toLowerCase();
        const matchText = (mode: CFMatchMode, needle: string, hay: string): boolean => {
            if (mode === "any") return true;
            const n = norm(needle);
            if (!n) return true;
            const h = norm(hay);
            if (mode === "equals") return h === n;
            return h.includes(n);
        };

        const parseNum = (s: string): number => {
            const n = Number((s || "").trim());
            return Number.isFinite(n) ? n : NaN;
        };

        let out: CFStyle | null = null;
        for (const rule of config.rules) {
            if (!rule.enabled) continue;
            if (rule.target !== ctx.target) continue;
            if (ctx.formatProp && (rule.formatProp || "background") !== ctx.formatProp) continue;

            const rowHay = (ctx.rowLabel || ctx.rowCode || "");
            const colHay = (ctx.colText || ctx.colKey || "");

            if (!matchText(rule.rowMatchMode, rule.rowMatchText, rowHay)) continue;
            if (!matchText(rule.colMatchMode, rule.colMatchText, colHay)) continue;

            let ok = true;
            if (rule.conditionType === "value") {
                const v = ctx.value;
                if (v == null || !Number.isFinite(v)) {
                    ok = false;
                } else {
                    const a = parseNum(rule.value1);
                    const b = parseNum(rule.value2);
                    switch (rule.operator) {
                        case "gt": ok = Number.isFinite(a) ? v > a : false; break;
                        case "gte": ok = Number.isFinite(a) ? v >= a : false; break;
                        case "lt": ok = Number.isFinite(a) ? v < a : false; break;
                        case "lte": ok = Number.isFinite(a) ? v <= a : false; break;
                        case "eq": ok = Number.isFinite(a) ? v === a : false; break;
                        case "between":
                            ok = Number.isFinite(a) && Number.isFinite(b) ? (v >= Math.min(a, b) && v <= Math.max(a, b)) : false;
                            break;
                        case "contains":
                            ok = false;
                            break;
                        default:
                            ok = false;
                    }
                }
            } else {
                const t = norm(ctx.text || "");
                const n = norm(rule.value1);
                switch (rule.operator) {
                    case "contains": ok = !!n ? t.includes(n) : false; break;
                    case "eq": ok = !!n ? t === n : false; break;
                    default:
                        ok = false;
                }
            }

            if (!ok) continue;

            out = { ...(out || {}), ...(rule.style || {}) };
        }
        return out;
    }

    private evalCfRulesForValueCell(
        config: CFConfig,
        ctx: {
            rowCode: string;
            rowLabel: string;
            colKey: string;
            colText: string;
            value: number | null;
            formatProp?: CFFormatProp;
        }
    ): CFStyle | null {
        const norm = (s: string) => (s || "").trim().toLowerCase();
        const matchText = (mode: CFMatchMode, needle: string, hay: string): boolean => {
            if (mode === "any") return true;
            const n = norm(needle);
            if (!n) return true;
            const h = norm(hay);
            if (mode === "equals") return h === n;
            return h.includes(n);
        };
        const parseNum = (s: string): number => {
            const n = Number((s || "").trim());
            return Number.isFinite(n) ? n : NaN;
        };

        let out: CFStyle | null = null;
        for (const rule of config.rules) {
            if (!rule.enabled) continue;
            if (ctx.formatProp && (rule.formatProp || "background") !== ctx.formatProp) continue;

            const rowHay = ctx.rowLabel || ctx.rowCode || "";
            const colHay = ctx.colText || ctx.colKey || "";

            // Apply scoped header rules to value cells when requested.
            if (rule.target === "rowHeader" && (rule.scope || "self") === "entireRow") {
                if (!matchText(rule.rowMatchMode, rule.rowMatchText, rowHay)) continue;
                if (!matchText(rule.colMatchMode, rule.colMatchText, colHay)) continue;

                let ok = true;
                if (rule.conditionType === "text") {
                    const t = norm(rowHay);
                    const n = norm(rule.value1);
                    switch (rule.operator) {
                        case "contains": ok = !!n ? t.includes(n) : false; break;
                        case "eq": ok = !!n ? t === n : false; break;
                        default: ok = false;
                    }
                } else {
                    ok = false;
                }
                if (!ok) continue;
                out = { ...(out || {}), ...(rule.style || {}) };
                continue;
            }

            if (rule.target === "columnHeader" && (rule.scope || "self") === "entireColumn") {
                if (!matchText(rule.rowMatchMode, rule.rowMatchText, rowHay)) continue;
                if (!matchText(rule.colMatchMode, rule.colMatchText, colHay)) continue;

                let ok = true;
                if (rule.conditionType === "text") {
                    const t = norm(colHay);
                    const n = norm(rule.value1);
                    switch (rule.operator) {
                        case "contains": ok = !!n ? t.includes(n) : false; break;
                        case "eq": ok = !!n ? t === n : false; break;
                        default: ok = false;
                    }
                } else {
                    ok = false;
                }
                if (!ok) continue;
                out = { ...(out || {}), ...(rule.style || {}) };
                continue;
            }

            // Regular value-cell rules.
            if (rule.target !== "cell") continue;
            if (!matchText(rule.rowMatchMode, rule.rowMatchText, rowHay)) continue;
            if (!matchText(rule.colMatchMode, rule.colMatchText, colHay)) continue;

            let ok = true;
            if (rule.conditionType === "value") {
                const v = ctx.value;
                if (v == null || !Number.isFinite(v)) {
                    ok = false;
                } else {
                    const a = parseNum(rule.value1);
                    const b = parseNum(rule.value2);
                    switch (rule.operator) {
                        case "gt": ok = Number.isFinite(a) ? v > a : false; break;
                        case "gte": ok = Number.isFinite(a) ? v >= a : false; break;
                        case "lt": ok = Number.isFinite(a) ? v < a : false; break;
                        case "lte": ok = Number.isFinite(a) ? v <= a : false; break;
                        case "eq": ok = Number.isFinite(a) ? v === a : false; break;
                        case "between":
                            ok = Number.isFinite(a) && Number.isFinite(b) ? (v >= Math.min(a, b) && v <= Math.max(a, b)) : false;
                            break;
                        default:
                            ok = false;
                    }
                }
            } else {
                ok = false;
            }
            if (!ok) continue;

            out = { ...(out || {}), ...(rule.style || {}) };
        }
        return out;
    }

    private ensureCfDraft(): CFConfig {
        if (this.cfEditorDraft) return this.cfEditorDraft;
        this.cfEditorDraft = this.safeParseCfConfig(this.cfRulesJson);
        return this.cfEditorDraft;
    }

    private closeCfEditor(): void {
        this.cfEditorDraft = null;
        // Turn off the formatting toggle so the pane doesn't stay ON.
        this.persistObjectProperties("conditionalFormatting", { showEditor: false });
        // Re-render using existing persisted rules.
        // (Power BI will call update again, but we keep a quick redraw for responsiveness.)
    }

    private safeParseCustomTableConfig(json: string): CustomTableConfig {
        const empty: CustomTableConfig = { version: 1, parents: [], children: [] };
        try {
            const parsed = JSON.parse(json || "") as Partial<CustomTableConfig>;
            if (!parsed || (parsed as any).version !== 1) return empty;
            const parentsRaw: any[] = Array.isArray((parsed as any).parents) ? ((parsed as any).parents as any[]) : [];
            const childrenRaw: any[] = Array.isArray((parsed as any).children) ? ((parsed as any).children as any[]) : [];

            const normAgg = (a: any): CTAggregation => {
                const s = String(a ?? "none").trim().toLowerCase();
                if (s === "sum" || s === "avg" || s === "min" || s === "max" || s === "count" || s === "none") return s as any;
                return "none";
            };

            const normFont = (s: any): string | undefined => {
                const v = String(s ?? "").trim();
                if (!v) return undefined;
                // Keep it simple: allow a short whitelist so we don't store arbitrary css.
                const ok = ["Segoe UI", "Arial", "Calibri"].includes(v);
                return ok ? v : undefined;
            };

            const normFontSize = (n: any): number | undefined => {
                const v = Number(n);
                if (!Number.isFinite(v)) return undefined;
                const vv = Math.max(6, Math.min(72, Math.round(v)));
                return vv;
            };

            const normColor = (c: any): string | undefined => {
                const s = String(c ?? "").trim();
                if (!s) return undefined;
                if (/^#([0-9a-fA-F]{6})$/.test(s)) return s;
                if (/^#([0-9a-fA-F]{3})$/.test(s)) return s;
                return undefined;
            };

            const normFormat = (f: any): any => {
                if (!f || typeof f !== "object") return undefined;
                const fontFamily = normFont((f as any).fontFamily);
                const fontSize = normFontSize((f as any).fontSize);
                const color = normColor((f as any).color);
                const bold = (f as any).bold === true ? true : undefined;
                if (!fontFamily && !fontSize && !color && !bold) return undefined;
                return { fontFamily, fontSize, color, bold };
            };

            const parents: CTParentRow[] = parentsRaw
                .map(p => ({
                    parentNo: String(p?.parentNo ?? "").trim(),
                    parentName: String(p?.parentName ?? "").trim(),
                    values: (() => {
                        const valuesRaw: any[] = Array.isArray((p as any)?.values) ? ((p as any).values as any[]) : [];
                        return valuesRaw
                            .map(v => ({
                                field: String(v?.field ?? "").trim(),
                                aggregation: normAgg(v?.aggregation)
                            }))
                            .filter(v => !!v.field);
                    })(),
                    format: normFormat(p?.format)
                }))
                .filter(p => !!p.parentNo);

            const children: CTChildRow[] = childrenRaw
                .map(c => {
                    const id = String(c?.id ?? "").trim() || cryptoId("ct");
                    const setParentNo = String(c?.setParentNo ?? "").trim();
                    const childName = String(c?.childName ?? "").trim();
                    const childNameFromField = String((c as any)?.childNameFromField ?? "").trim();
                    const parentMatchField = String((c as any)?.parentMatchField ?? "").trim();
                    const valuesRaw: any[] = Array.isArray(c?.values) ? (c.values as any[]) : [];
                    const values: CTValueMapping[] = valuesRaw
                        .map(v => ({
                            field: String(v?.field ?? "").trim(),
                            aggregation: normAgg(v?.aggregation)
                        }))
                        .filter(v => !!v.field);
                    return { id, setParentNo, childName, childNameFromField: childNameFromField || undefined, parentMatchField: parentMatchField || undefined, values, format: normFormat(c?.format) };
                })
                .filter(c => !!c.childName || !!c.childNameFromField);

            const showValueNamesInColumns = (parsed as any)?.showValueNamesInColumns === true;
            const hiddenColumnFieldKeys = Array.isArray((parsed as any)?.hiddenColumnFieldKeys)
                ? ((parsed as any).hiddenColumnFieldKeys as any[]).map(x => String(x ?? "").trim()).filter(x => !!x)
                : [];

            return { version: 1, parents, children, showValueNamesInColumns, hiddenColumnFieldKeys };
        } catch {
            return empty;
        }
    }

    private ensureCustomTableDraft(): CustomTableConfig {
        if (this.customTableEditorDraft) return this.customTableEditorDraft;
        this.customTableEditorDraft = this.safeParseCustomTableConfig(this.customTableJson);
        return this.customTableEditorDraft;
    }

    private closeCustomTableEditor(): void {
        this.customTableEditorDraft = null;
        this.persistObjectProperties("customTable", { showEditor: false });
    }

    private openCustomTableEditor(): void {
        this.ensureCustomTableDraft();
        this.persistObjectProperties("customTable", { showEditor: true });
    }

    private renderCustomTableEditor(): void {
        const ctCard: any = (this.formattingSettings as any).customTable;
        const enabled = (ctCard?.enabled?.value as boolean | undefined) ?? false;
        const showEditor = (ctCard?.showEditor?.value as boolean | undefined) ?? false;
        if (!enabled || !showEditor) return;

        const draft = this.ensureCustomTableDraft();

        const overlay = document.createElement("div");
        overlay.className = "fm-cf-overlay";

        const panel = document.createElement("div");
        panel.className = "fm-cf-panel fm-ct-panel";
        overlay.appendChild(panel);

        const header = document.createElement("div");
        header.className = "fm-cf-header";
        panel.appendChild(header);

        const title = document.createElement("div");
        title.className = "fm-cf-title";
        title.textContent = "Custom table editor";
        header.appendChild(title);

        const note = document.createElement("div");
        note.className = "fm-cf-style-note";
        note.textContent = "Define a custom Parent/Child hierarchy and pick measures/columns for values. Aggregation is disabled for measures (Power BI handles measures automatically).";
        header.appendChild(note);

        if (draft.showValueNamesInColumns == null) draft.showValueNamesInColumns = false;
        if (!Array.isArray(draft.hiddenColumnFieldKeys)) draft.hiddenColumnFieldKeys = [];

        let rerenderChildrenFn: (() => void) | null = null;

        // Global Custom Table options
        const optionsWrap = document.createElement("div");
        optionsWrap.className = "fm-ct-section";
        panel.appendChild(optionsWrap);

        const optionsTitle = document.createElement("div");
        optionsTitle.className = "fm-cf-style-title";
        optionsTitle.textContent = "Options";
        optionsWrap.appendChild(optionsTitle);

        const showNamesLabel = document.createElement("label");
        showNamesLabel.className = "fm-ct-bold";
        const showNamesToggle = document.createElement("input");
        showNamesToggle.type = "checkbox";
        showNamesToggle.checked = draft.showValueNamesInColumns === true;
        showNamesToggle.onchange = () => {
            draft.showValueNamesInColumns = showNamesToggle.checked;
            if (rerenderChildrenFn) rerenderChildrenFn();
        };
        showNamesLabel.appendChild(showNamesToggle);
        const showNamesTxt = document.createElement("span");
        showNamesTxt.textContent = "Show value name in column headers (enables multiple values per child)";
        showNamesLabel.appendChild(showNamesTxt);
        optionsWrap.appendChild(showNamesLabel);

        const boundCols = Array.isArray(this.customTableBoundColumnFields) ? this.customTableBoundColumnFields : [];
        if (boundCols.length > 0) {
            const colsTitle = document.createElement("div");
            colsTitle.className = "fm-cf-style-note";
            colsTitle.textContent = "Column fields visible in Custom Table:";
            optionsWrap.appendChild(colsTitle);

            const hiddenSet = new Set((draft.hiddenColumnFieldKeys || []).map(x => String(x || "").trim()).filter(x => !!x));

            for (const f of boundCols) {
                const lab = document.createElement("label");
                lab.className = "fm-ct-bold";
                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.checked = !hiddenSet.has(f.key);
                cb.onchange = () => {
                    const set = new Set((draft.hiddenColumnFieldKeys || []).map(x => String(x || "").trim()).filter(x => !!x));
                    if (cb.checked) set.delete(f.key);
                    else set.add(f.key);
                    draft.hiddenColumnFieldKeys = Array.from(set.values());
                };
                lab.appendChild(cb);
                const txt = document.createElement("span");
                txt.textContent = f.label;
                lab.appendChild(txt);
                optionsWrap.appendChild(lab);
            }
        }

        const makeTextInput = (value: string, placeholder: string): HTMLInputElement => {
            const inp = document.createElement("input");
            inp.className = "fm-cf-input";
            inp.type = "text";
            inp.value = value || "";
            inp.placeholder = placeholder;
            return inp;
        };

        const wrapField = (labelText: string, control: HTMLElement): HTMLDivElement => {
            const wrap = document.createElement("div");
            wrap.className = "fm-cf-field";
            const lab = document.createElement("div");
            lab.className = "fm-cf-label";
            lab.textContent = labelText;
            wrap.appendChild(lab);
            wrap.appendChild(control);
            return wrap;
        };

        const makeSelect = (): FancySelect => {
            return createFancySelect("fm-cf-select");
        };

        const makeNumInput = (value: number | undefined, placeholder: string): HTMLInputElement => {
            const inp = document.createElement("input");
            inp.className = "fm-cf-input";
            inp.type = "number";
            inp.value = value != null && Number.isFinite(value as any) ? String(value) : "";
            inp.placeholder = placeholder;
            inp.min = "6";
            inp.max = "72";
            inp.step = "1";
            return inp;
        };

        const fontOptions = ["", "Segoe UI", "Arial", "Calibri"];
        const makeFontSelect = (value: string | undefined): FancySelect => {
            const s = makeSelect();
            addOption(s.select, "", "(Default font)");
            for (const f of fontOptions) {
                if (!f) continue;
                addOption(s.select, f, f);
            }
            s.setValue((value || "").trim());
            s.syncFromSelect();
            return s;
        };

        const makeColorInput = (value: string | undefined): HTMLInputElement => {
            const inp = document.createElement("input");
            inp.className = "fm-ct-color";
            inp.type = "color";
            const v = String(value || "").trim();
            inp.value = /^#([0-9a-fA-F]{6})$/.test(v) ? v : "#000000";
            return inp;
        };

        const makeBoldToggle = (value: boolean | undefined): HTMLInputElement => {
            const inp = document.createElement("input");
            inp.type = "checkbox";
            inp.checked = value === true;
            return inp;
        };

        const fields = (Array.isArray(this.customTableAvailableFields) ? this.customTableAvailableFields : []).slice();

        const aggOptions: Array<{ v: CTAggregation; t: string }> = [
            { v: "none", t: "No aggregation" },
            { v: "sum", t: "Sum" },
            { v: "avg", t: "Average" },
            { v: "min", t: "Min" },
            { v: "max", t: "Max" },
            { v: "count", t: "Count" }
        ];

        const parentsWrap = document.createElement("div");
        parentsWrap.className = "fm-ct-section";
        panel.appendChild(parentsWrap);

        const parentsTitle = document.createElement("div");
        parentsTitle.className = "fm-cf-style-title";
        parentsTitle.textContent = "Parents";
        parentsWrap.appendChild(parentsTitle);

        const parentsList = document.createElement("div");
        parentsList.className = "fm-ct-list";
        parentsWrap.appendChild(parentsList);

        const renderParents = () => {
            parentsList.textContent = "";
            for (let i = 0; i < draft.parents.length; i++) {
                const p = draft.parents[i];
                const block = document.createElement("div");
                block.className = "fm-ct-card";

                const row = document.createElement("div");
                row.className = "fm-ct-row";

                const no = makeTextInput(p.parentNo, "Parent No");
                const name = makeTextInput(p.parentName, "Parent Name");
                const del = document.createElement("button");
                del.className = "fm-cf-iconbtn";
                del.textContent = "✕";
                del.title = "Remove";
                del.onclick = () => {
                    draft.parents.splice(i, 1);
                    // Clear invalid parent references.
                    const valid = new Set(draft.parents.map(x => x.parentNo));
                    for (const c of draft.children) {
                        if (c.setParentNo && !valid.has(c.setParentNo)) c.setParentNo = "";
                    }
                    renderParents();
                    renderChildren();
                };

                no.oninput = () => {
                    const v = no.value.trim();
                    p.parentNo = v;
                    renderChildren();
                };
                name.oninput = () => {
                    p.parentName = name.value;
                };

                // Parent formatting controls
                if (!p.format) p.format = {};
                const font = makeFontSelect(p.format.fontFamily);
                const size = makeNumInput(p.format.fontSize, "Size");
                const color = makeColorInput(p.format.color);
                const bold = makeBoldToggle(p.format.bold);

                font.select.onchange = () => {
                    const v = String(font.getValue() || "").trim();
                    p.format = p.format || {};
                    p.format.fontFamily = v || undefined;
                };
                size.oninput = () => {
                    const v = Number(size.value);
                    p.format = p.format || {};
                    p.format.fontSize = Number.isFinite(v) ? Math.max(6, Math.min(72, Math.round(v))) : undefined;
                };
                color.oninput = () => {
                    p.format = p.format || {};
                    p.format.color = String(color.value || "").trim() || undefined;
                };
                bold.onchange = () => {
                    p.format = p.format || {};
                    p.format.bold = bold.checked ? true : undefined;
                };

                const boldWrap = document.createElement("label");
                boldWrap.className = "fm-ct-bold";
                boldWrap.appendChild(bold);
                const boldTxt = document.createElement("span");
                boldTxt.textContent = "Bold";
                boldWrap.appendChild(boldTxt);

                row.appendChild(wrapField("Parent No", no));
                row.appendChild(wrapField("Parent Name", name));
                row.appendChild(wrapField("Font", font.host));
                row.appendChild(wrapField("Size", size));
                row.appendChild(wrapField("Color", color));
                row.appendChild(wrapField("Bold", boldWrap));

                const rowActions = document.createElement("div");
                rowActions.className = "fm-ct-row-actions";
                rowActions.appendChild(del);
                row.appendChild(rowActions);

                block.appendChild(row);

                // Parent value mappings (optional)
                const mappings = document.createElement("div");
                mappings.className = "fm-ct-mappings";

                const mappingsTitle = document.createElement("div");
                mappingsTitle.className = "fm-cf-label";
                mappingsTitle.textContent = "Values";
                mappings.appendChild(mappingsTitle);

                const ensureValues = () => {
                    if (!Array.isArray((p as any).values) || (p as any).values.length === 0) {
                        (p as any).values = [{ field: "", aggregation: "none" }];
                    }
                    if (!draft.showValueNamesInColumns) {
                        (p as any).values = [(p as any).values[0]];
                    }
                };
                ensureValues();

                const renderMappings = () => {
                    mappings.textContent = "";
                    mappings.appendChild(mappingsTitle);
                    ensureValues();
                    const vals: any[] = (p as any).values;
                    for (let mi = 0; mi < vals.length; mi++) {
                        const m = vals[mi];
                        const mRow = document.createElement("div");
                        mRow.className = "fm-ct-map";

                        const fieldSel = makeFieldSelect(m.field);
                        const aggSel = makeAggSelect(m.aggregation);

                        const applyMeasureRules = () => {
                            const isMeasure = isMeasureField(fieldSel.getValue());
                            if (isMeasure) {
                                m.aggregation = "none";
                                aggSel.setValue("none");
                                aggSel.select.disabled = true;
                            } else {
                                aggSel.select.disabled = false;
                            }
                        };

                        fieldSel.select.onchange = () => {
                            m.field = fieldSel.getValue();
                            applyMeasureRules();
                        };
                        aggSel.select.onchange = () => {
                            m.aggregation = (aggSel.getValue() as CTAggregation) || "none";
                        };
                        applyMeasureRules();

                        const add = document.createElement("button");
                        add.className = "fm-cf-iconbtn";
                        add.textContent = "+";
                        add.title = "Add value";
                        add.onclick = () => {
                            if (!draft.showValueNamesInColumns) {
                                draft.showValueNamesInColumns = true;
                                showNamesToggle.checked = true;
                            }
                            if (!Array.isArray((p as any).values)) (p as any).values = [];
                            (p as any).values.push({ field: "", aggregation: "none" });
                            renderParents();
                            if (rerenderChildrenFn) rerenderChildrenFn();
                        };

                        const rem = document.createElement("button");
                        rem.className = "fm-cf-iconbtn";
                        rem.textContent = "−";
                        rem.title = "Remove value";
                        rem.disabled = !draft.showValueNamesInColumns || vals.length <= 1;
                        rem.onclick = () => {
                            if (!draft.showValueNamesInColumns) return;
                            if (vals.length <= 1) return;
                            vals.splice(mi, 1);
                            renderParents();
                            if (rerenderChildrenFn) rerenderChildrenFn();
                        };

                        mRow.appendChild(fieldSel.host);
                        mRow.appendChild(aggSel.host);
                        mRow.appendChild(add);
                        mRow.appendChild(rem);
                        mappings.appendChild(mRow);
                    }
                };

                renderMappings();
                block.appendChild(mappings);

                parentsList.appendChild(block);
            }
        };

        const addParentBtn = document.createElement("button");
        addParentBtn.className = "fm-cf-btn";
        addParentBtn.textContent = "Add parent";
        addParentBtn.onclick = () => {
            draft.parents.push({ parentNo: "", parentName: "" });
            renderParents();
            renderChildren();
        };
        parentsWrap.appendChild(addParentBtn);

        const childrenWrap = document.createElement("div");
        childrenWrap.className = "fm-ct-section";
        panel.appendChild(childrenWrap);

        const childrenTitle = document.createElement("div");
        childrenTitle.className = "fm-cf-style-title";
        childrenTitle.textContent = "Children";
        childrenWrap.appendChild(childrenTitle);

        const childrenList = document.createElement("div");
        childrenList.className = "fm-ct-list";
        childrenWrap.appendChild(childrenList);

        const makeParentNoSelect = (value: string): FancySelect => {
            const s = makeSelect();
            addOption(s.select, "", "(No parent)");
            for (const p of draft.parents) {
                const label = p.parentName ? `${p.parentNo} — ${p.parentName}` : p.parentNo;
                addOption(s.select, p.parentNo, label);
            }
            s.setValue(value || "");
            s.syncFromSelect();
            return s;
        };

        const makeFieldSelect = (value: string): FancySelect => {
            const s = makeSelect();
            addOption(s.select, "", "Select value field");
            for (const f of fields) addOption(s.select, f, f);
            s.setValue(value || "");
            s.syncFromSelect();
            return s;
        };

        const makeAggSelect = (value: CTAggregation): FancySelect => {
            const s = makeSelect();
            for (const a of aggOptions) addOption(s.select, a.v, a.t);
            s.setValue(value || "none");
            s.syncFromSelect();
            return s;
        };

        const makeCategoryFieldSelect = (value: string, placeholder: string): FancySelect => {
            const s = makeSelect();
            addOption(s.select, "", placeholder);
            for (const f of (this.customTableAvailableCategoryFields || [])) addOption(s.select, f, f);
            s.setValue(value || "");
            s.syncFromSelect();
            return s;
        };

        const isMeasureField = (field: string): boolean => {
            const key = String(field || "").trim();
            return this.customTableFieldIsMeasure.get(key) === true;
        };

        const renderChildren = () => {
            childrenList.textContent = "";
            for (let i = 0; i < draft.children.length; i++) {
                const c = draft.children[i];
                const row = document.createElement("div");
                row.className = "fm-ct-child fm-ct-card";

                const top = document.createElement("div");
                top.className = "fm-ct-child-top";

                const parentSel = makeParentNoSelect(c.setParentNo);
                const childName = makeTextInput(c.childName, "Child Name");

                const autoChk = document.createElement("input");
                autoChk.type = "checkbox";
                autoChk.checked = !!String((c as any).childNameFromField || "").trim();

                const autoWrap = document.createElement("label");
                autoWrap.className = "fm-ct-bold";
                autoWrap.appendChild(autoChk);
                const autoTxt = document.createElement("span");
                autoTxt.textContent = "From field";
                autoWrap.appendChild(autoTxt);

                const childFieldSel = makeCategoryFieldSelect(String((c as any).childNameFromField || "").trim(), "Child name field");
                const parentMatchSel = makeCategoryFieldSelect(String((c as any).parentMatchField || "").trim(), "(Optional) Parent match field");

                const childFieldWrap = wrapField("Child name field", childFieldSel.host);
                const parentMatchWrap = wrapField("Parent match field", parentMatchSel.host);

                const syncAutoUi = () => {
                    const enabled = autoChk.checked;
                    childFieldWrap.style.display = enabled ? "" : "none";
                    parentMatchWrap.style.display = enabled ? "" : "none";
                    childName.disabled = enabled;
                    if (enabled) {
                        c.childName = "";
                    } else {
                        (c as any).childNameFromField = undefined;
                        (c as any).parentMatchField = undefined;
                    }
                };

                if (!c.format) c.format = {};
                const font = makeFontSelect(c.format.fontFamily);
                const size = makeNumInput(c.format.fontSize, "Size");
                const color = makeColorInput(c.format.color);
                const bold = makeBoldToggle(c.format.bold);

                font.select.onchange = () => {
                    const v = String(font.getValue() || "").trim();
                    c.format = c.format || {};
                    c.format.fontFamily = v || undefined;
                };
                size.oninput = () => {
                    const v = Number(size.value);
                    c.format = c.format || {};
                    c.format.fontSize = Number.isFinite(v) ? Math.max(6, Math.min(72, Math.round(v))) : undefined;
                };
                color.oninput = () => {
                    c.format = c.format || {};
                    c.format.color = String(color.value || "").trim() || undefined;
                };
                bold.onchange = () => {
                    c.format = c.format || {};
                    c.format.bold = bold.checked ? true : undefined;
                };

                const boldWrap = document.createElement("label");
                boldWrap.className = "fm-ct-bold";
                boldWrap.appendChild(bold);
                const boldTxt = document.createElement("span");
                boldTxt.textContent = "Bold";
                boldWrap.appendChild(boldTxt);

                parentSel.select.onchange = () => {
                    c.setParentNo = parentSel.getValue();
                };
                childName.oninput = () => {
                    c.childName = childName.value;
                };

                autoChk.onchange = () => {
                    if (autoChk.checked) {
                        (c as any).childNameFromField = String(childFieldSel.getValue() || "").trim() || undefined;
                        (c as any).parentMatchField = String(parentMatchSel.getValue() || "").trim() || undefined;
                    } else {
                        (c as any).childNameFromField = undefined;
                        (c as any).parentMatchField = undefined;
                    }
                    syncAutoUi();
                };

                childFieldSel.select.onchange = () => {
                    (c as any).childNameFromField = String(childFieldSel.getValue() || "").trim() || undefined;
                };
                parentMatchSel.select.onchange = () => {
                    (c as any).parentMatchField = String(parentMatchSel.getValue() || "").trim() || undefined;
                };

                const del = document.createElement("button");
                del.className = "fm-cf-iconbtn";
                del.textContent = "✕";
                del.title = "Remove";
                del.onclick = () => {
                    draft.children.splice(i, 1);
                    renderChildren();
                };

                top.appendChild(wrapField("Parent", parentSel.host));
                top.appendChild(wrapField("Child name", childName));
                top.appendChild(wrapField("Auto", autoWrap));
                top.appendChild(childFieldWrap);
                top.appendChild(parentMatchWrap);
                top.appendChild(wrapField("Font", font.host));
                top.appendChild(wrapField("Size", size));
                top.appendChild(wrapField("Color", color));
                top.appendChild(wrapField("Bold", boldWrap));

                const topActions = document.createElement("div");
                topActions.className = "fm-ct-row-actions";
                topActions.appendChild(del);
                top.appendChild(topActions);
                row.appendChild(top);

                syncAutoUi();

                const mappings = document.createElement("div");
                mappings.className = "fm-ct-mappings";

                const mappingsTitle = document.createElement("div");
                mappingsTitle.className = "fm-cf-label";
                mappingsTitle.textContent = "Values";
                mappings.appendChild(mappingsTitle);

                const ensureValues = () => {
                    if (!Array.isArray(c.values) || c.values.length === 0) {
                        c.values = [{ field: "", aggregation: "none" }];
                    }
                    if (!draft.showValueNamesInColumns) {
                        c.values = [c.values[0]];
                    }
                };
                ensureValues();

                const renderMappings = () => {
                    mappings.textContent = "";
                    mappings.appendChild(mappingsTitle);
                    ensureValues();
                    for (let mi = 0; mi < c.values.length; mi++) {
                        const m = c.values[mi];
                        const mRow = document.createElement("div");
                        mRow.className = "fm-ct-map";

                        const fieldSel = makeFieldSelect(m.field);
                        const aggSel = makeAggSelect(m.aggregation);

                        const applyMeasureRules = () => {
                            const isMeasure = isMeasureField(fieldSel.getValue());
                            if (isMeasure) {
                                m.aggregation = "none";
                                aggSel.setValue("none");
                                aggSel.select.disabled = true;
                            } else {
                                aggSel.select.disabled = false;
                            }
                        };

                        fieldSel.select.onchange = () => {
                            m.field = fieldSel.getValue();
                            applyMeasureRules();
                        };
                        aggSel.select.onchange = () => {
                            m.aggregation = (aggSel.getValue() as CTAggregation) || "none";
                        };

                        applyMeasureRules();

                        const add = document.createElement("button");
                        add.className = "fm-cf-iconbtn";
                        add.textContent = "+";
                        add.title = "Add value";
                        add.onclick = () => {
                            if (!draft.showValueNamesInColumns) {
                                draft.showValueNamesInColumns = true;
                                showNamesToggle.checked = true;
                            }
                            if (!Array.isArray(c.values)) c.values = [];
                            c.values.push({ field: "", aggregation: "none" });
                            renderChildren();
                        };

                        const rem = document.createElement("button");
                        rem.className = "fm-cf-iconbtn";
                        rem.textContent = "−";
                        rem.title = "Remove value";
                        rem.disabled = !draft.showValueNamesInColumns || c.values.length <= 1;
                        rem.onclick = () => {
                            if (!draft.showValueNamesInColumns) return;
                            if (c.values.length <= 1) return;
                            c.values.splice(mi, 1);
                            renderChildren();
                        };

                        mRow.appendChild(fieldSel.host);
                        mRow.appendChild(aggSel.host);
                        mRow.appendChild(add);
                        mRow.appendChild(rem);
                        mappings.appendChild(mRow);
                    }
                };

                renderMappings();
                row.appendChild(mappings);
                childrenList.appendChild(row);
            }
        };

        const addChildBtn = document.createElement("button");
        addChildBtn.className = "fm-cf-btn";
        addChildBtn.textContent = "Add child";
        addChildBtn.onclick = () => {
            draft.children.push({ id: cryptoId("ctc"), setParentNo: "", childName: "", childNameFromField: undefined, parentMatchField: undefined, values: [{ field: "", aggregation: "none" }] } as any);
            renderChildren();
        };
        childrenWrap.appendChild(addChildBtn);

        renderParents();
        renderChildren();
        rerenderChildrenFn = renderChildren;

        const actions = document.createElement("div");
        actions.className = "fm-cf-actions";
        panel.appendChild(actions);

        const cancel = document.createElement("button");
        cancel.className = "fm-cf-btn";
        cancel.textContent = "Cancel";
        cancel.onclick = () => this.closeCustomTableEditor();
        actions.appendChild(cancel);

        const apply = document.createElement("button");
        apply.className = "fm-cf-btn fm-cf-btn-primary";
        apply.textContent = "Apply";
        apply.onclick = () => {
            // Clean and normalize draft before saving.
            const cleanFont = (s: any): string | undefined => {
                const v = String(s ?? "").trim();
                if (!v) return undefined;
                const ok = ["Segoe UI", "Arial", "Calibri"].includes(v);
                return ok ? v : undefined;
            };
            const cleanFontSize = (n: any): number | undefined => {
                const v = Number(n);
                if (!Number.isFinite(v)) return undefined;
                return Math.max(6, Math.min(72, Math.round(v)));
            };
            const cleanColor = (c: any): string | undefined => {
                const s = String(c ?? "").trim();
                if (!s) return undefined;
                if (/^#([0-9a-fA-F]{6})$/.test(s)) return s;
                if (/^#([0-9a-fA-F]{3})$/.test(s)) return s;
                return undefined;
            };
            const cleanFormat = (f: any): any => {
                if (!f || typeof f !== "object") return undefined;
                const fontFamily = cleanFont((f as any).fontFamily);
                const fontSize = cleanFontSize((f as any).fontSize);
                const color = cleanColor((f as any).color);
                const bold = (f as any).bold === true ? true : undefined;
                if (!fontFamily && !fontSize && !color && !bold) return undefined;
                return { fontFamily, fontSize, color, bold };
            };

            draft.parents = (draft.parents || [])
                .map(p => ({
                    parentNo: String(p.parentNo || "").trim(),
                    parentName: String(p.parentName || "").trim(),
                    values: (() => {
                        const cleaned = (Array.isArray((p as any).values) ? ((p as any).values as any[]) : [])
                            .map(v => ({ field: String(v.field || "").trim(), aggregation: (String((v as any).aggregation || "none").trim().toLowerCase() as CTAggregation) || "none" }))
                            .filter(v => !!v.field);
                        return (draft.showValueNamesInColumns === true) ? cleaned : cleaned.slice(0, 1);
                    })(),
                    format: cleanFormat((p as any).format)
                }))
                .filter(p => !!p.parentNo);
            draft.children = (draft.children || [])
                .map(c => ({
                    id: String(c.id || "").trim() || cryptoId("ctc"),
                    setParentNo: String(c.setParentNo || "").trim(),
                    childName: String(c.childName || "").trim(),
                    childNameFromField: String((c as any).childNameFromField || "").trim() || undefined,
                    parentMatchField: String((c as any).parentMatchField || "").trim() || undefined,
                    values: (() => {
                        const cleaned = (Array.isArray(c.values) ? c.values : [])
                            .map(v => ({ field: String(v.field || "").trim(), aggregation: (String((v as any).aggregation || "none").trim().toLowerCase() as CTAggregation) || "none" }))
                            .filter(v => !!v.field);
                        return (draft.showValueNamesInColumns === true) ? cleaned : cleaned.slice(0, 1);
                    })(),
                    format: cleanFormat((c as any).format)
                }))
                .filter(c => !!c.childName || !!(c as any).childNameFromField);

            draft.showValueNamesInColumns = draft.showValueNamesInColumns === true;
            draft.hiddenColumnFieldKeys = Array.isArray(draft.hiddenColumnFieldKeys)
                ? (draft.hiddenColumnFieldKeys as any[]).map(x => String(x ?? "").trim()).filter(x => !!x)
                : [];

            // Force measure mappings to "none" aggregation.
            for (const c of draft.children) {
                for (const v of c.values) {
                    if (isMeasureField(v.field)) v.aggregation = "none";
                }
            }
            for (const p of draft.parents) {
                for (const v of (p.values || [])) {
                    if (isMeasureField(v.field)) v.aggregation = "none";
                }
            }

            const json = JSON.stringify({
                version: 1,
                parents: draft.parents,
                children: draft.children,
                showValueNamesInColumns: draft.showValueNamesInColumns === true,
                hiddenColumnFieldKeys: draft.hiddenColumnFieldKeys
            } as CustomTableConfig);
            this.persistObjectProperties("customTable", { tableJson: json, showEditor: false });

            // Keep local state in sync even before Power BI calls update again.
            this.customTableJson = json;
            this.pendingCustomTableJson = json;
            this.pendingCustomTableSetAt = Date.now();
            this.customTableEditorDraft = null;

            // Force an immediate redraw so the report reflects changes right away.
            if (this.lastDataView && this.lastViewport) {
                const model = this.buildModel(this.lastDataView);
                this.render(model, this.lastViewport);
            }
        };
        actions.appendChild(apply);

        // Clicking outside closes.
        overlay.onclick = (e: MouseEvent) => {
            if (e.target === overlay) this.closeCustomTableEditor();
        };

        this.root.appendChild(overlay);
    }

    private openCfEditor(): void {
        this.ensureCfDraft();
        this.persistObjectProperties("conditionalFormatting", { showEditor: true });
    }

    private renderCfEditor(model: Model): void {
        const cfCard: any = (this.formattingSettings as any).conditionalFormatting;
        const enabled = (cfCard?.enabled?.value as boolean | undefined) ?? false;
        const showEditor = (cfCard?.showEditor?.value as boolean | undefined) ?? false;
        if (!enabled) return;

        if (!showEditor) return;

        const draft = this.ensureCfDraft();

        const uniqueFields: string[] = (() => {
            if (Array.isArray(this.cfAvailableFields) && this.cfAvailableFields.length > 0) {
                return this.cfAvailableFields.slice();
            }

            const inferLabel = (c: BucketedColumnKey): string => {
                const direct = (c.leafLabel || "").trim();
                if (direct) return direct;
                const key = (c.key || "").trim();
                if (!key) return "";
                const parts = key.split("||");
                if (parts.length < 2) return "";
                const measure = (parts[parts.length - 1] || "").trim();
                if (!measure) return "";
                return measure;
            };

            const set = new Set<string>();
            for (const c of model.columns) {
                const label = inferLabel(c);
                if (label) set.add(label);
            }
            return Array.from(set.values()).sort((a, b) => a.localeCompare(b));
        })();

        const overlay = document.createElement("div");
        overlay.className = "fm-cf-overlay";

        const panel = document.createElement("div");
        panel.className = "fm-cf-panel";
        overlay.appendChild(panel);

        const header = document.createElement("div");
        header.className = "fm-cf-header";
        panel.appendChild(header);

        const title = document.createElement("div");
        title.className = "fm-cf-title";
        title.textContent = "Conditional formatting";
        header.appendChild(title);

        type FormatProp = CFFormatProp;
        type ApplyToView = CFRuleTarget | "all";
        let currentProp: FormatProp = "background";

        // Format tabs (separate Background vs Font views)
        const tabs = document.createElement("div");
        tabs.className = "fm-cf-tabs";
        header.appendChild(tabs);

        const mkTab = (label: string, prop: FormatProp) => {
            const b = document.createElement("button");
            b.type = "button";
            b.className = "fm-cf-tab";
            b.textContent = label;
            b.addEventListener("click", () => {
                currentProp = prop;
                rerender();
            });
            tabs.appendChild(b);
            return b;
        };

        const tabBg = mkTab("Background", "background");
        const tabFont = mkTab("Font", "fontColor");
        const tabIcon = mkTab("Icons", "icon");

        const toolbar = document.createElement("div");
        toolbar.className = "fm-cf-toolbar";
        header.appendChild(toolbar);

        // Default to showing all saved rules so history is always visible.
        let currentTarget: ApplyToView = "all";
        let lastSpecificTarget: CFRuleTarget = "cell";
        let basedOnField: string = ""; // leaf label filter for values (optional)
        let iconPlacement: CFIconPlacement = "left";

        const getCurrentSurface = (): CFSurfaceSettingsEx => {
            if (currentTarget === "all") return { formatStyle: "rules" };
            const t = currentTarget as CFRuleTarget;
            return this.getSurfaceSettings(draft, t, currentProp) as any;
        };

        const setCurrentSurface = (s: CFSurfaceSettingsEx): void => {
            if (currentTarget === "all") return;
            const t = currentTarget as CFRuleTarget;
            this.setSurfaceSettings(draft, t, currentProp, s as any);
        };

        const mkField = (label: string, child: HTMLElement) => {
            const wrap = document.createElement("div");
            wrap.className = "fm-cf-field";
            const l = document.createElement("div");
            l.className = "fm-cf-label";
            l.textContent = label;
            wrap.appendChild(l);
            wrap.appendChild(child);
            toolbar.appendChild(wrap);
            return wrap;
        };

        const applyToSel = createFancySelect("fm-cf-select");
        addOption(applyToSel.select, "all", "All");
        addOption(applyToSel.select, "cell", "Values only");
        addOption(applyToSel.select, "rowHeader", "Row headers");
        addOption(applyToSel.select, "columnHeader", "Column headers");
        applyToSel.setValue(currentTarget);
        applyToSel.syncFromSelect();
        applyToSel.select.addEventListener("change", () => {
            const v = applyToSel.getValue() as ApplyToView;
            currentTarget = v;
            if (v !== "all") lastSpecificTarget = v;
            rerender();
        });
        mkField("Apply to", applyToSel.host);

        const formatStyleSel = createFancySelect("fm-cf-select");
        addOption(formatStyleSel.select, "rules", "Rules");
        addOption(formatStyleSel.select, "gradient", "Gradient");
        addOption(formatStyleSel.select, "fieldValue", "Field value");
        formatStyleSel.syncFromSelect();
        formatStyleSel.select.addEventListener("change", () => {
            const v = formatStyleSel.getValue() as CFFormatStyle;
            if (currentProp === "icon" && v === "gradient") {
                // Icons don't support gradients.
                formatStyleSel.setValue("rules");
                setCurrentSurface({ formatStyle: "rules", iconPlacement } as any);
                rerender();
                return;
            }
            if (v === "rules") {
                const prev = getCurrentSurface();
                if (currentProp === "icon") {
                    setCurrentSurface({ formatStyle: "rules", iconPlacement: (prev as any).iconPlacement ?? iconPlacement } as any);
                } else {
                    setCurrentSurface({ formatStyle: "rules" } as any);
                }
            } else if (v === "gradient") {
                const existing = getCurrentSurface();
                if (existing.formatStyle === "gradient") {
                    // keep
                    setCurrentSurface({ ...existing, basedOnField } as any);
                } else {
                    const s: CFGradientSettings = {
                        formatStyle: "gradient",
                        basedOnField: basedOnField,
                        emptyValues: "blank",
                        min: { type: "auto", color: "#ffffff" },
                        max: { type: "auto", color: "#000000" },
                        useMid: false,
                        midColor: "#808080"
                    };
                    setCurrentSurface(s as any);
                }
            } else {
                const s: CFFieldValueSettings = {
                    formatStyle: "fieldValue",
                    basedOnField: basedOnField
                };
                const prev = getCurrentSurface();
                if (currentProp === "icon") {
                    setCurrentSurface({ ...s, iconPlacement: (prev as any).iconPlacement ?? iconPlacement } as any);
                } else {
                    setCurrentSurface(s as any);
                }
            }
            rerender();
        });
        mkField("Format style", formatStyleSel.host);

        const iconPlacementSel = createFancySelect("fm-cf-select");
        addOption(iconPlacementSel.select, "left", "Icon left of value");
        addOption(iconPlacementSel.select, "right", "Icon right of value");
        addOption(iconPlacementSel.select, "only", "Only icon (hide value)");
        iconPlacementSel.setValue(iconPlacement);
        iconPlacementSel.syncFromSelect();
        iconPlacementSel.select.addEventListener("change", () => {
            iconPlacement = iconPlacementSel.getValue() as CFIconPlacement;
            const s = getCurrentSurface();
            if (currentTarget !== "all" && currentProp === "icon") {
                setCurrentSurface({ ...(s as any), iconPlacement } as any);
            }
            rerender();
        });
        const iconPlacementField = mkField("Icon display", iconPlacementSel.host);

        const basedOnSel = createFancySelect("fm-cf-select");
        addOption(basedOnSel.select, "", "(All fields)");
        for (const f of uniqueFields) addOption(basedOnSel.select, f, f);
        basedOnSel.setValue(basedOnField);
        basedOnSel.syncFromSelect();
        basedOnSel.select.addEventListener("change", () => {
            basedOnField = basedOnSel.getValue();
            const s = getCurrentSurface();
            if (currentTarget !== "all" && s.formatStyle !== "rules") {
                if (s.formatStyle === "gradient") {
                    setCurrentSurface({ ...s, basedOnField } as any);
                } else if (s.formatStyle === "fieldValue") {
                    setCurrentSurface({ ...s, basedOnField } as any);
                }
            }
            rerender();
        });
        mkField("Based on field", basedOnSel.host);

        const reverseBtn = document.createElement("button");
        reverseBtn.type = "button";
        reverseBtn.className = "fm-cf-btn";
        reverseBtn.textContent = "Reverse order";
        reverseBtn.addEventListener("click", () => {
            draft.rules = [...(draft.rules || [])].reverse();
            rerender();
        });
        const reverseWrap = document.createElement("div");
        reverseWrap.className = "fm-cf-field";
        reverseWrap.appendChild(document.createElement("div"));
        reverseWrap.appendChild(reverseBtn);
        toolbar.appendChild(reverseWrap);

        const list = document.createElement("div");
        list.className = "fm-cf-list";
        panel.appendChild(list);

        const stylePanel = document.createElement("div");
        stylePanel.className = "fm-cf-style";
        panel.appendChild(stylePanel);

        const renderRuleRow = (rule: CFRule, index: number, total: number) => {
            const row = document.createElement("div");
            row.className = "fm-cf-rule";

            const enabledEl = document.createElement("input");
            enabledEl.type = "checkbox";
            enabledEl.checked = rule.enabled;
            enabledEl.addEventListener("change", () => { rule.enabled = enabledEl.checked; });
            row.appendChild(enabledEl);

            const sentence = document.createElement("div");
            sentence.className = "fm-cf-sentence";
            row.appendChild(sentence);

            const addText = (t: string) => {
                const s = document.createElement("span");
                s.className = "fm-cf-text";
                s.textContent = t;
                sentence.appendChild(s);
            };

            const mkSel = (opts: Array<{ v: string; l: string }>, value: string, onChange: (v: string) => void) => {
                const sel = createFancySelect("fm-cf-select");
                for (const o of opts) addOption(sel.select, o.v, o.l);
                sel.setValue(value);
                sel.syncFromSelect();
                sel.select.addEventListener("change", () => onChange(sel.getValue()));
                sentence.appendChild(sel.host);
                return sel;
            };

            const mkInput = (value: string, placeholder: string, onInput: (v: string) => void) => {
                const inp = document.createElement("input");
                inp.className = "fm-cf-input";
                inp.type = "text";
                inp.value = value;
                inp.placeholder = placeholder;
                inp.addEventListener("input", () => onInput(inp.value));
                sentence.appendChild(inp);
                return inp;
            };

            const ruleTarget: CFRuleTarget = (currentTarget === "all") ? (rule.target || "cell") : (currentTarget as CFRuleTarget);

            if (!rule.formatProp) {
                // fallback (older draft)
                rule.formatProp = (rule.style?.iconName ? "icon" : (rule.style?.background ? "background" : (rule.style?.fontColor ? "fontColor" : currentProp)));
            }

            // In All view, allow changing target per-rule.
            if (currentTarget === "all") {
                mkSel(
                    [
                        { v: "cell", l: "Values" },
                        { v: "rowHeader", l: "Row headers" },
                        { v: "columnHeader", l: "Column headers" }
                    ],
                    ruleTarget,
                    (v) => { rule.target = v as CFRuleTarget; rerender(); }
                );

                mkSel(
                    [
                        { v: "background", l: "Background" },
                        { v: "fontColor", l: "Font" },
                        { v: "icon", l: "Icons" }
                    ],
                    rule.formatProp,
                    (v) => { rule.formatProp = v as CFFormatProp; rerender(); }
                );
            }

            if (ruleTarget === "cell") {
                rule.conditionType = "value";
                addText("If value is");
                mkSel(
                    [
                        { v: "gte", l: ">=" },
                        { v: "gt", l: ">" },
                        { v: "lte", l: "<=" },
                        { v: "lt", l: "<" },
                        { v: "eq", l: "=" },
                        { v: "between", l: "between" }
                    ],
                    rule.operator,
                    (v) => { rule.operator = v as any; rerender(); }
                );

                mkInput(rule.value1, "Enter a value", (v) => { rule.value1 = v; });

                if (rule.operator === "between") {
                    addText("and");
                    mkInput(rule.value2, "Enter a value", (v) => { rule.value2 = v; });
                } else {
                    rule.value2 = "";
                }

                // Apply filter based on field (maps to column match for the current rule set)
                if (currentTarget === "cell" && basedOnField) {
                    rule.colMatchMode = "contains";
                    rule.colMatchText = basedOnField;
                } else {
                    rule.colMatchMode = "any";
                    rule.colMatchText = "";
                }
            } else {
                rule.conditionType = "text";
                addText("If text");
                mkSel(
                    [
                        { v: "contains", l: "contains" },
                        { v: "eq", l: "equals" }
                    ],
                    rule.operator,
                    (v) => { rule.operator = v as any; }
                );
                mkInput(rule.value1, "Enter text", (v) => { rule.value1 = v; });
                rule.value2 = "";
            }

            // Scope option for row/column headers.
            if (ruleTarget === "rowHeader") {
                const wrap = document.createElement("span");
                wrap.className = "fm-cf-text";
                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.checked = (rule.scope || "self") === "entireRow";
                cb.addEventListener("change", () => {
                    rule.scope = cb.checked ? "entireRow" : "self";
                });
                wrap.appendChild(cb);
                const lbl = document.createElement("span");
                lbl.textContent = " Apply to entire row";
                wrap.appendChild(lbl);
                sentence.appendChild(wrap);
            } else if (ruleTarget === "columnHeader") {
                const wrap = document.createElement("span");
                wrap.className = "fm-cf-text";
                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.checked = (rule.scope || "self") === "entireColumn";
                cb.addEventListener("change", () => {
                    rule.scope = cb.checked ? "entireColumn" : "self";
                });
                wrap.appendChild(cb);
                const lbl = document.createElement("span");
                lbl.textContent = " Apply to entire column";
                wrap.appendChild(lbl);
                sentence.appendChild(wrap);
            }

            addText("then");

            const effectiveProp: CFFormatProp = (currentTarget === "all") ? (rule.formatProp || currentProp) : currentProp;

            if (effectiveProp === "icon") {
                const iconSel = createFancySelect("fm-cf-select");
                addOption(iconSel.select, "", "(None)");
                for (const ic of CF_BUILTIN_ICONS) addOption(iconSel.select, ic.key, ic.label);
                iconSel.setValue((rule.style?.iconName || "").trim());
                iconSel.syncFromSelect();
                iconSel.select.addEventListener("change", () => {
                    if (currentTarget !== "all") rule.formatProp = currentProp;
                    const v = iconSel.getValue();
                    if (!v) {
                        delete (rule.style as any).iconName;
                    } else {
                        rule.style.iconName = v;
                    }
                });
                row.appendChild(iconSel.host);
            } else {
                const color = document.createElement("input");
                color.type = "color";
                color.className = "fm-cf-color";
                const current = effectiveProp === "background" ? (rule.style.background || "") : (rule.style.fontColor || "");
                color.value = current && /^#/.test(current) ? current : (effectiveProp === "background" ? "#ffffff" : "#000000");
                color.addEventListener("input", () => {
                    if (currentTarget !== "all") rule.formatProp = currentProp;
                    if (effectiveProp === "background") rule.style.background = color.value;
                    else rule.style.fontColor = color.value;
                });
                row.appendChild(color);
            }

            const up = document.createElement("button");
            up.type = "button";
            up.className = "fm-cf-iconbtn";
            up.textContent = "↑";
            up.disabled = index <= 0;
            up.addEventListener("click", () => {
                const arr = [...draft.rules];
                const tmp = arr[index - 1];
                arr[index - 1] = arr[index];
                arr[index] = tmp;
                draft.rules = arr;
                rerender();
            });
            row.appendChild(up);

            const down = document.createElement("button");
            down.type = "button";
            down.className = "fm-cf-iconbtn";
            down.textContent = "↓";
            down.disabled = index >= total - 1;
            down.addEventListener("click", () => {
                const arr = [...draft.rules];
                const tmp = arr[index + 1];
                arr[index + 1] = arr[index];
                arr[index] = tmp;
                draft.rules = arr;
                rerender();
            });
            row.appendChild(down);

            const remove = document.createElement("button");
            remove.type = "button";
            remove.className = "fm-cf-remove";
            remove.textContent = "✕";
            remove.addEventListener("click", () => {
                draft.rules = draft.rules.filter(r => r.id !== rule.id);
                rerender();
            });
            row.appendChild(remove);

            list.appendChild(row);
        };

        const rerender = () => {
            while (list.firstChild) list.removeChild(list.firstChild);
            while (stylePanel.firstChild) stylePanel.removeChild(stylePanel.firstChild);

            const surface = getCurrentSurface();
            if (currentProp === "icon") {
                const ip = (surface as any).iconPlacement as CFIconPlacement | undefined;
                iconPlacement = ip ?? iconPlacement;
                iconPlacementSel.setValue(iconPlacement);
            }

            // For Field value formatting, Power BI requires a field selection.
            // If none is selected, default to the first available field so it doesn't look "broken".
            if (currentTarget !== "all" && surface.formatStyle === "fieldValue" && !((surface as any).basedOnField || "").trim()) {
                const first = (Array.isArray(this.cfAvailableMeasureFields) && this.cfAvailableMeasureFields.length > 0)
                    ? this.cfAvailableMeasureFields[0]
                    : uniqueFields[0];
                if (first) {
                    basedOnField = first;
                    setCurrentSurface({ ...(surface as any), basedOnField } as any);
                    basedOnSel.setValue(basedOnField);
                }
            }

            basedOnSel.select.disabled = currentTarget === "all";
            formatStyleSel.select.disabled = currentTarget === "all";
            iconPlacementSel.select.disabled = currentTarget === "all" || currentProp !== "icon";
            iconPlacementField.style.display = currentProp === "icon" ? "" : "none";
            formatStyleSel.setValue((currentTarget === "all") ? "rules" : surface.formatStyle);

            tabBg.classList.toggle("fm-cf-tab-active", currentProp === "background");
            tabFont.classList.toggle("fm-cf-tab-active", currentProp === "fontColor");
            tabIcon.classList.toggle("fm-cf-tab-active", currentProp === "icon");

            const optGradient = formatStyleSel.select.querySelector('option[value="gradient"]') as HTMLOptionElement | null;
            if (optGradient) optGradient.disabled = currentProp === "icon";

            // Show the list area for both Rules and non-rules surfaces.
            // For non-rules, we show a single applied entry so users can remove/edit like Rules mode.
            const showRules = surface.formatStyle === "rules";
            list.style.display = "grid";
            reverseBtn.disabled = !showRules;
            // Based on field acts as a column filter in Rules mode only.
            if (!showRules) {
                // Keep basedOnField in the surface settings if applicable.
                if (surface.formatStyle === "gradient" || surface.formatStyle === "fieldValue") {
                    // ensure the dropdown reflects surface
                    const bof = (surface as any).basedOnField as string | undefined;
                    basedOnField = bof ?? basedOnField;
                    basedOnSel.setValue(basedOnField);
                }
            }

            if (!showRules) {
                const sectionTitle = document.createElement("div");
                sectionTitle.className = "fm-cf-style-title";
                sectionTitle.textContent = surface.formatStyle === "gradient" ? "Gradient" : "Field value";
                stylePanel.appendChild(sectionTitle);

                // Summary + clear button (non-rules surfaces don't show rows in the rule list).
                const summary = document.createElement("div");
                summary.className = "fm-cf-style-note";
                const bof = ((surface as any).basedOnField || "").trim();
                summary.textContent = bof
                    ? `Applied: ${sectionTitle.textContent} (Based on: ${bof})`
                    : `Applied: ${sectionTitle.textContent}`;
                stylePanel.appendChild(summary);

                const clearWrap = document.createElement("div");
                clearWrap.className = "fm-cf-actions";
                const clearBtn = document.createElement("button");
                clearBtn.type = "button";
                clearBtn.className = "fm-cf-btn";
                clearBtn.textContent = "Clear formatting";
                clearBtn.disabled = currentTarget === "all";
                clearBtn.addEventListener("click", () => {
                    // Reset current surface back to Rules.
                    if (currentTarget === "all") return;
                    if (currentProp === "icon") {
                        setCurrentSurface({ formatStyle: "rules", iconPlacement } as any);
                    } else {
                        setCurrentSurface({ formatStyle: "rules" } as any);
                    }
                    // Keep UI state reasonable.
                    basedOnField = "";
                    basedOnSel.setValue("");
                    rerender();
                });
                clearWrap.appendChild(clearBtn);
                stylePanel.appendChild(clearWrap);

                // Show a single applied entry in the list, so users can remove/edit like Rules mode.
                const appliedRow = document.createElement("div");
                appliedRow.className = "fm-cf-rule";

                const enabledEl = document.createElement("input");
                enabledEl.type = "checkbox";
                enabledEl.checked = true;
                enabledEl.disabled = true;
                appliedRow.appendChild(enabledEl);

                const sentence = document.createElement("div");
                sentence.className = "fm-cf-sentence";
                const addText = (t: string) => {
                    const sp = document.createElement("span");
                    sp.className = "fm-cf-text";
                    sp.textContent = t;
                    sentence.appendChild(sp);
                };
                const targetLabel = currentTarget === "cell" ? "Values" : (currentTarget === "rowHeader" ? "Row headers" : "Column headers");
                const propLabel = currentProp === "background" ? "Background" : (currentProp === "fontColor" ? "Font" : "Icons");
                addText(`${sectionTitle.textContent} applied to ${targetLabel} (${propLabel})`);
                if (bof) addText(`Based on: ${bof}`);
                appliedRow.appendChild(sentence);

                const editBtn = document.createElement("button");
                editBtn.type = "button";
                editBtn.className = "fm-cf-btn";
                editBtn.textContent = "Edit";
                editBtn.addEventListener("click", () => {
                    // Keep it simple: focus the most relevant control.
                    if (surface.formatStyle === "gradient" || surface.formatStyle === "fieldValue") {
                        const btn = basedOnSel.host.querySelector(".fm-dd-btn") as HTMLButtonElement | null;
                        if (btn) btn.focus();
                    }
                });
                appliedRow.appendChild(editBtn);

                const up = document.createElement("button");
                up.type = "button";
                up.className = "fm-cf-iconbtn";
                up.textContent = "↑";
                up.disabled = true;
                appliedRow.appendChild(up);

                const down = document.createElement("button");
                down.type = "button";
                down.className = "fm-cf-iconbtn";
                down.textContent = "↓";
                down.disabled = true;
                appliedRow.appendChild(down);

                const remove = document.createElement("button");
                remove.type = "button";
                remove.className = "fm-cf-remove";
                remove.textContent = "✕";
                remove.addEventListener("click", () => {
                    // Same as Clear formatting.
                    if (currentTarget === "all") return;
                    if (currentProp === "icon") {
                        setCurrentSurface({ formatStyle: "rules", iconPlacement } as any);
                    } else {
                        setCurrentSurface({ formatStyle: "rules" } as any);
                    }
                    basedOnField = "";
                    basedOnSel.setValue("");
                    rerender();
                });
                appliedRow.appendChild(remove);

                list.appendChild(appliedRow);

                if (surface.formatStyle === "gradient") {
                    const s = surface as CFGradientSettings;

                    const grid = document.createElement("div");
                    grid.className = "fm-cf-style-grid";
                    stylePanel.appendChild(grid);

                    const mkStyleField = (label: string, child: HTMLElement) => {
                        const wrap = document.createElement("div");
                        wrap.className = "fm-cf-field";
                        const l = document.createElement("div");
                        l.className = "fm-cf-label";
                        l.textContent = label;
                        wrap.appendChild(l);
                        wrap.appendChild(child);
                        grid.appendChild(wrap);
                    };

                    const emptySel = createFancySelect("fm-cf-select");
                    addOption(emptySel.select, "blank", "Blank (no color)");
                    addOption(emptySel.select, "zero", "Treat as 0");
                    emptySel.setValue(s.emptyValues);
                    emptySel.syncFromSelect();
                    emptySel.select.addEventListener("change", () => {
                        s.emptyValues = emptySel.getValue() as any;
                    });
                    mkStyleField("Empty values", emptySel.host);

                    const mkBoundEditor = (name: "min" | "max") => {
                        const bound = name === "min" ? s.min : s.max;

                        const row = document.createElement("div");
                        row.className = "fm-cf-style-bound";

                        const typeSel = createFancySelect("fm-cf-select");
                        addOption(typeSel.select, "auto", "Auto");
                        addOption(typeSel.select, "number", "Number");
                        typeSel.setValue(bound.type);
                        typeSel.syncFromSelect();

                        const valInp = document.createElement("input");
                        valInp.className = "fm-cf-input";
                        valInp.type = "text";
                        valInp.placeholder = "Value";
                        valInp.value = bound.value ?? "";

                        const colInp = document.createElement("input");
                        colInp.type = "color";
                        colInp.className = "fm-cf-color";
                        colInp.value = bound.color && /^#/.test(bound.color) ? bound.color : (name === "min" ? "#ffffff" : "#000000");

                        const syncEnabled = () => {
                            valInp.disabled = typeSel.getValue() !== "number";
                        };
                        syncEnabled();

                        typeSel.select.addEventListener("change", () => {
                            bound.type = typeSel.getValue() as any;
                            if (bound.type !== "number") bound.value = undefined;
                            syncEnabled();
                        });
                        valInp.addEventListener("input", () => {
                            bound.value = valInp.value;
                        });
                        colInp.addEventListener("input", () => {
                            bound.color = colInp.value;
                        });

                        row.appendChild(typeSel.host);
                        row.appendChild(valInp);
                        row.appendChild(colInp);

                        return row;
                    };

                    mkStyleField("Minimum", mkBoundEditor("min"));
                    mkStyleField("Maximum", mkBoundEditor("max"));

                    const midWrap = document.createElement("div");
                    midWrap.className = "fm-cf-style-mid";
                    const cb = document.createElement("input");
                    cb.type = "checkbox";
                    cb.checked = !!s.useMid;
                    cb.addEventListener("change", () => {
                        s.useMid = cb.checked;
                    });
                    const lbl = document.createElement("span");
                    lbl.className = "fm-cf-text";
                    lbl.textContent = " Use center color";
                    const midColor = document.createElement("input");
                    midColor.type = "color";
                    midColor.className = "fm-cf-color";
                    midColor.value = s.midColor && /^#/.test(s.midColor) ? s.midColor : "#808080";
                    midColor.addEventListener("input", () => {
                        s.midColor = midColor.value;
                    });
                    midWrap.appendChild(cb);
                    midWrap.appendChild(lbl);
                    midWrap.appendChild(midColor);
                    mkStyleField("Center", midWrap);
                }

                if (surface.formatStyle === "fieldValue") {
                    const note = document.createElement("div");
                    note.className = "fm-cf-style-note";
                    note.textContent = currentProp === "icon"
                        ? "Uses the selected field's values as icon names (e.g. ragGreen) or data:image/svg+xml/..."
                        : "Uses the selected field's values as colors (hex or color name).";
                    stylePanel.appendChild(note);
                }
            }

            const filtered = (draft.rules || []).filter(r => {
                const okTarget = currentTarget === "all" ? true : ((r.target || currentTarget) === currentTarget);
                if (!okTarget) return false;
                const rp = (r.formatProp as CFFormatProp | undefined)
                    ?? (r.style?.iconName ? "icon" : undefined)
                    ?? (r.style?.background ? "background" : (r.style?.fontColor ? "fontColor" : "background"));
                return rp === currentProp;
            });
            if (surface.formatStyle === "rules") {
                for (let i = 0; i < filtered.length; i++) {
                    renderRuleRow(filtered[i], i, filtered.length);
                }
            }
        };

        const actions = document.createElement("div");
        actions.className = "fm-cf-actions";
        panel.appendChild(actions);

        const add = document.createElement("button");
        add.type = "button";
        add.className = "fm-cf-btn";
        add.textContent = "+ New rule";
        add.addEventListener("click", () => {
            // Create within the current view (Apply to)
            const surface = getCurrentSurface();
            if (surface.formatStyle !== "rules") return;
            const t: CFRuleTarget = (currentTarget === "all") ? lastSpecificTarget : (currentTarget as CFRuleTarget);
            const newRule: CFRule = {
                id: cryptoId("r"),
                enabled: true,
                target: t,
                scope: "self",
                formatProp: currentProp,
                rowMatchMode: "any",
                rowMatchText: "",
                colMatchMode: "any",
                colMatchText: "",
                conditionType: t === "cell" ? "value" : "text",
                operator: t === "cell" ? "gte" : "contains",
                value1: t === "cell" ? "0" : "",
                value2: "",
                style: {}
            };
            if (t === "cell" && currentTarget === "cell" && basedOnField) {
                newRule.colMatchMode = "contains";
                newRule.colMatchText = basedOnField;
            }
            draft.rules.push({
                ...newRule
            });
            rerender();
        });
        actions.appendChild(add);

        const cancel = document.createElement("button");
        cancel.type = "button";
        cancel.className = "fm-cf-btn";
        cancel.textContent = "Cancel";
        cancel.addEventListener("click", () => {
            this.closeCfEditor();
            this.root.removeChild(overlay);
        });
        actions.appendChild(cancel);

        const ok = document.createElement("button");
        ok.type = "button";
        ok.className = "fm-cf-btn fm-cf-btn-primary";
        ok.textContent = "OK";
        ok.addEventListener("click", () => {
            const cleanDraft: CFConfig = {
                version: 2,
                rules: (draft.rules || []).map(r => this.normalizeCfRule(r)),
                surfaces: draft.surfaces
            };
            const json = JSON.stringify(cleanDraft);
            this.persistObjectProperties("conditionalFormatting", { rulesJson: json, showEditor: false });
            this.cfRulesJson = json;
            this.cfEditorDraft = null;
            this.root.removeChild(overlay);
            // Re-render quickly so changes apply immediately.
            // (Power BI will also invoke update after persist.)
        });
        actions.appendChild(ok);

        rerender();

        this.root.appendChild(overlay);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private clearRoot(): void {
        const children = Array.from(this.root.childNodes);
        for (const n of children) {
            const el = n as any;
            if (el && el.nodeType === 1 && (el as HTMLElement).dataset?.fmPersist === "1") {
                continue;
            }
            this.root.removeChild(n);
        }
    }

    private renderMessage(message: string, viewport: powerbi.IViewport): void {
        this.root.style.width = `${viewport.width}px`;
        this.root.style.height = `${viewport.height}px`;
        this.clearRoot();
        const el = document.createElement("div");
        el.className = "fm-message";
        el.textContent = message;
        this.root.appendChild(el);
    }

    private renderEmptyState(viewport: powerbi.IViewport, ctx: { reason: "noDataView" | "missingBindings" | "filteredOut"; missingRow?: boolean; missingValues?: boolean }): void {
        this.root.style.width = `${viewport.width}px`;
        this.root.style.height = `${viewport.height}px`;
        this.clearRoot();

        const wrap = document.createElement("div");
        wrap.className = "fm-empty";

        const card = document.createElement("div");
        card.className = "fm-empty-card";

        const title = document.createElement("div");
        title.className = "fm-empty-title";
        title.textContent = "Financial Matrix Pro";
        card.appendChild(title);

        const sub = document.createElement("div");
        sub.className = "fm-empty-sub";
        if (ctx.reason === "noDataView" || (ctx.missingRow && ctx.missingValues)) {
            sub.textContent = "Premium guide: follow the steps below.";
        } else if (ctx.missingRow) {
            sub.textContent = "Row is missing. Add at least one field to Row to start. This screen auto-hides when the visual is ready.";
        } else if (ctx.missingValues) {
            sub.textContent = "Values are missing. Add at least one measure to Values. This screen auto-hides when ready.";
        } else {
            sub.textContent = "No rows to display. Filters may have removed all data.";
        }
        card.appendChild(sub);

        const accent = document.createElement("div");
        accent.className = "fm-empty-accent";
        card.appendChild(accent);

        const scroll = document.createElement("div");
        scroll.className = "fm-empty-scroll";

        const addSection = (heading: string, lines: string[], ordered?: boolean) => {
            const box = document.createElement("div");
            box.className = "fm-empty-section";

            const h = document.createElement("div");
            h.className = "fm-empty-section-title";
            h.textContent = heading;
            box.appendChild(h);

            const list = document.createElement(ordered ? "ol" : "ul");
            list.className = "fm-empty-list";
            for (const t of lines) {
                const li = document.createElement("li");
                li.textContent = t;
                list.appendChild(li);
            }
            box.appendChild(list);
            scroll.appendChild(box);
        };

        addSection(
            "Quick start (recommended)",
            [
                "Row: add your main dimension(s) (Account, Segment, Product, etc.).",
                "Values: add at least one numeric measure (Sales, Amount, Balance, etc.).",
                "Column (optional): add one or more fields for multi-level column headers (Year > Month > …).",
                "Presets: choose a preset for instant styling; then fine-tune Header / Rows / Totals.",
                "Conditional Formatting: open editor to add Rules, Gradients, Field-value formatting, and Icons.",
                "Custom Table (optional): enable it to define Parent/Child layout and control which values appear in each row."
            ],
            true
        );

        addSection(
            "Fields (what goes where)",
            [
                "Row (required): creates the left-side hierarchy and row headers.",
                "Values (required): measures shown as columns and cell values.",
                "Column (optional): builds column grouping / hierarchy in the header.",
                "Group (optional): lets you group rows for layout/format scenarios.",
                "Format by fields (CF): bind fields you want to use in Conditional Formatting 'Based on field'.",
                "Custom table value fields / values: dedicated bindings for Custom Table mapping (optional)."
            ]
        );

        addSection(
            "Presets (instant styling)",
            [
                "Presets are safe defaults for spacing, grid visibility, zebra rows, and premium look.",
                "You can always switch to Custom for full control.",
                "Header contrast remains readable across presets (dark header + light text)."
            ]
        );

        addSection(
            "Header & Column hierarchy",
            [
                "Enable Column hierarchy to show multi-level columns (e.g., Year > Month > Measure).",
                "Use Freeze header to keep headers visible while scrolling.",
                "Control font, alignment, and background from the Header card."
            ]
        );

        addSection(
            "Rows & hierarchy",
            [
                "Hierarchy view enables expand/collapse style behavior for parent/child rows.",
                "Indent size + left padding control how deep levels are aligned.",
                "Auto aggregate parents can roll-up child values into parent totals.",
                "Blank as zero helps financial statements render consistently."
            ]
        );

        addSection(
            "Totals",
            [
                "Configure Grand totals, Subtotals, and Column totals from the Totals card.",
                "Use grid options to visually separate totals from detail rows.",
                "Combine with units/decimals settings for clean statement formatting."
            ]
        );

        addSection(
            "Conditional Formatting (CF)",
            [
                "Surfaces: apply formatting to Cells, Row headers, or Column headers.",
                "Rules: apply background/font/icon when conditions match (>, <, between, contains, etc.).",
                "Gradient: color scale across values for heatmap-style insight.",
                "Field-value formatting: drive formatting using a 'Based on field' value.",
                "Icons: show KPI icons (RAG, arrows, triangles, check/cross) based on thresholds."
            ]
        );

        addSection(
            "Custom Table (advanced layout)",
            [
                "Turn on Custom Table to design a statement layout (P&L / Balance Sheet) with custom Parent/Child rows.",
                "Map each row to a bound field/measure (Values) and choose aggregation behavior.",
                "Optionally hide selected column groups to keep the layout clean.",
                "Use 'Show value names in columns' when you want value labels visible in headers."
            ]
        );

        addSection(
            "Troubleshooting",
            [
                "If the matrix is blank: check if filters removed all data, then test by removing filters.",
                "If you see this guide: ensure at least one Row field and one measure in Values are bound.",
                "If Custom Table is ON: ensure a measure is bound to Values (and configure mapping in the editor).",
                "If formatting looks off: use 'Reset to default' in the format pane for a clean baseline."
            ]
        );

        const tip = document.createElement("div");
        tip.className = "fm-empty-tip";
        tip.textContent = "";

        scroll.appendChild(tip);
        card.appendChild(scroll);

        wrap.appendChild(card);
        this.root.appendChild(wrap);
    }

    private render(model: Model, viewport: powerbi.IViewport): void {
        this.root.style.width = `${viewport.width}px`;
        this.root.style.height = `${viewport.height}px`;

        const tableBg = this.formattingSettings.table.backgroundColor.value.value;
        this.root.style.background = tableBg || "";

        const tablePreset = (this.formattingSettings.table as any).preset?.value?.value as string | undefined;
        const showGridRaw = this.formattingSettings.table.showGrid.value;
        const horizontalGrid = ((this.formattingSettings.table as any).horizontalGrid?.value as boolean | undefined) ?? true;
        const horizontalGridColor = (((this.formattingSettings.table as any).horizontalGridColor?.value?.value as string | undefined) || "");
        const horizontalGridThickness = Math.max(0, ((this.formattingSettings.table as any).horizontalGridThickness?.value as number | undefined) ?? 1);
        const horizontalGridStyle = (((this.formattingSettings.table as any).horizontalGridStyle?.value?.value as string | undefined) || "solid");

        const verticalGrid = ((this.formattingSettings.table as any).verticalGrid?.value as boolean | undefined) ?? true;
        const verticalGridColor = (((this.formattingSettings.table as any).verticalGridColor?.value?.value as string | undefined) || "");
        const verticalGridThickness = Math.max(0, ((this.formattingSettings.table as any).verticalGridThickness?.value as number | undefined) ?? 1);
        const verticalGridStyle = (((this.formattingSettings.table as any).verticalGridStyle?.value?.value as string | undefined) || "solid");

        const totalsCard: any = (this.formattingSettings as any).totals;

        // Totals grid: prefer Totals card; fallback to legacy Table card
        const grandTotalGrid = ((totalsCard as any)?.grandTotalGrid?.value as boolean | undefined)
            ?? ((this.formattingSettings.table as any).grandTotalGrid?.value as boolean | undefined)
            ?? false;
        const grandTotalGridColor = (((totalsCard as any)?.grandTotalGridColor?.value?.value as string | undefined)
            ?? (((this.formattingSettings.table as any).grandTotalGridColor?.value?.value as string | undefined) || ""));
        const grandTotalGridThickness = Math.max(0, ((totalsCard as any)?.grandTotalGridThickness?.value as number | undefined)
            ?? ((this.formattingSettings.table as any).grandTotalGridThickness?.value as number | undefined)
            ?? 2);
        const grandTotalGridStyle = (((totalsCard as any)?.grandTotalGridStyle?.value?.value as string | undefined)
            ?? (((this.formattingSettings.table as any).grandTotalGridStyle?.value?.value as string | undefined) || "solid"));

        const subtotalGrid = ((totalsCard as any)?.subtotalGrid?.value as boolean | undefined)
            ?? ((this.formattingSettings.table as any).subtotalGrid?.value as boolean | undefined)
            ?? false;
        const subtotalGridColor = (((totalsCard as any)?.subtotalGridColor?.value?.value as string | undefined)
            ?? (((this.formattingSettings.table as any).subtotalGridColor?.value?.value as string | undefined) || ""));
        const subtotalGridThickness = Math.max(0, ((totalsCard as any)?.subtotalGridThickness?.value as number | undefined)
            ?? ((this.formattingSettings.table as any).subtotalGridThickness?.value as number | undefined)
            ?? 1);
        const subtotalGridStyle = (((totalsCard as any)?.subtotalGridStyle?.value?.value as string | undefined)
            ?? (((this.formattingSettings.table as any).subtotalGridStyle?.value?.value as string | undefined) || "solid"));

        const columnTotalGrid = ((totalsCard as any)?.columnTotalGrid?.value as boolean | undefined)
            ?? ((this.formattingSettings.table as any).columnTotalGrid?.value as boolean | undefined)
            ?? false;
        const columnTotalGridColor = (((totalsCard as any)?.columnTotalGridColor?.value?.value as string | undefined)
            ?? (((this.formattingSettings.table as any).columnTotalGridColor?.value?.value as string | undefined) || ""));
        const columnTotalGridThickness = Math.max(0, ((totalsCard as any)?.columnTotalGridThickness?.value as number | undefined)
            ?? ((this.formattingSettings.table as any).columnTotalGridThickness?.value as number | undefined)
            ?? 2);
        const columnTotalGridStyle = (((totalsCard as any)?.columnTotalGridStyle?.value?.value as string | undefined)
            ?? (((this.formattingSettings.table as any).columnTotalGridStyle?.value?.value as string | undefined) || "solid"));
        const zebraRaw = this.formattingSettings.table.zebra.value;
        const wrapText = ((this.formattingSettings.table as any).wrapText?.value as boolean | undefined) ?? true;
        // Moved to Header card as "Column hierarchy" (keep fallback for older reports if it exists).
        const columnHierarchy = (((this.formattingSettings.header as any).columnHierarchy?.value as boolean | undefined)
            ?? ((this.formattingSettings.table as any).columnHierarchy?.value as boolean | undefined)
            ?? true);

        // Presets are lightweight: they only tweak sizing/toggles (no hard-coded new colors)
        const preset = (tablePreset || "custom").toLowerCase();

        // Apply a theme class at the container level so editor overlays also match the selected preset.
        const themeKey = (() => {
            if (preset.startsWith("financeblue")) return "financeblue";
            if (preset.startsWith("financepurple")) return "financepurple";
            if (preset.startsWith("financeemerald")) return "financeemerald";
            return "";
        })();
        this.root.className = themeKey ? `fm-container fm-theme-${themeKey}` : "fm-container";

        const cfCard: any = (this.formattingSettings as any).conditionalFormatting;
        const cfEnabled = (cfCard?.enabled?.value as boolean | undefined) ?? false;
        const cfConfig = cfEnabled ? this.safeParseCfConfig(this.cfRulesJson) : { version: 1 as const, rules: [] as CFRule[] };

        const cfCellBgSurface = cfEnabled ? this.getSurfaceSettings(cfConfig as any, "cell", "background") : ({ formatStyle: "rules" } as const);
        const cfCellFontSurface = cfEnabled ? this.getSurfaceSettings(cfConfig as any, "cell", "fontColor") : ({ formatStyle: "rules" } as const);
        const cfCellIconSurface = cfEnabled ? this.getSurfaceSettings(cfConfig as any, "cell", "icon") : ({ formatStyle: "rules" } as const);
        const cfRowHeaderBgSurface = cfEnabled ? this.getSurfaceSettings(cfConfig as any, "rowHeader", "background") : ({ formatStyle: "rules" } as const);
        const cfRowHeaderFontSurface = cfEnabled ? this.getSurfaceSettings(cfConfig as any, "rowHeader", "fontColor") : ({ formatStyle: "rules" } as const);
        const cfColHeaderBgSurface = cfEnabled ? this.getSurfaceSettings(cfConfig as any, "columnHeader", "background") : ({ formatStyle: "rules" } as const);
        const cfColHeaderFontSurface = cfEnabled ? this.getSurfaceSettings(cfConfig as any, "columnHeader", "fontColor") : ({ formatStyle: "rules" } as const);
        const gradientStatsByKey = new Map<string, { min: number; max: number }>();
        const rowHeaderGradientStatsByKey = new Map<string, { min: number; max: number }>();

        const colKeyByLevelsAndLeafLabel = new Map<string, string>();
        for (const c of model.columns) {
            const ll = (c.leafLabel || "").trim();
            if (!ll) continue;
            const levelsKey = (c.columnLevels || []).join("||");
            colKeyByLevelsAndLeafLabel.set(`${levelsKey}||${ll}`, c.key);
        }

        const resolveColKeyForBasedOn = (col: BucketedColumnKey, basedOnField: string): string | null => {
            const want = (basedOnField || "").trim();
            if (!want) return col.key;
            const levelsKey = (col.columnLevels || []).join("||");
            return colKeyByLevelsAndLeafLabel.get(`${levelsKey}||${want}`) || null;
        };

        type PresetOverride = {
            showGrid?: boolean;
            zebra?: boolean;
            horizontalGrid?: boolean;
            verticalGrid?: boolean;
            horizontalGridThickness?: number;
            verticalGridThickness?: number;
            horizontalGridStyle?: string;
            verticalGridStyle?: string;
        };

        const presetOverrides: Record<string, PresetOverride> = {
            // legacy
            clean: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            compact: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: false, horizontalGridThickness: 1, horizontalGridStyle: "solid" },
            softzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },

            // Modern attractive finance themes
            financeblue: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            financebluezebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            financepurple: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            financepurplezebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            financeemerald: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            financeemeraldzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },

            // preset pack
            minimal: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            minimalzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            ledger: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: false, horizontalGridThickness: 1, horizontalGridStyle: "solid" },
            ledgerzebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: false, horizontalGridThickness: 1, horizontalGridStyle: "solid" },
            statement: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            statementzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            spreadsheet: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            spreadsheetzebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            boxed: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            boxedzebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            airy: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            airyzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            boldgrid: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 2, verticalGridThickness: 2, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            boldgridzebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 2, verticalGridThickness: 2, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            dottedgrid: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "dotted", verticalGridStyle: "dotted" },
            dashedgrid: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "dashed", verticalGridStyle: "dashed" },
            blueprint: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "dashed", verticalGridStyle: "dotted" },
            blueprintzebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "dashed", verticalGridStyle: "dotted" },
            mono: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "solid", verticalGridStyle: "solid" },
            monozebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: true, horizontalGridThickness: 1, verticalGridThickness: 1, horizontalGridStyle: "solid", verticalGridStyle: "solid" },

            // More distinct patterns
            card: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            cardzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            underlineheader: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            underlineheaderzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            borderlessdividers: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: false, horizontalGridThickness: 1, horizontalGridStyle: "solid" },
            borderlessdividerszebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: false, horizontalGridThickness: 1, horizontalGridStyle: "solid" },
            columndividers: { showGrid: true, zebra: false, horizontalGrid: false, verticalGrid: true, verticalGridThickness: 1, verticalGridStyle: "solid" },
            columndividerszebra: { showGrid: true, zebra: true, horizontalGrid: false, verticalGrid: true, verticalGridThickness: 1, verticalGridStyle: "solid" },
            groupdivider: { showGrid: true, zebra: false, horizontalGrid: true, verticalGrid: false, horizontalGridThickness: 1, horizontalGridStyle: "solid" },
            groupdividerzebra: { showGrid: true, zebra: true, horizontalGrid: true, verticalGrid: false, horizontalGridThickness: 1, horizontalGridStyle: "solid" },

            // Premium SaaS-style
            premiumsoft: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            premiumsoftzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            premiumaccent: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            premiumaccentzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false },
            premiumcompact: { showGrid: false, zebra: false, horizontalGrid: false, verticalGrid: false },
            premiumcompactzebra: { showGrid: false, zebra: true, horizontalGrid: false, verticalGrid: false }
        };

        const o = presetOverrides[preset] || {};
        const showGrid = (o.showGrid ?? showGridRaw);
        const zebra = (o.zebra ?? zebraRaw);
        const headerShow = this.formattingSettings.header.show.value;
        const headerBg = this.formattingSettings.header.backgroundColor.value.value;
        const headerColor = this.formattingSettings.header.fontColor.value.value;
        const headerFontSizeRaw = this.formattingSettings.header.fontSize.value;
        const headerFontFamily = (((this.formattingSettings.header as any).fontFamily?.value?.value as string | undefined) || "Segoe UI");
        const headerTextAlign = (((this.formattingSettings.header as any).textAlign?.value?.value as string | undefined) || "center");
        const rowHeightRaw = Math.max(16, this.formattingSettings.rows.rowHeight.value);
        const hierarchyView = ((this.formattingSettings.rows as any).hierarchyView?.value as boolean | undefined) ?? false;
        const labelPaddingLeft = Math.max(0, (this.formattingSettings.rows as any).labelPaddingLeft?.value ?? 8);

        const headerFontSize = preset === "compact" ? Math.max(10, Math.min(11, headerFontSizeRaw)) : headerFontSizeRaw;
        const rowHeight = preset === "compact" ? Math.max(16, Math.min(20, rowHeightRaw)) : rowHeightRaw;
        const indentSize = Math.max(0, this.formattingSettings.rows.indentSize.value);

        const parentFontSize = Math.max(6, ((this.formattingSettings.rows as any).parentFontSize?.value as number | undefined) ?? 12);
        const parentFontFamily = (((this.formattingSettings.rows as any).parentFontFamily?.value?.value as string | undefined) || "Segoe UI");
        const parentTextAlign = (((this.formattingSettings.rows as any).parentTextAlign?.value?.value as string | undefined) || "left");

        const childFontSize = Math.max(6, ((this.formattingSettings.rows as any).childFontSize?.value as number | undefined) ?? 12);
        const childFontFamily = (((this.formattingSettings.rows as any).childFontFamily?.value?.value as string | undefined) || "Segoe UI");
        const childTextAlign = (((this.formattingSettings.rows as any).childTextAlign?.value?.value as string | undefined) || "left");

        const decimals = Math.max(0, this.formattingSettings.numbers.decimals.value);
        const units = (this.formattingSettings.numbers.displayUnits.value.value as DisplayUnit) || "auto";
        const blankAsZero = this.formattingSettings.rows.blankAsZero.value;
        const autoAggregateParents = this.formattingSettings.rows.autoAggregateParents.value;

        const valueFontFamily = (((this.formattingSettings.numbers as any).valueFontFamily?.value?.value as string | undefined) || "Segoe UI");
        const valueFontSize = Math.max(6, ((this.formattingSettings.numbers as any).valueFontSize?.value as number | undefined) ?? 12);
        const valueTextAlign = (((this.formattingSettings.numbers as any).valueTextAlign?.value?.value as string | undefined) || "right");
        const valueFontColor = (((this.formattingSettings.numbers as any).valueFontColor?.value?.value as string | undefined) || "");
        const valueBold = (((this.formattingSettings.numbers as any).valueBold?.value as boolean | undefined) ?? false) === true;

        type TotalStyle = {
            bold: boolean;
            color: string;
            background: string;
            fontFamily: string;
            fontSize: number;
            labelAlign: string;
            valueAlign: string;
        };

        const rowsAny = (this.formattingSettings.rows as any);
        const totalsAny = (totalsCard as any) || null;
        const cleanColor = (v: any): string => {
            const s = (typeof v === "string" ? v : "") || "";
            return s.trim();
        };
        const cleanAlign = (v: any, fallback: string): string => {
            const s = (typeof v === "string" ? v : "") || "";
            const t = s.trim().toLowerCase();
            if (t === "left" || t === "center" || t === "right") return t;
            return fallback;
        };

        const legacyTotals: TotalStyle = {
            bold: (rowsAny.totalsBold?.value as boolean | undefined) ?? true,
            color: cleanColor(rowsAny.totalsFontColor?.value?.value),
            background: cleanColor(rowsAny.totalsBackground?.value?.value),
            fontFamily: (rowsAny.totalsFontFamily?.value?.value as string | undefined) || "Segoe UI",
            fontSize: Math.max(6, (rowsAny.totalsFontSize?.value as number | undefined) ?? 12),
            labelAlign: cleanAlign(rowsAny.totalsLabelAlign?.value?.value, "left"),
            valueAlign: cleanAlign(rowsAny.totalsValueAlign?.value?.value, "right")
        };

        const grandTotalStyle: TotalStyle = {
            bold: ((totalsAny?.grandTotalBold?.value as boolean | undefined) ?? (rowsAny.grandTotalBold?.value as boolean | undefined)) ?? legacyTotals.bold,
            color: (cleanColor(totalsAny?.grandTotalFontColor?.value?.value) || cleanColor(rowsAny.grandTotalFontColor?.value?.value) || legacyTotals.color),
            background: (cleanColor(totalsAny?.grandTotalBackground?.value?.value) || cleanColor(rowsAny.grandTotalBackground?.value?.value) || legacyTotals.background),
            fontFamily: ((totalsAny?.grandTotalFontFamily?.value?.value as string | undefined) || (rowsAny.grandTotalFontFamily?.value?.value as string | undefined) || legacyTotals.fontFamily),
            fontSize: Math.max(6, (totalsAny?.grandTotalFontSize?.value as number | undefined) ?? (rowsAny.grandTotalFontSize?.value as number | undefined) ?? legacyTotals.fontSize),
            labelAlign: cleanAlign((totalsAny?.grandTotalLabelAlign?.value?.value as string | undefined) ?? (rowsAny.grandTotalLabelAlign?.value?.value as string | undefined), legacyTotals.labelAlign),
            valueAlign: cleanAlign((totalsAny?.grandTotalValueAlign?.value?.value as string | undefined) ?? (rowsAny.grandTotalValueAlign?.value?.value as string | undefined), legacyTotals.valueAlign)
        };

        const subtotalStyle: TotalStyle = {
            bold: ((totalsAny?.subtotalBold?.value as boolean | undefined) ?? (rowsAny.subtotalBold?.value as boolean | undefined)) ?? legacyTotals.bold,
            color: (cleanColor(totalsAny?.subtotalFontColor?.value?.value) || cleanColor(rowsAny.subtotalFontColor?.value?.value) || legacyTotals.color),
            background: (cleanColor(totalsAny?.subtotalBackground?.value?.value) || cleanColor(rowsAny.subtotalBackground?.value?.value) || legacyTotals.background),
            fontFamily: ((totalsAny?.subtotalFontFamily?.value?.value as string | undefined) || (rowsAny.subtotalFontFamily?.value?.value as string | undefined) || legacyTotals.fontFamily),
            fontSize: Math.max(6, (totalsAny?.subtotalFontSize?.value as number | undefined) ?? (rowsAny.subtotalFontSize?.value as number | undefined) ?? legacyTotals.fontSize),
            labelAlign: cleanAlign((totalsAny?.subtotalLabelAlign?.value?.value as string | undefined) ?? (rowsAny.subtotalLabelAlign?.value?.value as string | undefined), legacyTotals.labelAlign),
            valueAlign: cleanAlign((totalsAny?.subtotalValueAlign?.value?.value as string | undefined) ?? (rowsAny.subtotalValueAlign?.value?.value as string | undefined), legacyTotals.valueAlign)
        };

        const columnTotalStyle: TotalStyle = {
            bold: ((totalsAny?.columnTotalBold?.value as boolean | undefined) ?? (rowsAny.columnTotalBold?.value as boolean | undefined)) ?? legacyTotals.bold,
            color: (cleanColor(totalsAny?.columnTotalFontColor?.value?.value) || cleanColor(rowsAny.columnTotalFontColor?.value?.value) || legacyTotals.color),
            background: (cleanColor(totalsAny?.columnTotalBackground?.value?.value) || cleanColor(rowsAny.columnTotalBackground?.value?.value) || legacyTotals.background),
            fontFamily: ((totalsAny?.columnTotalFontFamily?.value?.value as string | undefined) || (rowsAny.columnTotalFontFamily?.value?.value as string | undefined) || legacyTotals.fontFamily),
            fontSize: Math.max(6, (totalsAny?.columnTotalFontSize?.value as number | undefined) ?? (rowsAny.columnTotalFontSize?.value as number | undefined) ?? legacyTotals.fontSize),
            labelAlign: cleanAlign((totalsAny?.columnTotalLabelAlign?.value?.value as string | undefined) ?? (rowsAny.columnTotalLabelAlign?.value?.value as string | undefined), legacyTotals.labelAlign),
            valueAlign: cleanAlign((totalsAny?.columnTotalValueAlign?.value?.value as string | undefined) ?? (rowsAny.columnTotalValueAlign?.value?.value as string | undefined), legacyTotals.valueAlign)
        };

        const grandTotalLabel = (((totalsAny?.grandTotalLabel?.value as string | undefined) ?? ((this.formattingSettings.rows as any).grandTotalLabel?.value as string | undefined) ?? "Grand Total").trim()) || "Grand Total";
        const subtotalLabelTemplate = (((totalsAny?.subtotalLabelTemplate?.value as string | undefined) ?? ((this.formattingSettings.rows as any).subtotalLabelTemplate?.value as string | undefined) ?? "Total {label}").trim()) || "Total {label}";
        const columnTotalLabel = (((totalsAny?.columnTotalLabel?.value as string | undefined) ?? ((this.formattingSettings.rows as any).columnTotalLabel?.value as string | undefined) ?? "Column total").trim()) || "Column total";

        const visibleRows = model.rows.filter(r => this.isVisible(r, model.rows));

        if (cfEnabled) {
            const gradientRowKeys = new Set<string>();
            const considerRowSurface = (s: CFSurfaceSettings) => {
                if (s.formatStyle !== "gradient") return;
                const bof = ((s as CFGradientSettings).basedOnField || "").trim();
                gradientRowKeys.add(bof || "*");
            };
            considerRowSurface(cfRowHeaderBgSurface as any);
            considerRowSurface(cfRowHeaderFontSurface as any);

            const sumRow = (rowCode: string, fieldKey: string, includeZeroForBlanks: boolean): number | null => {
                const wantLeaf = fieldKey === "*" ? "" : fieldKey;
                let sum = 0;
                let hasAny = false;
                for (const col of model.columns) {
                    const ll = (col.leafLabel || "").trim();
                    if (wantLeaf && ll !== wantLeaf) continue;
                    const raw = model.values.get(rowCode)?.get(col.key) ?? null;
                    const v = raw == null ? null : raw;
                    if (v == null || !Number.isFinite(v)) {
                        if (includeZeroForBlanks) {
                            hasAny = true;
                            // add 0
                        }
                        continue;
                    }
                    hasAny = true;
                    sum += v;
                }
                return hasAny ? sum : (includeZeroForBlanks ? 0 : null);
            };

            for (const key of gradientRowKeys) {
                const includeZeroForBlanks = (
                    (cfRowHeaderBgSurface as any).formatStyle === "gradient" && (((cfRowHeaderBgSurface as any).basedOnField || "").trim() || "*") === key && (cfRowHeaderBgSurface as any).emptyValues === "zero"
                ) || (
                    (cfRowHeaderFontSurface as any).formatStyle === "gradient" && (((cfRowHeaderFontSurface as any).basedOnField || "").trim() || "*") === key && (cfRowHeaderFontSurface as any).emptyValues === "zero"
                );

                let mn: number | null = null;
                let mx: number | null = null;
                for (const r of visibleRows) {
                    if (r.type === "blank") continue;
                    if (r.isTotal) continue;
                    const v = sumRow(r.code, key, includeZeroForBlanks);
                    if (v == null || !Number.isFinite(v)) continue;
                    mn = mn == null ? v : Math.min(mn, v);
                    mx = mx == null ? v : Math.max(mx, v);
                }
                if (mn != null && mx != null) rowHeaderGradientStatsByKey.set(key, { min: mn, max: mx });
            }
        }

        if (cfEnabled) {
            const gradientLeafLabels = (() => {
                const set = new Set<string>();
                const consider = (s: CFSurfaceSettings) => {
                    if (s.formatStyle !== "gradient") return;
                    const bof = ((s as CFGradientSettings).basedOnField || "").trim();
                    if (bof) {
                        set.add(bof);
                        return;
                    }
                    for (const c of model.columns) {
                        const ll = (c.leafLabel || "").trim();
                        if (ll) set.add(ll);
                    }
                };
                consider(cfCellBgSurface as any);
                consider(cfCellFontSurface as any);
                return set;
            })();

            if (gradientLeafLabels.size > 0) {
                const includeZeroForBlanks = (
                    (cfCellBgSurface as any).formatStyle === "gradient" && ((cfCellBgSurface as any).emptyValues === "zero")
                ) || (
                    (cfCellFontSurface as any).formatStyle === "gradient" && ((cfCellFontSurface as any).emptyValues === "zero")
                );
                const minBy = new Map<string, number>();
                const maxBy = new Map<string, number>();

                const leafSet = new Set<string>();
                for (const c of model.columns) {
                    const ll = (c.leafLabel || "").trim();
                    if (ll) leafSet.add(ll);
                }
                for (const r of visibleRows) {
                    if (r.type === "blank") continue;
                    if (r.isTotal) continue;
                    for (const col of model.columns) {
                        const ll = (col.leafLabel || "").trim();

                        // 1) Displayed measures / leaf labels
                        if (ll && gradientLeafLabels.has(ll)) {
                            const raw = model.values.get(r.code)?.get(col.key) ?? null;
                            const v = raw == null ? null : raw;
                            if (v == null || !Number.isFinite(v)) {
                                if (includeZeroForBlanks) {
                                    const prevMin = minBy.get(ll);
                                    const prevMax = maxBy.get(ll);
                                    minBy.set(ll, prevMin == null ? 0 : Math.min(prevMin, 0));
                                    maxBy.set(ll, prevMax == null ? 0 : Math.max(prevMax, 0));
                                }
                            } else {
                                const prevMin = minBy.get(ll);
                                const prevMax = maxBy.get(ll);
                                minBy.set(ll, prevMin == null ? v : Math.min(prevMin, v));
                                maxBy.set(ll, prevMax == null ? v : Math.max(prevMax, v));
                            }
                        }

                        // 2) Non-displayed based-on measures (e.g., formatBy)
                        // For Custom Table tuple-only columns (no measure), include leafLabel to uniquely identify the column cell.
                        const levelsKeyBase = (col.columnLevels || []).join("||");
                        const levelsKey = (col as any).measure == null
                            ? [...(col.columnLevels || []), (col.leafLabel || "")].filter(p => !!p).join("||")
                            : levelsKeyBase;
                        for (const key of gradientLeafLabels) {
                            if (!key || leafSet.has(key)) continue;
                            const rawAny = model.cfFieldRaw.get(`${r.code}||${levelsKey}||${key}`);
                            const num = toNumber(rawAny);
                            if (!Number.isFinite(num)) {
                                if (!includeZeroForBlanks) continue;
                                const prevMin = minBy.get(key);
                                const prevMax = maxBy.get(key);
                                minBy.set(key, prevMin == null ? 0 : Math.min(prevMin, 0));
                                maxBy.set(key, prevMax == null ? 0 : Math.max(prevMax, 0));
                                continue;
                            }
                            const prevMin = minBy.get(key);
                            const prevMax = maxBy.get(key);
                            minBy.set(key, prevMin == null ? num : Math.min(prevMin, num));
                            maxBy.set(key, prevMax == null ? num : Math.max(prevMax, num));
                        }
                    }
                }
                for (const ll of gradientLeafLabels) {
                    const mn = minBy.get(ll);
                    const mx = maxBy.get(ll);
                    if (mn == null || mx == null) continue;
                    gradientStatsByKey.set(ll, { min: mn, max: mx });
                }
            }
        }

        const unitScaleByCol = new Map<string, number>();
        for (const col of model.columns) {
            unitScaleByCol.set(col.key, this.computeUnitScale(units, model, visibleRows, col.key));
        }

        this.clearRoot();

        const table = document.createElement("table");
        const colWidthModeRaw = ((this.formattingSettings.table as any).columnWidthMode?.value?.value as string | undefined) || "auto";
        const colWidthMode = (colWidthModeRaw || "auto").trim().toLowerCase();
        const isManualWidth = colWidthMode === "manual";
        const effHorizontalGrid = (o.horizontalGrid ?? horizontalGrid);
        const effVerticalGrid = (o.verticalGrid ?? verticalGrid);
        const effHorizontalThickness = Math.max(0, (o.horizontalGridThickness ?? horizontalGridThickness));
        const effVerticalThickness = Math.max(0, (o.verticalGridThickness ?? verticalGridThickness));
        const effHorizontalStyle = (o.horizontalGridStyle ?? horizontalGridStyle);
        const effVerticalStyle = (o.verticalGridStyle ?? verticalGridStyle);

        const freezeHeader = ((this.formattingSettings.table as any).freezeHeader?.value as boolean | undefined) ?? false;
        const freezeRowHeaders = ((this.formattingSettings.table as any).freezeRowHeaders?.value as boolean | undefined) ?? false;

        const presetClass = preset ? `fm-preset-${preset}` : "";
        const gridHClass = showGrid && effHorizontalGrid && effHorizontalThickness > 0 ? "fm-grid-h" : "";
        const gridVClass = showGrid && effVerticalGrid && effVerticalThickness > 0 ? "fm-grid-v" : "";
        table.className = `fm-table ${presetClass} ${showGrid ? "fm-grid" : ""} ${gridHClass} ${gridVClass} ${isManualWidth ? "fm-width-manual" : ""} ${(freezeHeader || freezeRowHeaders) ? "fm-freeze" : ""} ${wrapText ? "fm-wrap" : ""}`;
        if (this.activeSelectionKeys && this.activeSelectionKeys.size > 0) {
            table.classList.add("fm-has-selection");
        }

        const defaultGridColor = "var(--fm-grid-color, rgba(0, 0, 0, 0.12))";
        if (showGrid && effHorizontalGrid && effHorizontalThickness > 0) {
            table.style.setProperty("--fm-grid-h", `${effHorizontalThickness}px ${effHorizontalStyle} ${horizontalGridColor || defaultGridColor}`);
        }
        if (showGrid && effVerticalGrid && effVerticalThickness > 0) {
            table.style.setProperty("--fm-grid-v", `${effVerticalThickness}px ${effVerticalStyle} ${verticalGridColor || defaultGridColor}`);
        }

        const mkBorder = (thickness: number, style: string, color: string): string => {
            const t = Math.max(0, thickness);
            if (t <= 0) return "";
            return `${t}px ${style || "solid"} ${color || defaultGridColor}`;
        };

        const grandRowBorder = mkBorder(grandTotalGridThickness, grandTotalGridStyle, grandTotalGridColor);
        const subRowBorder = mkBorder(subtotalGridThickness, subtotalGridStyle, subtotalGridColor);
        const colTotalBorder = mkBorder(columnTotalGridThickness, columnTotalGridStyle, columnTotalGridColor);

        const columnWidthPx = Math.max(60, ((this.formattingSettings.table as any).columnWidthPx?.value as number | undefined) ?? 120);
        // In manual-width mode, this is treated as *per row-header column* width.
        const rowHeaderColWidthPx = Math.max(60, ((this.formattingSettings.table as any).rowHeaderWidthPx?.value as number | undefined) ?? 220);

        // Manual widths can be flaky with colspan headers; colgroup makes widths reliable.
        if (isManualWidth) {
            const colgroup = document.createElement("colgroup");
            for (let i = 0; i < model.rowHeaderColumns.length; i++) {
                const colEl = document.createElement("col");
                colEl.style.width = `${rowHeaderColWidthPx}px`;
                colEl.style.minWidth = `${rowHeaderColWidthPx}px`;
                colEl.style.maxWidth = `${rowHeaderColWidthPx}px`;
                colgroup.appendChild(colEl);
            }
            for (let i = 0; i < model.columns.length; i++) {
                const colEl = document.createElement("col");
                colEl.style.width = `${columnWidthPx}px`;
                colEl.style.minWidth = `${columnWidthPx}px`;
                colEl.style.maxWidth = `${columnWidthPx}px`;
                colgroup.appendChild(colEl);
            }
            if (model.showColumnTotal) {
                const colEl = document.createElement("col");
                colEl.style.width = `${columnWidthPx}px`;
                colEl.style.minWidth = `${columnWidthPx}px`;
                colEl.style.maxWidth = `${columnWidthPx}px`;
                colgroup.appendChild(colEl);
            }
            table.appendChild(colgroup);

            // Force the table to be at least the sum of manual column widths.
            // This prevents the browser from squeezing columns to fit the viewport,
            // and makes overflow-x show a scrollbar when the table is wider.
            const dataCols = model.columns.length + (model.showColumnTotal ? 1 : 0);
            const totalWidthPx = (model.rowHeaderColumns.length * rowHeaderColWidthPx) + (dataCols * columnWidthPx);
            if (Number.isFinite(totalWidthPx) && totalWidthPx > 0) {
                table.style.width = `${totalWidthPx}px`;
                table.style.minWidth = "100%";
            }
        }

        if (headerShow) {
            const thead = document.createElement("thead");
            // Column hierarchy ON  => single compact header row: "Year Country Measure"
            // Column hierarchy OFF => stacked rows: Column1 row, Column2 row, ..., Value row (with parent merging)
            const headerLevels = columnHierarchy
                ? 1
                : Math.max(1, model.columns.reduce((m, c) => Math.max(m, (c.columnLevels.length + 1)), 1));

            const headerText = (c: BucketedColumnKey, level: number): string => {
                if (columnHierarchy) {
                    // Single row only
                    const parts = [...(c.columnLevels || []), c.leafLabel].filter(p => !!p);
                    return parts.join(" ");
                }
                if (level < c.columnLevels.length) return c.columnLevels[level] ?? "";
                if (level === c.columnLevels.length) return c.leafLabel;
                return "";
            };

            const samePrefix = (a: BucketedColumnKey, b: BucketedColumnKey, level: number): boolean => {
                for (let i = 0; i <= level; i++) {
                    if (headerText(a, i) !== headerText(b, i)) return false;
                }
                return true;
            };

            const sumGroup = (startIdx: number, endIdx: number, leafLabelFilter: string, includeZeroForBlanks: boolean): number | null => {
                const wantLeaf = (leafLabelFilter || "").trim();
                let sum = 0;
                let hasAny = false;
                for (const r of visibleRows) {
                    if (r.type === "blank") continue;
                    if (r.isTotal) continue;
                    for (let ci = startIdx; ci < endIdx; ci++) {
                        const col = model.columns[ci];
                        const ll = (col.leafLabel || "").trim();
                        if (wantLeaf && ll !== wantLeaf) continue;
                        const raw = model.values.get(r.code)?.get(col.key) ?? null;
                        const v = raw == null ? null : raw;
                        if (v == null || !Number.isFinite(v)) {
                            if (includeZeroForBlanks) {
                                hasAny = true;
                            }
                            continue;
                        }
                        hasAny = true;
                        sum += v;
                    }
                }
                return hasAny ? sum : (includeZeroForBlanks ? 0 : null);
            };

            const findGroupColor = (startIdx: number, endIdx: number, leafLabelFilter: string): string | null => {
                const wantLeaf = (leafLabelFilter || "").trim();
                for (let ci = startIdx; ci < endIdx; ci++) {
                    const col = model.columns[ci];
                    const ll = (col.leafLabel || "").trim();
                    if (wantLeaf && ll !== wantLeaf) continue;
                    for (const r of visibleRows) {
                        if (r.type === "blank") continue;
                        if (r.isTotal) continue;
                        const raw = model.valuesRaw.get(r.code)?.get(col.key);
                        const parsed = this.parseSafeCssColor(raw);
                        if (parsed) return parsed;
                    }
                }
                return null;
            };

            const gradientColorFor = (s: CFGradientSettings, v: number, autoMin: number, autoMax: number): string => {
                const parseBound = (b: CFGradientBound, fallback: number): number => {
                    if (b.type === "number") {
                        const n = Number((b.value || "").trim());
                        return Number.isFinite(n) ? n : fallback;
                    }
                    return fallback;
                };
                const minVal = parseBound(s.min, autoMin);
                const maxVal = parseBound(s.max, autoMax);
                const denom = (maxVal - minVal);
                if (!Number.isFinite(denom) || denom === 0) return s.max.color;
                const t = this.clamp01((v - minVal) / denom);
                if (s.useMid) {
                    const midT = 0.5;
                    if (t <= midT) return this.lerpColorHex(s.min.color, s.midColor, t / midT);
                    return this.lerpColorHex(s.midColor, s.max.color, (t - midT) / (1 - midT));
                }
                return this.lerpColorHex(s.min.color, s.max.color, t);
            };

            for (let level = 0; level < headerLevels; level++) {
                const tr = document.createElement("tr");
                tr.className = "fm-header";
                if (headerBg) tr.style.background = headerBg;
                if (headerColor) tr.style.color = headerColor;
                if (headerFontSize) tr.style.fontSize = `${headerFontSize}px`;
                if (headerFontFamily) tr.style.fontFamily = headerFontFamily;

                if (level === 0) {
                    for (let rhIndex = 0; rhIndex < model.rowHeaderColumns.length; rhIndex++) {
                        const rh = model.rowHeaderColumns[rhIndex];
                        const thRow = document.createElement("th");
                        thRow.className = "fm-th fm-th-row";
                        thRow.textContent = rh;
                        thRow.rowSpan = headerLevels;
                        thRow.style.textAlign = headerTextAlign as any;
                        // Note: Freeze is handled via scroll-synced transforms (more reliable than sticky in Power BI).
                        if (freezeRowHeaders) {
                            thRow.classList.add("fm-freeze-left");
                        }
                        if (isManualWidth) {
                            thRow.style.width = `${rowHeaderColWidthPx}px`;
                            thRow.style.minWidth = `${rowHeaderColWidthPx}px`;
                            thRow.style.maxWidth = `${rowHeaderColWidthPx}px`;
                        }
                        tr.appendChild(thRow);
                    }
                }

                // Merge repeated parent labels (e.g., Year) using prefix grouping.
                const groups: Array<{ start: number; end: number; span: number; current: BucketedColumnKey; text: string }> = [];
                let i = 0;
                while (i < model.columns.length) {
                    const start = i;
                    const current = model.columns[i];
                    while (i < model.columns.length && samePrefix(model.columns[start], model.columns[i], level)) {
                        i++;
                    }
                    const end = i;
                    const span = end - start;
                    groups.push({ start, end, span, current, text: headerText(current, level) });
                }

                const computeAutoMinMax = (surface: CFSurfaceSettings): { min: number; max: number } | null => {
                    if (surface.formatStyle !== "gradient") return null;
                    const s = surface as CFGradientSettings;
                    const includeZero = s.emptyValues === "zero";
                    let mn: number | null = null;
                    let mx: number | null = null;
                    for (const g of groups) {
                        const v = sumGroup(g.start, g.end, s.basedOnField || "", includeZero);
                        if (v == null || !Number.isFinite(v)) continue;
                        mn = mn == null ? v : Math.min(mn, v);
                        mx = mx == null ? v : Math.max(mx, v);
                    }
                    if (mn == null || mx == null) return null;
                    return { min: mn, max: mx };
                };

                const bgAuto = computeAutoMinMax(cfColHeaderBgSurface as any);
                const fontAuto = computeAutoMinMax(cfColHeaderFontSurface as any);

                for (const g of groups) {
                    const th = document.createElement("th");
                    th.className = "fm-th";
                    th.colSpan = g.span;
                    th.textContent = g.text;
                    th.style.textAlign = headerTextAlign as any;

                    if (cfEnabled) {
                        const bgStyle = this.evalCfRules(cfConfig, {
                            target: "columnHeader",
                            formatProp: "background",
                            colKey: g.current.key,
                            colText: g.text,
                            text: g.text
                        });
                        if (bgStyle) this.applyCfStyle(th, bgStyle);

                        const fontStyle = this.evalCfRules(cfConfig, {
                            target: "columnHeader",
                            formatProp: "fontColor",
                            colKey: g.current.key,
                            colText: g.text,
                            text: g.text
                        });
                        if (fontStyle) this.applyCfStyle(th, fontStyle);

                        // Surface styles for headers: Gradient / Field value.
                        const applySurface = (surface: CFSurfaceSettings, prop: CFFormatProp) => {
                            if (surface.formatStyle === "fieldValue") {
                                const s = surface as CFFieldValueSettings;
                                const based = ((s.basedOnField || "").trim());
                                let rawAny: any = null;
                                if (based) {
                                    rawAny = model.colFieldRaw.get(g.current.key)?.get(based);
                                    if (rawAny == null) rawAny = model.colFieldRaw.get(g.current.columnLevels.join("||"))?.get(based);
                                }
                                const fieldColor = this.parseSafeCssColor(rawAny);
                                const textColor = this.parseSafeCssColor(g.text);
                                const color = fieldColor || textColor || findGroupColor(g.start, g.end, s.basedOnField || "");
                                if (!color) return;
                                if (prop === "background") th.style.background = color;
                                else th.style.color = color;
                                return;
                            }
                            if (surface.formatStyle === "gradient") {
                                const s = surface as CFGradientSettings;
                                const includeZero = s.emptyValues === "zero";
                                const gv = sumGroup(g.start, g.end, s.basedOnField || "", includeZero);
                                const v = gv == null ? (includeZero ? 0 : null) : gv;
                                if (v == null || !Number.isFinite(v)) return;
                                const auto = prop === "background" ? bgAuto : fontAuto;
                                const autoMin = auto ? auto.min : v;
                                const autoMax = auto ? auto.max : v;
                                const c = gradientColorFor(s, v, autoMin, autoMax);
                                if (prop === "background") th.style.background = c;
                                else th.style.color = c;
                            }
                        };

                        applySurface(cfColHeaderBgSurface as any, "background");
                        applySurface(cfColHeaderFontSurface as any, "fontColor");
                    }

                    // Note: Freeze is handled via scroll-synced transforms.
                    if (isManualWidth) {
                        const px = Math.max(1, g.span) * columnWidthPx;
                        th.style.minWidth = `${px}px`;
                        th.style.width = `${px}px`;
                    }
                    tr.appendChild(th);
                }

                // Column total header
                if (model.showColumnTotal) {
                    if (columnHierarchy) {
                        const thTotal = document.createElement("th");
                        thTotal.className = "fm-th";
                        thTotal.textContent = columnTotalLabel;
                        thTotal.style.textAlign = columnTotalStyle.labelAlign as any;
                        if (columnTotalStyle.fontFamily) thTotal.style.fontFamily = columnTotalStyle.fontFamily;
                        if (columnTotalStyle.fontSize) thTotal.style.fontSize = `${columnTotalStyle.fontSize}px`;
                        if (columnTotalStyle.bold) thTotal.style.fontWeight = "700";
                        if (columnTotalStyle.color) thTotal.style.color = columnTotalStyle.color;
                        if (columnTotalStyle.background) thTotal.style.background = columnTotalStyle.background;
                        if (columnTotalGrid && colTotalBorder) {
                            thTotal.style.borderLeft = colTotalBorder;
                            thTotal.style.borderRight = colTotalBorder;
                        }
                        // Note: Freeze is handled via scroll-synced transforms.
                        if (isManualWidth) {
                            thTotal.style.minWidth = `${columnWidthPx}px`;
                            thTotal.style.width = `${columnWidthPx}px`;
                        }
                        tr.appendChild(thTotal);
                    } else if (level === 0) {
                        const thTotal = document.createElement("th");
                        thTotal.className = "fm-th";
                        thTotal.textContent = columnTotalLabel;
                        thTotal.rowSpan = headerLevels;
                        thTotal.style.textAlign = columnTotalStyle.labelAlign as any;
                        if (columnTotalStyle.fontFamily) thTotal.style.fontFamily = columnTotalStyle.fontFamily;
                        if (columnTotalStyle.fontSize) thTotal.style.fontSize = `${columnTotalStyle.fontSize}px`;
                        if (columnTotalStyle.bold) thTotal.style.fontWeight = "700";
                        if (columnTotalStyle.color) thTotal.style.color = columnTotalStyle.color;
                        if (columnTotalStyle.background) thTotal.style.background = columnTotalStyle.background;
                        if (columnTotalGrid && colTotalBorder) {
                            thTotal.style.borderLeft = colTotalBorder;
                            thTotal.style.borderRight = colTotalBorder;
                        }
                        // Note: Freeze is handled via scroll-synced transforms.
                        if (isManualWidth) {
                            thTotal.style.minWidth = `${columnWidthPx}px`;
                            thTotal.style.width = `${columnWidthPx}px`;
                        }
                        tr.appendChild(thTotal);
                    }
                }

                thead.appendChild(tr);
            }

            table.appendChild(thead);
        }

        const tbody = document.createElement("tbody");
        let zebraIndex = 0;

        const tooltipMeasures = Array.isArray(model.tooltipMeasureNames) ? model.tooltipMeasureNames : [];

        for (const r of visibleRows) {
            const isParent = r.children.length > 0;
            const hasToggle = isParent && r.type !== "blank";
            const isCollapsed = this.collapsed.has(r.code);
            const indentPx = hierarchyView ? (r.depth * indentSize) : (model.rowFieldCount > 1 ? 0 : (r.depth * indentSize));

            const rowSelectionId = model.rowSelectionIdByCode?.get(r.code) ?? null;
            const rowSelectionKey = rowSelectionId && typeof (rowSelectionId as any).getKey === "function"
                ? (rowSelectionId as any).getKey()
                : "";
            const hasSelection = !!this.activeSelectionKeys && this.activeSelectionKeys.size > 0;
            const isSelected = hasSelection && !!rowSelectionKey && this.activeSelectionKeys.has(rowSelectionKey);

            const isGrand = r.code === "__grand_total";
            const isSub = r.code.endsWith("||__subtotal");
            const rowTotalStyle: TotalStyle = isGrand ? grandTotalStyle : (isSub ? subtotalStyle : legacyTotals);

            const baseStyle = this.getRowStyle(r, isParent);
            const tr = document.createElement("tr");
            tr.className = `fm-tr ${showGrid ? "fm-grid" : ""}`;
            tr.dataset.code = r.code;
            if (hasSelection && r.type !== "blank") {
                if (isSelected) tr.classList.add("fm-row-selected");
                else tr.classList.add("fm-row-dim");
            }
            if (!wrapText) tr.style.height = `${rowHeight}px`;
            else tr.style.minHeight = `${rowHeight}px`;
            if (r.isTotal) {
                if (rowTotalStyle.bold) tr.style.fontWeight = "700";
                if (rowTotalStyle.color) tr.style.color = rowTotalStyle.color;
                else if (baseStyle.color) tr.style.color = baseStyle.color;
                if (rowTotalStyle.background) tr.style.background = rowTotalStyle.background;
                else if (baseStyle.background) tr.style.background = baseStyle.background;
            } else {
                if (baseStyle.color) tr.style.color = baseStyle.color;
                if (baseStyle.background) tr.style.background = baseStyle.background;
            }

            if (zebra && r.type !== "blank") {
                if (zebraIndex % 2 === 1) {
                    tr.classList.add("fm-zebra");
                }
                zebraIndex++;
            }

            const activeRowHeaderIndex = (() => {
                if (hierarchyView) {
                    if (model.hasGroup) {
                        if ((r.rowLevel ?? undefined) === -1) return 0;
                        return 1;
                    }
                    return 0;
                }

                if (model.hasGroup) {
                    if ((r.rowLevel ?? undefined) === -1) return 0;
                    if ((r.rowLevel ?? undefined) != null && (r.rowLevel as number) >= 0) return 1 + (r.rowLevel as number);
                    return 0;
                }
                if ((r.rowLevel ?? undefined) != null && (r.rowLevel as number) >= 0) return (r.rowLevel as number);
                return 0;
            })();

            for (let h = 0; h < model.rowHeaderColumns.length; h++) {
                const td = document.createElement("td");
                td.className = "fm-td fm-td-row";

                // Make the entire row-header area clickable to pass filter context.
                if (rowSelectionId && r.type !== "blank" && !r.isTotal) {
                    td.style.cursor = "pointer";
                    const onPointerDown = (ev: PointerEvent) => {
                        // Left click / primary pointer only
                        if ((ev as any).button != null && (ev as any).button !== 0) return;
                        const target = ev.target as HTMLElement | null;
                        if (target && target.closest(".fm-toggle")) return;
                        this.selectRow(rowSelectionId, false);
                    };
                    // Capture-phase to avoid host/inner handlers swallowing the click.
                    td.addEventListener("pointerdown", onPointerDown as any, true);

                    td.addEventListener("click", (ev) => {
                        const target = ev.target as HTMLElement | null;
                        if (target && target.closest(".fm-toggle")) return;
                        this.selectRow(rowSelectionId, false);
                    });
                    td.addEventListener("contextmenu", (ev) => {
                        ev.preventDefault();
                        try {
                            (this.selectionManager as any).showContextMenu(rowSelectionId, { x: ev.clientX, y: ev.clientY });
                        } catch {
                            // ignore context menu errors
                        }
                    });
                }
                if (r.isTotal) {
                    const isGrand = r.code === "__grand_total";
                    const isSub = r.code.endsWith("||__subtotal");
                    if (isGrand && grandTotalGrid && grandRowBorder) {
                        td.style.borderTop = grandRowBorder;
                        td.style.borderBottom = grandRowBorder;
                    }
                    if (isSub && subtotalGrid && subRowBorder) {
                        td.style.borderTop = subRowBorder;
                        td.style.borderBottom = subRowBorder;
                    }
                }
                if (isManualWidth) {
                    td.style.width = `${rowHeaderColWidthPx}px`;
                    td.style.minWidth = `${rowHeaderColWidthPx}px`;
                    td.style.maxWidth = `${rowHeaderColWidthPx}px`;
                }

                if (freezeRowHeaders) {
                    td.classList.add("fm-freeze-left");
                }

                if (h === activeRowHeaderIndex) {
                    const rowFontFamily = (r.fontFamilyOverride || (isParent ? parentFontFamily : childFontFamily));
                    const rowFontSize = (r.fontSizeOverride != null ? r.fontSizeOverride : (isParent ? parentFontSize : childFontSize));
                    const rowTextAlign = (isParent ? parentTextAlign : childTextAlign) as any;
                    if (r.isTotal) {
                        if (rowTotalStyle.fontFamily) td.style.fontFamily = rowTotalStyle.fontFamily;
                        if (rowTotalStyle.fontSize) td.style.fontSize = `${rowTotalStyle.fontSize}px`;
                        td.style.textAlign = rowTotalStyle.labelAlign as any;
                    } else {
                        if (rowFontFamily) td.style.fontFamily = rowFontFamily;
                        if (rowFontSize) td.style.fontSize = `${rowFontSize}px`;
                        td.style.textAlign = rowTextAlign;
                    }

                    // Keep +/- aligned across rows: apply base padding to the cell,
                    // and indent only the label (not the toggle/spacer).
                    td.style.paddingLeft = `${labelPaddingLeft}px`;
                    if (r.isTotal) {
                        if (rowTotalStyle.bold) td.style.fontWeight = "700";
                    } else if (baseStyle.bold) {
                        td.style.fontWeight = "700";
                    }
                    if (baseStyle.italic) td.style.fontStyle = "italic";
                    if (baseStyle.underline) td.style.textDecoration = "underline";

                    const rowCell = document.createElement("div");
                    rowCell.className = "fm-rowcell";
                    td.appendChild(rowCell);

                    // Hierarchy view shows +/- and reserves spacing; matrix view hides +/- completely.
                    if (hierarchyView && r.type !== "blank") {
                        if (hasToggle) {
                            const btn = document.createElement("button");
                            btn.className = "fm-toggle";
                            btn.type = "button";
                            btn.setAttribute("aria-label", isCollapsed ? "Expand" : "Collapse");
                            btn.textContent = isCollapsed ? "+" : "-";
                            btn.addEventListener("click", (e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                if (this.collapsed.has(r.code)) this.collapsed.delete(r.code);
                                else this.collapsed.add(r.code);
                                this.render(model, viewport);
                            });
                            rowCell.appendChild(btn);
                        } else {
                            const spacer = document.createElement("span");
                            spacer.className = "fm-toggle-spacer";
                            rowCell.appendChild(spacer);
                        }
                    }

                    const label = document.createElement("span");
                    label.className = "fm-label";
                    label.style.marginLeft = `${indentPx}px`;
                    label.textContent = r.type === "blank" ? "" : r.label;
                    rowCell.appendChild(label);

                    // Tooltip + drill-through context menu on row header.
                    if (rowSelectionId && r.type !== "blank" && !r.isTotal) {
                        td.addEventListener("click", (ev) => {
                            const target = ev.target as HTMLElement | null;
                            if (target && target.closest(".fm-toggle")) return;
                            this.selectRow(rowSelectionId, false);
                        });
                        td.addEventListener("contextmenu", (ev) => {
                            ev.preventDefault();
                            try {
                                (this.selectionManager as any).showContextMenu(rowSelectionId, { x: ev.clientX, y: ev.clientY });
                            } catch {
                                // ignore context menu errors
                            }
                        });
                    }

                    if (r.type !== "blank") {
                        td.addEventListener("mouseenter", (ev) => {
                            const items = buildMatrixTooltipItems({ rowLabel: r.label });
                            this.showTooltip(ev as MouseEvent, items, rowSelectionId);
                        });
                        td.addEventListener("mousemove", (ev) => {
                            const items = buildMatrixTooltipItems({ rowLabel: r.label });
                            this.showTooltip(ev as MouseEvent, items, rowSelectionId);
                        });
                        td.addEventListener("mouseleave", () => {
                            this.hideTooltip(true);
                        });
                    }

                    if (cfEnabled) {
                        const bgStyle = this.evalCfRules(cfConfig, {
                            target: "rowHeader",
                            formatProp: "background",
                            rowCode: r.code,
                            rowLabel: r.label,
                            text: r.label
                        });
                        if (bgStyle) this.applyCfStyle(td, bgStyle);

                        const fontStyle = this.evalCfRules(cfConfig, {
                            target: "rowHeader",
                            formatProp: "fontColor",
                            rowCode: r.code,
                            rowLabel: r.label,
                            text: r.label
                        });
                        if (fontStyle) this.applyCfStyle(td, fontStyle);

                        const findRowColor = (leafLabelFilter: string): string | null => {
                            const wantLeaf = (leafLabelFilter || "").trim();
                            for (const col of model.columns) {
                                const ll = (col.leafLabel || "").trim();
                                if (wantLeaf && ll !== wantLeaf) continue;
                                const raw = model.valuesRaw.get(r.code)?.get(col.key);
                                const parsed = this.parseSafeCssColor(raw);
                                if (parsed) return parsed;
                            }
                            return null;
                        };

                        const sumRow = (fieldKey: string, includeZeroForBlanks: boolean): number | null => {
                            const wantLeaf = fieldKey === "*" ? "" : (fieldKey || "").trim();
                            let sum = 0;
                            let hasAny = false;
                            for (const col of model.columns) {
                                const ll = (col.leafLabel || "").trim();
                                if (wantLeaf && ll !== wantLeaf) continue;
                                const raw = model.values.get(r.code)?.get(col.key) ?? null;
                                const v = raw == null ? null : raw;
                                if (v == null || !Number.isFinite(v)) {
                                    if (includeZeroForBlanks) {
                                        hasAny = true;
                                    }
                                    continue;
                                }
                                hasAny = true;
                                sum += v;
                            }
                            return hasAny ? sum : (includeZeroForBlanks ? 0 : null);
                        };

                        const applySurface = (surface: CFSurfaceSettings, prop: CFFormatProp) => {
                            const excluded = (surface as any)?.excludedRowCodes as string[] | undefined;
                            if (Array.isArray(excluded) && excluded.includes(r.code)) return;
                            if (surface.formatStyle === "fieldValue") {
                                const s = surface as CFFieldValueSettings;
                                const based = ((s.basedOnField || "").trim());
                                const rawAny = based ? (model.rowFieldRaw.get(r.code)?.get(based) ?? null) : null;
                                const fieldColor = this.parseSafeCssColor(rawAny);
                                const textColor = this.parseSafeCssColor(r.label);
                                const color = fieldColor || textColor || findRowColor(s.basedOnField || "");
                                if (!color) return;
                                if (prop === "background") td.style.background = color;
                                else td.style.color = color;
                                return;
                            }
                            if (surface.formatStyle === "gradient") {
                                const s = surface as CFGradientSettings;
                                const key = ((s.basedOnField || "").trim()) || "*";
                                const stats = rowHeaderGradientStatsByKey.get(key);
                                if (!stats) return;
                                const includeZero = s.emptyValues === "zero";
                                const v = sumRow(key, includeZero);
                                if (v == null || !Number.isFinite(v)) return;

                                const parseBound = (b: CFGradientBound, fallback: number): number => {
                                    if (b.type === "number") {
                                        const n = Number((b.value || "").trim());
                                        return Number.isFinite(n) ? n : fallback;
                                    }
                                    return fallback;
                                };
                                const minVal = parseBound(s.min, stats.min);
                                const maxVal = parseBound(s.max, stats.max);
                                const denom = (maxVal - minVal);
                                if (!Number.isFinite(denom) || denom === 0) {
                                    const c = s.max.color;
                                    if (prop === "background") td.style.background = c;
                                    else td.style.color = c;
                                    return;
                                }

                                const t = this.clamp01((v - minVal) / denom);
                                const c = s.useMid
                                    ? (t <= 0.5
                                        ? this.lerpColorHex(s.min.color, s.midColor, t / 0.5)
                                        : this.lerpColorHex(s.midColor, s.max.color, (t - 0.5) / 0.5))
                                    : this.lerpColorHex(s.min.color, s.max.color, t);
                                if (prop === "background") td.style.background = c;
                                else td.style.color = c;
                            }
                        };

                        applySurface(cfRowHeaderBgSurface as any, "background");
                        applySurface(cfRowHeaderFontSurface as any, "fontColor");
                    }
                }

                tr.appendChild(td);
            }

            for (const col of model.columns) {
                const td = document.createElement("td");
                td.className = "fm-td fm-td-num";
                if (rowSelectionId && r.type !== "blank" && !r.isTotal) {
                    td.style.cursor = "pointer";
                    td.addEventListener("click", () => {
                        this.selectRow(rowSelectionId, false);
                    });
                }
                if (r.isTotal) {
                    const isGrand = r.code === "__grand_total";
                    const isSub = r.code.endsWith("||__subtotal");
                    if (isGrand && grandTotalGrid && grandRowBorder) {
                        td.style.borderTop = grandRowBorder;
                        td.style.borderBottom = grandRowBorder;
                    }
                    if (isSub && subtotalGrid && subRowBorder) {
                        td.style.borderTop = subRowBorder;
                        td.style.borderBottom = subRowBorder;
                    }
                }
                if (r.isTotal) {
                    if (rowTotalStyle.fontFamily) td.style.fontFamily = rowTotalStyle.fontFamily;
                    if (rowTotalStyle.fontSize) td.style.fontSize = `${rowTotalStyle.fontSize}px`;
                    td.style.textAlign = rowTotalStyle.valueAlign as any;
                    if (rowTotalStyle.bold) td.style.fontWeight = "700";
                    if (rowTotalStyle.color) td.style.color = rowTotalStyle.color;
                    if (rowTotalStyle.background) td.style.background = rowTotalStyle.background;
                } else {
                    if (valueFontFamily) td.style.fontFamily = valueFontFamily;
                    if (valueFontSize) td.style.fontSize = `${valueFontSize}px`;
                    td.style.textAlign = valueTextAlign as any;
                    if (valueBold) td.style.fontWeight = "700";
                    if (valueFontColor) td.style.color = valueFontColor;
                }
                if (isManualWidth) {
                    td.style.width = `${columnWidthPx}px`;
                    td.style.minWidth = `${columnWidthPx}px`;
                    td.style.maxWidth = `${columnWidthPx}px`;
                }
                const rowVals = model.values.get(r.code);
                const hasMappedValue = !!rowVals && rowVals.has(col.key);
                const raw = hasMappedValue ? (rowVals!.get(col.key) ?? null) : null;
                const canBlankAsZero = blankAsZero && (autoAggregateParents || !isParent) && hasMappedValue;
                const v = raw == null ? (canBlankAsZero ? 0 : null) : raw;
                const scale = unitScaleByCol.get(col.key) ?? 1;
                const text = r.type === "blank" ? "" : (units === "auto" ? this.formatAutoNumber(v) : this.formatNumber(v, decimals, scale));
                td.textContent = text;

                // Tooltip + drill-through context menu on value cells.
                if (r.type !== "blank") {
                    const colLabel = [...(col.columnLevels || []), col.leafLabel].filter(p => !!p).join(" ");
                    const levelsKey = (col.columnLevels || []).join("||");
                    const getExtra = () => {
                        const extra: Array<{ name: string; value: unknown }> = [];
                        if (!tooltipMeasures.length) return extra;
                        for (const name of tooltipMeasures) {
                            const rawAny = model.cfFieldRaw.get(`${r.code}||${levelsKey}||${name}`);
                            extra.push({ name, value: rawAny });
                        }
                        return extra;
                    };

                    td.addEventListener("mouseenter", (ev) => {
                        const items = buildMatrixTooltipItems({
                            rowLabel: r.label,
                            columnLabel: colLabel,
                            valueText: text,
                            extra: getExtra()
                        });
                        this.showTooltip(ev as MouseEvent, items, rowSelectionId);
                    });
                    td.addEventListener("mousemove", (ev) => {
                        const items = buildMatrixTooltipItems({
                            rowLabel: r.label,
                            columnLabel: colLabel,
                            valueText: text,
                            extra: getExtra()
                        });
                        this.showTooltip(ev as MouseEvent, items, rowSelectionId);
                    });
                    td.addEventListener("mouseleave", () => {
                        this.hideTooltip(true);
                    });

                    if (rowSelectionId && !r.isTotal) {
                        td.addEventListener("contextmenu", (ev) => {
                            ev.preventDefault();
                            try {
                                (this.selectionManager as any).showContextMenu(rowSelectionId, { x: ev.clientX, y: ev.clientY });
                            } catch {
                                // ignore context menu errors
                            }
                        });
                    }
                }

                if (cfEnabled) {
                    const colText = [...(col.columnLevels || []), col.leafLabel].filter(p => !!p).join(" ");
                    const bgRules = this.evalCfRulesForValueCell(cfConfig, {
                        rowCode: r.code,
                        rowLabel: r.label,
                        colKey: col.key,
                        colText,
                        value: v,
                        formatProp: "background"
                    });
                    if (bgRules) this.applyCfStyle(td, bgRules);

                    const fontRules = this.evalCfRulesForValueCell(cfConfig, {
                        rowCode: r.code,
                        rowLabel: r.label,
                        colKey: col.key,
                        colText,
                        value: v,
                        formatProp: "fontColor"
                    });
                    if (fontRules) this.applyCfStyle(td, fontRules);

                    const leafLabel = (col.leafLabel || "").trim();

                    const getRawForBasedOn = (basedOnField: string): any => {
                        const want = (basedOnField || "").trim();
                        if (!want) return model.valuesRaw.get(r.code)?.get(col.key);

                        // 1) Row-scoped grouping fields (Row/Group/FormatByField)
                        const rowMap = model.rowFieldRaw.get(r.code);
                        if (rowMap && rowMap.has(want)) return rowMap.get(want);

                        // 2) Column-scoped grouping fields (Column levels)
                        const colMap = model.colFieldRaw.get(col.key);
                        if (colMap && colMap.has(want)) return colMap.get(want);

                        // 3) Other displayed measures (leaf label match)
                        const srcKey = resolveColKeyForBasedOn(col, want);
                        if (srcKey) return model.valuesRaw.get(r.code)?.get(srcKey);

                        // 4) Non-displayed measures (Format by)
                        // For Custom Table tuple-only columns (no measure), include leafLabel to uniquely identify the column cell.
                        const levelsKeyBase = (col.columnLevels || []).join("||");
                        const levelsKey = (col as any).measure == null
                            ? [...(col.columnLevels || []), (col.leafLabel || "")].filter(p => !!p).join("||")
                            : levelsKeyBase;
                        return model.cfFieldRaw.get(`${r.code}||${levelsKey}||${want}`);
                    };

                    const applyGradient = (surface: CFSurfaceSettings, prop: CFFormatProp) => {
                        if (surface.formatStyle !== "gradient") return;
                        const excluded = (surface as any)?.excludedRowCodes as string[] | undefined;
                        if (Array.isArray(excluded) && excluded.includes(r.code)) return;
                        const s = surface as CFGradientSettings;
                        const key = ((s.basedOnField || "").trim()) || leafLabel;
                        if (!key) return;

                        let vv: number | null = null;
                        if (!((s.basedOnField || "").trim()) || key === leafLabel) {
                            vv = v == null ? (s.emptyValues === "zero" ? 0 : null) : v;
                        } else {
                            const rawAny = getRawForBasedOn(key);
                            const num = toNumber(rawAny);
                            vv = Number.isFinite(num) ? num : (s.emptyValues === "zero" ? 0 : null);
                        }
                        if (vv == null || !Number.isFinite(vv)) return;

                        const stats = gradientStatsByKey.get(key);

                        const parseBound = (b: CFGradientBound, fallback: number): number => {
                            if (b.type === "number") {
                                const n = Number((b.value || "").trim());
                                return Number.isFinite(n) ? n : fallback;
                            }
                            return fallback;
                        };

                        const autoMin = stats ? stats.min : vv;
                        const autoMax = stats ? stats.max : vv;
                        const minVal = parseBound(s.min, autoMin);
                        const maxVal = parseBound(s.max, autoMax);
                        const denom = (maxVal - minVal);
                        if (!Number.isFinite(denom) || denom === 0) {
                            const c = s.max.color;
                            if (prop === "background") td.style.background = c;
                            else td.style.color = c;
                            return;
                        }

                        const t = this.clamp01((vv - minVal) / denom);
                        let colorHex = "";
                        if (s.useMid) {
                            const midT = 0.5;
                            if (t <= midT) {
                                const tt = t / midT;
                                colorHex = this.lerpColorHex(s.min.color, s.midColor, tt);
                            } else {
                                const tt = (t - midT) / (1 - midT);
                                colorHex = this.lerpColorHex(s.midColor, s.max.color, tt);
                            }
                        } else {
                            colorHex = this.lerpColorHex(s.min.color, s.max.color, t);
                        }

                        if (prop === "background") td.style.background = colorHex;
                        else td.style.color = colorHex;
                    };

                    const applyFieldValue = (surface: CFSurfaceSettings, prop: CFFormatProp) => {
                        if (surface.formatStyle !== "fieldValue") return;
                        const excluded = (surface as any)?.excludedRowCodes as string[] | undefined;
                        if (Array.isArray(excluded) && excluded.includes(r.code)) return;
                        const s = surface as CFFieldValueSettings;
                        const rawColor = getRawForBasedOn(s.basedOnField);
                        const parsed = this.parseSafeCssColor(rawColor);
                        if (!parsed) return;
                        if (prop === "background") td.style.background = parsed;
                        else td.style.color = parsed;
                    };

                    applyGradient(cfCellBgSurface as any, "background");
                    applyGradient(cfCellFontSurface as any, "fontColor");
                    applyFieldValue(cfCellBgSurface as any, "background");
                    applyFieldValue(cfCellFontSurface as any, "fontColor");

                    // Icons (Rules + Field value)
                    let iconEl: HTMLElement | null = null;

                    // Rules-based icon
                    const iconRules = this.evalCfRulesForValueCell(cfConfig, {
                        rowCode: r.code,
                        rowLabel: r.label,
                        colKey: col.key,
                        colText,
                        value: v,
                        formatProp: "icon"
                    });
                    if (iconRules?.iconName) {
                        iconEl = this.createIconElementFromSpec(iconRules.iconName);
                    }

                    // Field-value icon surface
                    if (!iconEl && cfCellIconSurface.formatStyle === "fieldValue") {
                        const excluded = (cfCellIconSurface as any)?.excludedRowCodes as string[] | undefined;
                        if (Array.isArray(excluded) && excluded.includes(r.code)) {
                            // Excluded row: skip surface-based icons.
                        } else {
                        const s = cfCellIconSurface as CFFieldValueSettings;
                        const rawIcon = getRawForBasedOn(s.basedOnField);
                        iconEl = this.createIconElementFromSpec(rawIcon);
                        }
                    }

                    if (iconEl) {
                        const placement = ((cfCellIconSurface as any).iconPlacement as CFIconPlacement | undefined) || "left";
                        const wrap = document.createElement("span");
                        wrap.className = "fm-cell-flex";
                        td.textContent = "";
                        if (placement === "only") {
                            wrap.appendChild(iconEl);
                            td.appendChild(wrap);
                        } else {
                            const txt = document.createElement("span");
                            txt.className = "fm-cell-text";
                            txt.textContent = text;
                            if (placement === "left") {
                                wrap.appendChild(iconEl);
                                wrap.appendChild(txt);
                            } else {
                                wrap.appendChild(txt);
                                wrap.appendChild(iconEl);
                            }
                            td.appendChild(wrap);
                        }
                    }
                }
                tr.appendChild(td);
            }

            if (model.showColumnTotal) {
                const td = document.createElement("td");
                td.className = "fm-td fm-td-num";
                if (columnTotalGrid && colTotalBorder) {
                    td.style.borderLeft = colTotalBorder;
                    td.style.borderRight = colTotalBorder;
                }
                if (r.isTotal) {
                    const isGrand = r.code === "__grand_total";
                    const isSub = r.code.endsWith("||__subtotal");
                    if (isGrand && grandTotalGrid && grandRowBorder) {
                        td.style.borderTop = grandRowBorder;
                        td.style.borderBottom = grandRowBorder;
                    }
                    if (isSub && subtotalGrid && subRowBorder) {
                        td.style.borderTop = subRowBorder;
                        td.style.borderBottom = subRowBorder;
                    }
                }
                // Column Total column is its own style surface.
                const totalSurfaceStyle = r.isTotal ? rowTotalStyle : columnTotalStyle;
                if (totalSurfaceStyle.fontFamily) td.style.fontFamily = totalSurfaceStyle.fontFamily;
                if (totalSurfaceStyle.fontSize) td.style.fontSize = `${totalSurfaceStyle.fontSize}px`;
                td.style.textAlign = totalSurfaceStyle.valueAlign as any;
                if (totalSurfaceStyle.bold) td.style.fontWeight = "700";
                if (totalSurfaceStyle.color) td.style.color = totalSurfaceStyle.color;
                if (totalSurfaceStyle.background) td.style.background = totalSurfaceStyle.background;
                if (isManualWidth) {
                    td.style.width = `${columnWidthPx}px`;
                    td.style.minWidth = `${columnWidthPx}px`;
                    td.style.maxWidth = `${columnWidthPx}px`;
                }
                if (r.type === "blank") {
                    td.textContent = "";
                } else {
                    let sum = 0;
                    let hasAny = false;
                    for (const col of model.columns) {
                        const rowVals = model.values.get(r.code);
                        const hasMappedValue = !!rowVals && rowVals.has(col.key);
                        const raw = hasMappedValue ? (rowVals!.get(col.key) ?? null) : null;
                        const canBlankAsZero = blankAsZero && (autoAggregateParents || !isParent) && hasMappedValue;
                        const v = raw == null ? (canBlankAsZero ? 0 : null) : raw;
                        if (v == null || !Number.isFinite(v)) continue;
                        hasAny = true;
                        sum += v;
                    }
                    const vTotal = hasAny ? sum : (blankAsZero ? 0 : null);
                    const text = units === "auto" ? this.formatAutoNumber(vTotal) : this.formatNumber(vTotal, decimals, 1);
                    td.textContent = text;

                    if (cfEnabled) {
                        const bgRules = this.evalCfRulesForValueCell(cfConfig, {
                            rowCode: r.code,
                            rowLabel: r.label,
                            colKey: "__column_total",
                            colText: columnTotalLabel,
                            value: vTotal,
                            formatProp: "background"
                        });
                        if (bgRules) this.applyCfStyle(td, bgRules);

                        const fontRules = this.evalCfRulesForValueCell(cfConfig, {
                            rowCode: r.code,
                            rowLabel: r.label,
                            colKey: "__column_total",
                            colText: columnTotalLabel,
                            value: vTotal,
                            formatProp: "fontColor"
                        });
                        if (fontRules) this.applyCfStyle(td, fontRules);
                    }
                }
                tr.appendChild(td);
            }

            tbody.appendChild(tr);
        }

        table.appendChild(tbody);
        this.root.appendChild(table);

        // Render Conditional Formatting editor (popup) if enabled/open.
        this.renderCfEditor(model);

        // Render Custom Table editor (popup) if enabled/open.
        this.renderCustomTableEditor();

        // Freeze behavior (Power BI WebView friendly): use scroll-synced transforms.
        if (freezeRowHeaders || freezeHeader) {
            const thead = table.querySelector("thead") as HTMLElement | null;

            // Use scroll-synced transforms for freeze. Sticky on table cells is unreliable in Power BI WebView.

            const headerCells = thead
                ? (Array.from(thead.querySelectorAll("th")) as HTMLTableCellElement[])
                : ([] as HTMLTableCellElement[]);
            const headerRowHeaderCells = thead
                ? (Array.from(thead.querySelectorAll("th.fm-th-row")) as HTMLTableCellElement[])
                : ([] as HTMLTableCellElement[]);
            const bodyRowHeaderCells = Array.from(table.querySelectorAll("tbody td.fm-td-row")) as HTMLTableCellElement[];

            // Cover layer: hides row-header cells from painting over the header region during vertical scroll.
            // We keep row headers above month headers for horizontal scroll, but the header itself must always win vertically.
            const leftTopCover = (freezeRowHeaders && thead) ? document.createElement("div") : null;
            if (leftTopCover) {
                leftTopCover.className = "fm-freeze-left-top-cover";
                // Between header (20) and row-header title cells (40), but above body row-header cells (30).
                leftTopCover.style.zIndex = "35";
                this.root.insertBefore(leftTopCover, table);
            }

            const leftMask = null as HTMLDivElement | null;

            const isTransparent = (bg: string): boolean => {
                const s = (bg || "").trim().toLowerCase();
                return !s || s === "transparent" || s === "rgba(0, 0, 0, 0)";
            };

            const parseRgba = (bg: string): { r: number; g: number; b: number; a: number } | null => {
                const s = (bg || "").trim().toLowerCase();
                const m1 = s.match(/^rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([0-9]*\.?[0-9]+)\s*\)$/);
                if (m1) {
                    return { r: Number(m1[1]), g: Number(m1[2]), b: Number(m1[3]), a: Math.max(0, Math.min(1, Number(m1[4]))) };
                }
                const m2 = s.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/);
                if (m2) {
                    return { r: Number(m2[1]), g: Number(m2[2]), b: Number(m2[3]), a: 1 };
                }
                return null;
            };

            const blendToOpaque = (fg: string, bg: string): string => {
                const f = parseRgba(fg);
                const b = parseRgba(bg);
                if (!f) return fg;
                const under = b ?? { r: 255, g: 255, b: 255, a: 1 };
                const a = Math.max(0, Math.min(1, f.a));
                const r = Math.round((a * f.r) + ((1 - a) * under.r));
                const g = Math.round((a * f.g) + ((1 - a) * under.g));
                const bb = Math.round((a * f.b) + ((1 - a) * under.b));
                return `rgb(${r}, ${g}, ${bb})`;
            };

            const primeFreezeBackgrounds = () => {
                // We need row headers to be opaque so content underneath is hidden.
                // Use the cell's computed bg if available, else inherit from the row,
                // else fall back to the visual background (or a neutral light surface).
                const rootBg = window.getComputedStyle(this.root).backgroundColor || "";
                const rootFallback = isTransparent(rootBg) ? "rgb(255, 255, 255)" : blendToOpaque(rootBg, "rgb(255, 255, 255)");

                // Ensure the overall table surface is opaque in freeze mode so gaps (border-spacing/card presets)
                // don't reveal scrolled content underneath.
                table.style.backgroundColor = rootFallback;

                // Size and color the top-left cover.
                if (leftTopCover && thead) {
                    let coverWidth = 0;
                    if (isManualWidth) {
                        coverWidth = model.rowHeaderColumns.length * rowHeaderColWidthPx;
                    } else {
                        const base = headerRowHeaderCells.length ? headerRowHeaderCells : bodyRowHeaderCells.slice(0, model.rowHeaderColumns.length);
                        for (let i = 0; i < Math.min(model.rowHeaderColumns.length, base.length); i++) {
                            coverWidth += base[i].getBoundingClientRect().width;
                        }
                        if (!Number.isFinite(coverWidth) || coverWidth <= 0) {
                            coverWidth = model.rowHeaderColumns.length * rowHeaderColWidthPx;
                        }
                    }
                    const coverHeight = Math.max(0, Math.round((thead as HTMLElement).getBoundingClientRect().height || 0));
                    leftTopCover.style.width = `${Math.max(0, Math.round(coverWidth))}px`;
                    leftTopCover.style.height = `${coverHeight}px`;

                    const headerBgFromCells = headerCells.length ? (window.getComputedStyle(headerCells[0]).backgroundColor || "") : "";
                    const coverBg = !isTransparent(headerBgFromCells)
                        ? blendToOpaque(headerBgFromCells, rootFallback)
                        : rootFallback;
                    leftTopCover.style.backgroundColor = coverBg;
                }

                if (freezeRowHeaders) {
                    for (const cell of bodyRowHeaderCells) {
                        const cellBg = window.getComputedStyle(cell).backgroundColor || "";
                        const row = cell.parentElement as HTMLElement | null;
                        const rowBg = row ? (window.getComputedStyle(row).backgroundColor || "") : "";
                        const chosen = !isTransparent(cellBg)
                            ? blendToOpaque(cellBg, rootFallback)
                            : (!isTransparent(rowBg) ? blendToOpaque(rowBg, rootFallback) : rootFallback);
                        cell.style.backgroundColor = chosen;
                        // Subtle separation from the scrolled value area.
                        cell.style.boxShadow = "1px 0 0 rgba(17, 24, 39, 0.08)";
                    }

                    // Also make the header's row-header cells (e.g., "Segment") opaque so
                    // horizontally-scrolled column headers can't show underneath when only row-freeze is ON.
                    if (thead) {
                        const headerBgFromCells = headerCells.length ? (window.getComputedStyle(headerCells[0]).backgroundColor || "") : "";
                        const headerLeftFallback = !isTransparent(headerBgFromCells)
                            ? blendToOpaque(headerBgFromCells, rootFallback)
                            : rootFallback;
                        for (const th of headerRowHeaderCells) {
                            const thBg = window.getComputedStyle(th).backgroundColor || "";
                            th.style.backgroundColor = isTransparent(thBg) ? headerLeftFallback : blendToOpaque(thBg, rootFallback);
                            th.style.boxShadow = "1px 0 0 rgba(17, 24, 39, 0.08)";
                        }
                    }
                }

                if (freezeHeader && thead) {
                    // Use computed header background as the source of truth. Avoid CSS var(...) strings here,
                    // because they can't be parsed/blended reliably and can leave the header transparent.
                    const computedHeaderBg = headerCells.length ? (window.getComputedStyle(headerCells[0]).backgroundColor || "") : "";
                    const headerBgFallback = !isTransparent(computedHeaderBg)
                        ? blendToOpaque(computedHeaderBg, rootFallback)
                        : rootFallback;

                    // Also paint the thead background to cover any gaps between cells.
                    (thead as HTMLElement).style.backgroundColor = headerBgFallback;
                    for (const th of headerCells) {
                        const thBg = window.getComputedStyle(th).backgroundColor || "";
                        if (isTransparent(thBg)) {
                            (th as HTMLElement).style.backgroundColor = headerBgFallback;
                        } else {
                            (th as HTMLElement).style.backgroundColor = blendToOpaque(thBg, rootFallback);
                        }
                    }
                    for (const th of headerRowHeaderCells) {
                        const thBg = window.getComputedStyle(th).backgroundColor || "";
                        if (isTransparent(thBg)) {
                            (th as HTMLElement).style.backgroundColor = headerBgFallback;
                        } else {
                            (th as HTMLElement).style.backgroundColor = blendToOpaque(thBg, rootFallback);
                        }
                        (th as HTMLElement).style.boxShadow = "1px 0 0 rgba(17, 24, 39, 0.08)";
                    }
                }
            };

            const applyFreeze = () => {
                const sl = this.root.scrollLeft;
                const st = this.root.scrollTop;

                if (leftTopCover) {
                    // Only needed when vertically scrolled; otherwise it can visually mask the first row.
                    leftTopCover.style.display = st > 0 ? "block" : "none";
                    // Keep the cover pinned to the visible top-left of the scroll container.
                    leftTopCover.style.left = `${sl}px`;
                    leftTopCover.style.top = `${st}px`;
                }

                if (freezeHeader && thead) {
                    thead.style.position = "relative";
                    // Header must sit above values, but below frozen row headers.
                    thead.style.zIndex = "20";
                    // Use top instead of transforms to avoid font-weight looking bolder on Windows.
                    thead.style.top = `${st}px`;
                    thead.style.transform = "";

                    for (const th of headerCells) {
                        th.style.position = "relative";
                        th.style.zIndex = "20";
                    }
                } else if (thead) {
                    thead.style.top = "";
                    thead.style.transform = "";
                }

                if (freezeRowHeaders) {
                    // When header freeze is ON, keep row headers above the header area.
                    const zBodyRowHeader = freezeHeader ? "60" : "30";
                    const zTopLeftCover = freezeHeader ? "70" : "35";
                    const zHeaderRowHeader = freezeHeader ? "80" : "40";

                    if (leftTopCover) leftTopCover.style.zIndex = zTopLeftCover;

                    for (const td of bodyRowHeaderCells) {
                        td.style.position = "relative";
                        // Above month headers (20), but below the top-left cover and corner/header-left cells.
                        td.style.zIndex = zBodyRowHeader;
                        // Use left instead of transforms to preserve text rendering.
                        td.style.left = `${sl}px`;
                        td.style.transform = "";
                    }

                    for (const th of headerRowHeaderCells) {
                        th.style.position = "relative";
                        th.style.zIndex = zHeaderRowHeader;
                        // Only freeze horizontally. Vertical behavior follows the header (if frozen) or scroll (if not).
                        th.style.left = `${sl}px`;
                        th.style.transform = "";
                    }
                } else {
                    for (const td of bodyRowHeaderCells) {
                        td.style.left = "";
                        td.style.transform = "";
                    }
                    for (const th of headerRowHeaderCells) {
                        th.style.left = "";
                        th.style.transform = "";
                    }
                }
            };

            // Replace any prior handler on each render.
            (this.root as any).onscroll = applyFreeze;
            requestAnimationFrame(() => {
                primeFreezeBackgrounds();
                applyFreeze();
            });
        }
    }

    private isVisible(r: RowNode, all: RowNode[]): boolean {
        let p = r.parent;
        while (p) {
            if (this.collapsed.has(p)) return false;
            const parentNode = all.find(x => x.code === p);
            p = parentNode ? parentNode.parent : null;
        }
        return true;
    }

    private getRowStyle(r: RowNode, isParent: boolean): Required<RowStyleOverride> {
        const parentBold = this.formattingSettings.rows.parentBold.value;
        const parentColor = this.formattingSettings.rows.parentFontColor.value.value;
        const parentBg = this.formattingSettings.rows.parentBackground.value.value;
        const childColor = this.formattingSettings.rows.childFontColor.value.value;
        const childBg = this.formattingSettings.rows.childBackground.value.value;

        const base: Required<RowStyleOverride> = {
            bold: isParent ? parentBold : false,
            color: isParent ? (parentColor || "") : (childColor || ""),
            background: isParent ? (parentBg || "") : (childBg || ""),
            italic: false,
            underline: false
        };

        if (r.style) {
            return {
                bold: r.style.bold ?? base.bold,
                color: r.style.color ?? base.color,
                background: r.style.background ?? base.background,
                italic: r.style.italic ?? base.italic,
                underline: r.style.underline ?? base.underline
            };
        }
        return base;
    }

    private formatColumnLeafLabel(col: Pick<BucketedColumnKey, "bucket" | "key"> & Partial<Pick<BucketedColumnKey, "measure">>): string {
        const suffix = col.measure || col.key;
        return suffix;
    }

    private computeUnitScale(units: DisplayUnit, model: Model, visibleRows: RowNode[], colKey: string): number {
        // Custom Auto formatting is applied per-cell; no column-wide scaling.
        if (units === "auto") return 1;
        if (units === "none") return 1;
        if (units === "thousands") return 1000;
        if (units === "millions") return 1000000;
        if (units === "billions") return 1000000000;

        let maxAbs = 0;
        for (const r of visibleRows) {
            const v = model.values.get(r.code)?.get(colKey);
            if (v == null || !Number.isFinite(v)) continue;
            maxAbs = Math.max(maxAbs, Math.abs(v));
        }
        if (maxAbs >= 1_000_000_000) return 1_000_000_000;
        if (maxAbs >= 1_000_000) return 1_000_000;
        if (maxAbs >= 1_000) return 1_000;
        return 1;
    }

    // Custom Auto number formatting logic (based on user-provided DAX format switch).
    // Ranges:
    // <= -1B => x.xxB
    // [-1B,-1M) => x.xxM
    // [-1M,-1K) => x.xxK
    // (-1K,0) => -x.xx
    // 0 => 0
    // (0,1) => 0.xx
    // [1,1K) => x.xx
    // [1K,1M) => x.xxK
    // [1M,1B) => x.xxM
    // >= 1B => x.xxbn
    private formatAutoNumber(value: number | null): string {
        if (value == null || !Number.isFinite(value)) return "";
        if (value === 0) return "0";

        const abs = Math.abs(value);
        const sign = value < 0 ? "-" : "";

        const fmt2 = (n: number): string => {
            return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        };

        if (value <= -1_000_000_000) return `${sign}${fmt2(abs / 1_000_000_000)}B`;
        if (value <= -1_000_000 && value > -1_000_000_000) return `${sign}${fmt2(abs / 1_000_000)}M`;
        if (value <= -1_000 && value > -1_000_000) return `${sign}${fmt2(abs / 1_000)}K`;
        if (value < 0 && value > -1_000) return `${sign}${fmt2(abs)}`;

        if (value > 0 && value < 1) return fmt2(value);
        if (value >= 1 && value < 1_000) return fmt2(value);
        if (value >= 1_000 && value < 1_000_000) return `${fmt2(value / 1_000)}K`;
        if (value >= 1_000_000 && value < 1_000_000_000) return `${fmt2(value / 1_000_000)}M`;
        if (value >= 1_000_000_000) return `${fmt2(value / 1_000_000_000)}bn`;

        return fmt2(value);
    }

    private formatNumber(value: number | null, decimals: number, scale: number): string {
        if (value == null || !Number.isFinite(value)) return "";
        const scaled = value / (scale || 1);
        return scaled.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    }

    private buildModel(dataView: DataView): Model {
        const ctCard: any = (this.formattingSettings as any).customTable;
        const customTableEnabled = (ctCard?.enabled?.value as boolean | undefined) ?? false;
        if (customTableEnabled) {
            const cfg = this.safeParseCustomTableConfig(this.customTableJson);
            return this.buildModelFromCustomTable(dataView, cfg);
        }

        const categorical = dataView.categorical;
        const categories = categorical?.categories ?? [];
        const values = categorical?.values ?? [];

        const rowCats = categories.filter(c => !!c.source.roles?.category);
        const groupCat = categories.find(c => !!c.source.roles?.group);
        const categoryClassCat = categories.find(c => !!c.source.roles?.categoryClass);
        const commentsCat = categories.find(c => !!c.source.roles?.comments);
        const columnCats = categories.filter(c => !!c.source.roles?.column);
        const formatByFieldCats = categories.filter(c => !!c.source.roles?.formatByField);

        const rowCatValues = rowCats.map(c => (c.values ?? []) as any[]);
        const groupValues = (groupCat?.values ?? []) as any[];
        const categoryClassValues = (categoryClassCat?.values ?? []) as any[];
        const commentsValues = (commentsCat?.values ?? []) as any[];
        const columnCatValues = columnCats.map(c => (c.values ?? []) as any[]);
        const formatByFieldValues = formatByFieldCats.map(c => (c.values ?? []) as any[]);

        const rowCount = rowCatValues.length > 0 ? rowCatValues[0].length : 0;

        const norm = (v: any): string => {
            const s = v == null ? "" : String(v);
            const t = s.trim();
            return t ? t : "(Blank)";
        };

        const getRowLevels = (i: number): string[] => {
            if (rowCatValues.length === 0) return [];
            // IMPORTANT: keep blanks as explicit "(Blank)" so parent totals stay correct.
            return rowCatValues.map(arr => norm(arr[i]));
        };
        const getGroup = (i: number) => {
            const v = groupValues[i];
            const s = v == null ? "" : String(v);
            const t = s.trim();
            return t ? t : null;
        };
        const getClass = (i: number) => {
            const v = categoryClassValues[i];
            const s = v == null ? "" : String(v);
            const t = s.trim();
            return t ? t : null;
        };
        const getComment = (i: number) => {
            const v = commentsValues[i];
            const s = v == null ? "" : String(v);
            const t = s.trim();
            return t ? t : null;
        };

        const getColumnLevels = (i: number): string[] => {
            if (columnCatValues.length === 0) return [];
            // Keep blank levels so tuple length remains stable.
            return columnCatValues.map(arr => norm(arr[i]));
        };

        const pickBucket = (v: powerbi.DataViewValueColumn): ColumnBucket | null => {
            const roles: any = v?.source?.roles || {};
            if (roles.values || roles.customTableValue || roles.formatBy) return "values";
            return null;
        };

        const getMeasureDisplayName = (v: powerbi.DataViewValueColumn) => (v.source?.displayName || v.source?.queryName || "Value");

        const columns: BucketedColumnKey[] = [];
        const orderedBuckets: ColumnBucket[] = ["values"];
        const columnsByKey = new Map<string, BucketedColumnKey>();

        const makeLevels = (colLevels: string[], leaf: string): string[] => {
            if (colLevels.length === 0) {
                return [leaf];
            }
            if (colLevels.length === 1) {
                return [`${colLevels[0]} ${leaf}`];
            }
            // Keep top-level grouping (e.g., Year) but also include it with lower levels
            // so users can see Year alongside Country/Measure text.
            const levels = colLevels.slice(0, colLevels.length - 1);
            levels.push(`${colLevels.join(" ")} ${leaf}`);
            return levels;
        };

        const ensureColumn = (bucket: ColumnBucket, measureName: string, colLevels: string[]): BucketedColumnKey => {
            const leaf = `${bucket}||${measureName}`;
            const key = `${colLevels.join("||")}||${leaf}`;
            const existing = columnsByKey.get(key);
            if (existing) return existing;
            const leafLabel = this.formatColumnLeafLabel({ bucket, measure: measureName, key: leaf });
            const col: BucketedColumnKey = {
                bucket,
                measure: measureName,
                key,
                levels: makeLevels(colLevels, leafLabel),
                columnLevels: colLevels.slice(),
                leafLabel
            };
            columnsByKey.set(key, col);
            columns.push(col);
            return col;
        };

        const valuesMap = new Map<string, Map<string, number | null>>();
        const valuesRawMap = new Map<string, Map<string, any>>();
        const cfFieldRaw = new Map<string, any>();
        const rowFieldRaw = new Map<string, Map<string, any>>();
        const colFieldRaw = new Map<string, Map<string, any>>();
        const rowSelectionIdByCode = new Map<string, powerbi.extensibility.ISelectionId>();

        const tooltipMeasureSet = new Set<string>();
        for (const v of values) {
            const roles: any = v?.source?.roles || {};
            if (roles.tooltips) {
                tooltipMeasureSet.add(getMeasureDisplayName(v));
            }
        }
        const tooltipMeasureNames = Array.from(tooltipMeasureSet.values());
        const dataLabelByCode = new Map<string, string>();
        const dataParentByCode = new Map<string, string | null>();
        const dataRowLevelByCode = new Map<string, number>();
        const dataGroupKeyByCode = new Map<string, string | null>();

        const setCell = (code: string, colKey: string, v: number | null) => {
            if (!valuesMap.has(code)) valuesMap.set(code, new Map());
            valuesMap.get(code)!.set(colKey, v);
        };

        const addCell = (code: string, colKey: string, rawAny: any) => {
            const n = toNumber(rawAny);
            if (!Number.isFinite(n)) {
                // Only set a null if nothing exists yet; don't clobber prior numeric sum.
                if (!valuesMap.has(code)) valuesMap.set(code, new Map());
                if (!valuesMap.get(code)!.has(colKey)) valuesMap.get(code)!.set(colKey, null);
                if (!valuesRawMap.has(code)) valuesRawMap.set(code, new Map());
                if (!valuesRawMap.get(code)!.has(colKey)) valuesRawMap.get(code)!.set(colKey, rawAny);
                return;
            }

            if (!valuesMap.has(code)) valuesMap.set(code, new Map());
            const prev = valuesMap.get(code)!.get(colKey);
            const next = (prev == null || !Number.isFinite(prev as any)) ? n : (prev + n);
            valuesMap.get(code)!.set(colKey, next);

            if (!valuesRawMap.has(code)) valuesRawMap.set(code, new Map());
            valuesRawMap.get(code)!.set(colKey, next);
        };

        const setCellRaw = (code: string, colKey: string, v: any) => {
            if (!valuesRawMap.has(code)) valuesRawMap.set(code, new Map());
            valuesRawMap.get(code)!.set(colKey, v);
        };

        const setCfFieldRaw = (rowCode: string, colLevels: string[], fieldLabel: string, v: any) => {
            const ll = (fieldLabel || "").trim();
            if (!ll) return;
            const key = `${rowCode}||${colLevels.join("||")}||${ll}`;
            cfFieldRaw.set(key, v);
        };

        const addCfMeasureRaw = (rowCode: string, colLevels: string[], fieldLabel: string, rawAny: any) => {
            const ll = (fieldLabel || "").trim();
            if (!ll) return;
            const key = `${rowCode}||${colLevels.join("||")}||${ll}`;
            const n = toNumber(rawAny);
            if (!Number.isFinite(n)) {
                if (!cfFieldRaw.has(key)) cfFieldRaw.set(key, rawAny);
                return;
            }
            const prevAny = cfFieldRaw.get(key);
            const prev = toNumber(prevAny);
            const next = Number.isFinite(prev) ? (prev + n) : n;
            cfFieldRaw.set(key, next);
        };

        const setRowFieldRaw = (rowCode: string, fieldLabel: string, v: any) => {
            const ll = (fieldLabel || "").trim();
            if (!ll) return;
            if (!rowFieldRaw.has(rowCode)) rowFieldRaw.set(rowCode, new Map());
            rowFieldRaw.get(rowCode)!.set(ll, v);
        };

        const setColFieldRaw = (colKey: string, fieldLabel: string, v: any) => {
            const ll = (fieldLabel || "").trim();
            if (!ll) return;
            if (!colFieldRaw.has(colKey)) colFieldRaw.set(colKey, new Map());
            colFieldRaw.get(colKey)!.set(ll, v);
        };

        for (let i = 0; i < rowCount; i++) {
            const rowLevels = getRowLevels(i);
            if (rowLevels.length === 0) continue;
            const grp = getGroup(i);
            const leafCode = rowLevels.join("||");
            const colLevels = getColumnLevels(i);

            // Create selection ids for each row level (prefix) so clicking parent/leaf rows filters correctly.
            try {
                for (let lvl = 0; lvl < rowCats.length; lvl++) {
                    const catCol = rowCats[lvl];
                    const code = rowLevels.slice(0, lvl + 1).join("||");
                    if (!catCol || !code) continue;
                    if (rowSelectionIdByCode.has(code)) continue;
                    const sid = this.host.createSelectionIdBuilder().withCategory(catCol, i).createSelectionId();
                    rowSelectionIdByCode.set(code, sid);
                }
            } catch {
                // ignore selection id issues
            }

            // Store raw bound grouping/category values for this row.
            for (let ri = 0; ri < rowCats.length; ri++) {
                const fieldName = rowCats[ri]?.source?.displayName || rowCats[ri]?.source?.queryName || "";
                if (!fieldName) continue;
                setRowFieldRaw(leafCode, fieldName, (rowCatValues[ri] || [])[i]);
            }
            for (let fi = 0; fi < formatByFieldCats.length; fi++) {
                const fieldName = formatByFieldCats[fi]?.source?.displayName || formatByFieldCats[fi]?.source?.queryName || "";
                if (!fieldName) continue;
                setRowFieldRaw(leafCode, fieldName, (formatByFieldValues[fi] || [])[i]);
            }
            if (groupCat) {
                const fieldName = groupCat?.source?.displayName || groupCat?.source?.queryName || "Group";
                setRowFieldRaw(leafCode, fieldName, groupValues[i]);
            }

            // Store raw bound grouping/category values for this column tuple (by rendered column keys).
            for (let ci = 0; ci < columnCats.length; ci++) {
                const fieldName = columnCats[ci]?.source?.displayName || columnCats[ci]?.source?.queryName || "";
                if (!fieldName) continue;
                setColFieldRaw(colLevels.join("||"), fieldName, (columnCatValues[ci] || [])[i]);
            }

            // Ensure row hierarchy nodes exist for each prefix
            let parent: string | null = grp;
            if (grp) {
                if (!dataLabelByCode.has(grp)) {
                    dataLabelByCode.set(grp, grp);
                    dataParentByCode.set(grp, null);
                    dataRowLevelByCode.set(grp, -1);
                    dataGroupKeyByCode.set(grp, grp);
                }
            }

            for (let lvl = 0; lvl < rowLevels.length; lvl++) {
                const code = rowLevels.slice(0, lvl + 1).join("||");
                const label = rowLevels[lvl];
                if (!dataLabelByCode.has(code)) {
                    dataLabelByCode.set(code, label);
                }
                if (!dataParentByCode.has(code)) {
                    dataParentByCode.set(code, parent);
                }
                if (!dataRowLevelByCode.has(code)) {
                    dataRowLevelByCode.set(code, lvl);
                }
                if (!dataGroupKeyByCode.has(code)) {
                    dataGroupKeyByCode.set(code, grp);
                }
                parent = code;
            }

            // Store measures into bucketed columns + column hierarchy
            for (const v of values) {
                const bucket = pickBucket(v);
                if (!bucket) continue;
                const m = getMeasureDisplayName(v);
                const col = ensureColumn(bucket, m, colLevels);
                // Attach column group values to the rendered column key so Field value can reference them.
                for (let ci = 0; ci < columnCats.length; ci++) {
                    const fieldName = columnCats[ci]?.source?.displayName || columnCats[ci]?.source?.queryName || "";
                    if (!fieldName) continue;
                    setColFieldRaw(col.key, fieldName, (columnCatValues[ci] || [])[i]);
                }
                const raw = v.values[i] as any;
                // IMPORTANT: If the dataview returns multiple rows for the same (rowLevels, colLevels),
                // accumulate numeric values to match Matrix behavior (Sum).
                addCell(leafCode, col.key, raw);
            }

            // Capture raw values for any bound measures (including formatBy) for Field value.
            for (const v of values) {
                const src: any = v?.source;
                const name = (src?.displayName || src?.queryName || "").toString().trim();
                if (!name) continue;
                const roles: any = src?.roles || {};
                // Only store measures.
                if (!src?.isMeasure && !(roles.values || roles.customTableValue || roles.formatBy || roles.tooltips)) continue;
                addCfMeasureRaw(leafCode, colLevels, name, v.values[i]);
            }

            // Optional (currently unused) metadata buckets
            void getClass(i);
            void getComment(i);
        }

        // Sort columns by level tuple (sequence is based on Column bucket field order)
        columns.sort((a, b) => {
            const len = Math.max(a.levels.length, b.levels.length);
            for (let i = 0; i < len; i++) {
                const av = a.levels[i] ?? "";
                const bv = b.levels[i] ?? "";
                const c = av.localeCompare(bv);
                if (c !== 0) return c;
            }
            return 0;
        });

        const layoutRows = this.parseLayoutRows(this.formattingSettings.layout.layoutJson.value);
        const formulas = this.parseFormulas(this.formattingSettings.layout.formulasJson.value);

        const rowByCode = new Map<string, RowNode>();
        let orderIndex = 0;

        const addRow = (r: LayoutRow, fallbackLabel?: string, fallbackParent?: string | null) => {
            const code = (r.code || "").trim();
            if (!code) return;
            const label = (r.label || fallbackLabel || code).trim();
            const parent = r.parent !== undefined ? (r.parent ? r.parent.trim() : null) : (fallbackParent ?? null);
            const type: LayoutRowType = (r.type || "data") as LayoutRowType;
            const formula = (r.formula || formulas.get(code) || undefined);
            rowByCode.set(code, {
                code,
                label,
                parent,
                type,
                orderIndex: orderIndex++,
                order: r.order,
                formula,
                style: r.style,
                children: [],
                depth: 0,
                rowLevel: dataRowLevelByCode.get(code),
                groupKey: dataGroupKeyByCode.get(code) ?? null
            });
        };

        for (const r of layoutRows) {
            addRow(r, undefined, null);
        }

        for (const [code, label] of dataLabelByCode.entries()) {
            if (rowByCode.has(code)) {
                const existing = rowByCode.get(code)!;
                if (!existing.parent) {
                    existing.parent = dataParentByCode.get(code) ?? null;
                }
                if (!existing.label) {
                    existing.label = label;
                }
                continue;
            }
            addRow({ code, type: "data" }, label, dataParentByCode.get(code) ?? null);
        }

        const hasGroup = !!groupCat;
        const rowFieldCount = rowCats.length;
        const rowHeaderColumns: string[] = [];

        const hierarchyView = ((this.formattingSettings.rows as any).hierarchyView?.value as boolean | undefined) ?? false;
        if (hasGroup) {
            rowHeaderColumns.push(groupCat?.source?.displayName || "Group");
        }

        if (hierarchyView) {
            const title = rowCats.length > 0
                ? rowCats.map(rc => rc.source?.displayName || "").join(" > ")
                : "";
            rowHeaderColumns.push(title);
        } else {
            if (rowFieldCount > 0) {
                for (const rc of rowCats) {
                    rowHeaderColumns.push(rc.source?.displayName || "");
                }
            } else {
                rowHeaderColumns.push("");
            }
        }

        for (const r of rowByCode.values()) {
            if (r.parent && rowByCode.has(r.parent)) {
                rowByCode.get(r.parent)!.children.push(r.code);
            }
        }

        const roots: string[] = [];
        for (const r of rowByCode.values()) {
            if (!r.parent || !rowByCode.has(r.parent)) {
                roots.push(r.code);
            }
        }

        const sortChildren = (codes: string[]) => {
            codes.sort((a, b) => {
                const ra = rowByCode.get(a)!;
                const rb = rowByCode.get(b)!;
                const oa = ra.order ?? ra.orderIndex;
                const ob = rb.order ?? rb.orderIndex;
                if (oa !== ob) return oa - ob;
                return ra.label.localeCompare(rb.label);
            });
        };

        for (const r of rowByCode.values()) {
            sortChildren(r.children);
        }
        sortChildren(roots);

        const flatRows: RowNode[] = [];
        const visit = (code: string, depth: number) => {
            const r = rowByCode.get(code);
            if (!r) return;
            r.depth = depth;
            flatRows.push(r);
            for (const c of r.children) {
                visit(c, depth + 1);
            }
        };

        for (const rootCode of roots) {
            visit(rootCode, 0);
        }

        // Totals settings
        const totalsAny = ((this.formattingSettings as any).totals as any) || null;
        const rowsAny = (this.formattingSettings.rows as any);
        const showGrandTotal = ((totalsAny?.showGrandTotal?.value as boolean | undefined) ?? (rowsAny.showGrandTotal?.value as boolean | undefined)) ?? false;
        const showSubtotals = ((totalsAny?.showSubtotals?.value as boolean | undefined) ?? (rowsAny.showSubtotals?.value as boolean | undefined)) ?? false;
        const showColumnTotal = ((totalsAny?.showColumnTotal?.value as boolean | undefined) ?? (rowsAny.showColumnTotal?.value as boolean | undefined)) ?? false;

        const grandTotalLabel = (((totalsAny?.grandTotalLabel?.value as string | undefined) ?? (rowsAny.grandTotalLabel?.value as string | undefined) ?? "Grand Total").trim()) || "Grand Total";
        const subtotalLabelTemplate = (((totalsAny?.subtotalLabelTemplate?.value as string | undefined) ?? (rowsAny.subtotalLabelTemplate?.value as string | undefined) ?? "Total {label}").trim()) || "Total {label}";

        const formatSubtotalLabel = (label: string): string => {
            const base = String(label ?? "").trim();
            const tpl = subtotalLabelTemplate;
            if (tpl.includes("{label}")) {
                const out = tpl.replace(/\{label\}/g, base);
                return out.trim() || `Total ${base}`;
            }
            // If user didn't include placeholder, use the template as-is.
            return tpl.trim();
        };

        for (const r of rowByCode.values()) {
            if (!valuesMap.has(r.code)) valuesMap.set(r.code, new Map());
            const rowVals = valuesMap.get(r.code)!;
            for (const col of columns) {
                if (!rowVals.has(col.key)) rowVals.set(col.key, null);
            }
        }

        const blankAsZero = this.formattingSettings.rows.blankAsZero.value;
        const autoAggregateParents = this.formattingSettings.rows.autoAggregateParents.value;

        const getCell = (code: string, colKey: string): number | null => {
            const v = valuesMap.get(code)?.get(colKey);
            if (v == null) return blankAsZero ? 0 : null;
            return v;
        };

        const setComputed = (code: string, colKey: string, v: number | null) => {
            if (!valuesMap.has(code)) valuesMap.set(code, new Map());
            valuesMap.get(code)!.set(colKey, v);
        };

        const evalAst = (ast: any, ctx: { currentColKey: string }): any => {
            switch (ast.type) {
                case "num":
                    return ast.value;
                case "str":
                    return ast.value;
                case "ref":
                    return getCell(ast.value, ctx.currentColKey);
                case "ident": {
                    const upper = String(ast.value).toUpperCase();
                    if (upper === "TRUE") return true;
                    if (upper === "FALSE") return false;
                    return 0;
                }
                case "un": {
                    const v = evalAst(ast.expr, ctx);
                    const op = String(ast.op).toUpperCase();
                    if (op === "+") return toNumber(v);
                    if (op === "-") return -toNumber(v);
                    if (op === "NOT" || op === "!") return !toBool(v);
                    return toNumber(v);
                }
                case "bin": {
                    const l = evalAst(ast.left, ctx);
                    const r = evalAst(ast.right, ctx);
                    const op = String(ast.op).toUpperCase();
                    if (op === "+") return toNumber(l) + toNumber(r);
                    if (op === "-") return toNumber(l) - toNumber(r);
                    if (op === "*") return toNumber(l) * toNumber(r);
                    if (op === "/") return toNumber(l) / toNumber(r);
                    if (op === "^") return Math.pow(toNumber(l), toNumber(r));
                    if (op === ">") return toNumber(l) > toNumber(r);
                    if (op === "<") return toNumber(l) < toNumber(r);
                    if (op === ">=") return toNumber(l) >= toNumber(r);
                    if (op === "<=") return toNumber(l) <= toNumber(r);
                    if (op === "==") return toNumber(l) === toNumber(r);
                    if (op === "!=") return toNumber(l) !== toNumber(r);
                    if (op === "AND" || op === "&&") return toBool(l) && toBool(r);
                    if (op === "OR" || op === "||") return toBool(l) || toBool(r);
                    return 0;
                }
                case "call": {
                    const name = String(ast.name).toUpperCase();
                    const args = (ast.args || []).map((a: any) => evalAst(a, ctx));

                    const asRowCode = (x: any) => String(x ?? "").trim();

                    if (name === "VALUE") {
                        const code = asRowCode(args[0]);
                        const measure = args.length >= 2 ? asRowCode(args[1]) : "";
                        const period = args.length >= 3 ? asRowCode(args[2]) : "";
                        let colKey = ctx.currentColKey;
                        if (measure || period) {
                            if (period && measure) colKey = `${period}||${measure}`;
                            else if (measure) colKey = measure;
                            else colKey = period;
                        }
                        return getCell(code, colKey);
                    }

                    if (name === "IF") {
                        const cond = toBool(args[0]);
                        return cond ? toNumber(args[1]) : toNumber(args[2]);
                    }

                    if (name === "ABS") return Math.abs(toNumber(args[0]));

                    if (name === "ROUND") {
                        const n = toNumber(args[0]);
                        const d = args.length >= 2 ? Math.max(0, Math.floor(toNumber(args[1]))) : 0;
                        const p = Math.pow(10, d);
                        return Math.round(n * p) / p;
                    }

                    const agg = (mode: "SUM" | "AVG" | "MIN" | "MAX" | "COUNT") => {
                        const nums: number[] = [];
                        for (const a of args) {
                            if (typeof a === "string") {
                                const code = asRowCode(a);
                                const v = getCell(code, ctx.currentColKey);
                                if (v == null) continue;
                                nums.push(v);
                            } else {
                                const n = toNumber(a);
                                if (!Number.isFinite(n)) continue;
                                nums.push(n);
                            }
                        }
                        if (mode === "COUNT") return nums.length;
                        if (nums.length === 0) return 0;
                        if (mode === "SUM") return nums.reduce((s, x) => s + x, 0);
                        if (mode === "AVG") return nums.reduce((s, x) => s + x, 0) / nums.length;
                        if (mode === "MIN") return Math.min(...nums);
                        if (mode === "MAX") return Math.max(...nums);
                        return 0;
                    };

                    if (name === "SUM" || name === "AVG" || name === "MIN" || name === "MAX" || name === "COUNT") {
                        return agg(name as any);
                    }

                    if (name === "SUMCHILDREN" || name === "AVGCHILDREN" || name === "COUNTCHILDREN") {
                        const code = asRowCode(args[0]);
                        const node = rowByCode.get(code);
                        if (!node) return 0;
                        const childVals: number[] = [];
                        for (const c of node.children) {
                            const v = getCell(c, ctx.currentColKey);
                            if (v == null) continue;
                            childVals.push(v);
                        }
                        if (name === "COUNTCHILDREN") return childVals.length;
                        if (childVals.length === 0) return 0;
                        if (name === "SUMCHILDREN") return childVals.reduce((s, x) => s + x, 0);
                        return childVals.reduce((s, x) => s + x, 0) / childVals.length;
                    }

                    return 0;
                }
                default:
                    return 0;
            }
        };

        const calcRows = [...rowByCode.values()].filter(r => r.type === "calc" && !!r.formula);
        for (let pass = 0; pass < 6; pass++) {
            let changed = false;
            for (const r of calcRows) {
                let ast: any = null;
                try {
                    ast = new ExprParser(r.formula as string).parse();
                } catch {
                    ast = null;
                }
                for (const col of columns) {
                    if (!ast) {
                        setComputed(r.code, col.key, null);
                        continue;
                    }
                    const out = evalAst(ast, { currentColKey: col.key });
                    const num = toNumber(out);
                    const prev = valuesMap.get(r.code)?.get(col.key) ?? null;
                    const next = Number.isFinite(num) ? num : (blankAsZero ? 0 : null);
                    if (prev !== next) {
                        changed = true;
                        setComputed(r.code, col.key, next);
                    }
                }
            }
            if (!changed) break;
        }

        if (autoAggregateParents) {
            const postOrder = [...flatRows].reverse();
            for (const r of postOrder) {
                if (r.type === "blank" || r.type === "calc" || r.children.length === 0) continue;
                for (const col of columns) {
                    const existing = valuesMap.get(r.code)?.get(col.key);
                    if (existing != null) continue;
                    let sum = 0;
                    let hasAny = false;
                    for (const c of r.children) {
                        const cv = valuesMap.get(c)?.get(col.key);
                        if (cv == null) continue;
                        hasAny = true;
                        sum += cv;
                    }
                    setComputed(r.code, col.key, hasAny ? sum : (blankAsZero ? 0 : null));
                }
            }
        }

        // Inject subtotal rows at the end of each parent's child block.
        if (showSubtotals) {
            const out: RowNode[] = [];

            const pushSubtotalFor = (r: RowNode) => {
                if (r.rowLevel === -1) return; // skip Group-level subtotals
                if (r.type === "blank" || r.type === "calc") return;
                if (r.children.length === 0) return;

                const subtotalCode = `${r.code}||__subtotal`;
                const subtotalRow: RowNode = {
                    code: subtotalCode,
                    label: formatSubtotalLabel(r.label),
                    parent: r.parent,
                    type: "calc",
                    orderIndex: (r.orderIndex ?? 0) + 0.5,
                    order: (r.order ?? undefined),
                    formula: undefined,
                    style: undefined,
                    children: [],
                    depth: r.depth,
                    rowLevel: r.rowLevel,
                    groupKey: r.groupKey ?? null,
                    isTotal: true
                };

                if (!valuesMap.has(subtotalCode)) valuesMap.set(subtotalCode, new Map());
                for (const col of columns) {
                    let sum = 0;
                    let hasAny = false;
                    for (const c of r.children) {
                        const cv = valuesMap.get(c)?.get(col.key);
                        if (cv == null) continue;
                        hasAny = true;
                        sum += cv;
                    }
                    valuesMap.get(subtotalCode)!.set(col.key, hasAny ? sum : (blankAsZero ? 0 : null));
                }

                out.push(subtotalRow);
            };

            const visitWithSubtotals = (code: string) => {
                const r = rowByCode.get(code);
                if (!r) return;
                out.push(r);
                for (const c of r.children) {
                    visitWithSubtotals(c);
                }
                pushSubtotalFor(r);
            };

            out.splice(0, out.length);
            for (const rootCode of roots) {
                visitWithSubtotals(rootCode);
            }
            flatRows.splice(0, flatRows.length, ...out);
        }

        // Inject grand total row at bottom.
        if (showGrandTotal) {
            const gtCode = "__grand_total";
            const gtRow: RowNode = {
                code: gtCode,
                label: grandTotalLabel,
                parent: null,
                type: "calc",
                orderIndex: 1_000_000,
                children: [],
                depth: 0,
                rowLevel: undefined,
                groupKey: null,
                isTotal: true
            };
            if (!valuesMap.has(gtCode)) valuesMap.set(gtCode, new Map());

            // Sum all leaf/data rows with no children (data leaves) to avoid double counting parents.
            const sumRows = flatRows.filter(r => r.type !== "blank" && r.children.length === 0 && !r.isTotal);
            for (const col of columns) {
                let sum = 0;
                let hasAny = false;
                for (const r of sumRows) {
                    const v = valuesMap.get(r.code)?.get(col.key);
                    if (v == null) continue;
                    hasAny = true;
                    sum += v;
                }
                valuesMap.get(gtCode)!.set(col.key, hasAny ? sum : (blankAsZero ? 0 : null));
            }

            flatRows.push(gtRow);
        }

        const effectiveRowFieldCount = hierarchyView ? 1 : rowFieldCount;
        return {
            columns,
            rows: flatRows,
            values: valuesMap,
            valuesRaw: valuesRawMap,
            cfFieldRaw,
            rowFieldRaw,
            colFieldRaw,
            rowHeaderColumns,
            hasGroup,
            rowFieldCount: effectiveRowFieldCount,
            showColumnTotal,
            tooltipMeasureNames,
            rowSelectionIdByCode
        };
    }

    private buildModelFromCustomTable(dataView: DataView, cfg: CustomTableConfig): Model {
        const categorical = dataView.categorical;
        const categories = categorical?.categories ?? [];
        const values = categorical?.values ?? [];

        const hiddenColKeys = new Set((Array.isArray(cfg.hiddenColumnFieldKeys) ? cfg.hiddenColumnFieldKeys : [])
            .map(x => String(x ?? "").trim())
            .filter(x => !!x));

        const columnCatsAll = categories.filter(c => !!c.source.roles?.column);
        const columnCatValuesAll = columnCatsAll.map(c => (c.values ?? []) as any[]);
        const visibleColumnPairs = columnCatsAll
            .map((c, i) => {
                const src: any = c?.source;
                const key = String((src?.queryName || src?.displayName || "")).trim();
                const label = String((src?.displayName || src?.queryName || "")).trim();
                return { c, i, key, label: label || key, values: columnCatValuesAll[i] };
            })
            .filter(p => !!p.key && !hiddenColKeys.has(p.key));

        const columnCats = visibleColumnPairs.map(p => p.c);
        const columnCatValues = visibleColumnPairs.map(p => p.values);

        const norm = (v: any): string => {
            const s = v == null ? "" : String(v);
            const t = s.trim();
            return t ? t : "(Blank)";
        };

        const getColumnLevels = (i: number): string[] => {
            if (columnCatValues.length === 0) return [];
            return columnCatValues.map(arr => norm(arr[i]));
        };

        const rowCount = (() => {
            const anyArr = (columnCatValues[0] || []) as any[];
            const v0 = (values[0]?.values || []) as any[];
            return Math.max(anyArr.length || 0, v0.length || 0);
        })();

        // Group row indices by rendered column tuple.
        const idxByColTuple = new Map<string, number[]>();
        for (let i = 0; i < rowCount; i++) {
            const colLevels = getColumnLevels(i);
            const key = colLevels.join("||");
            if (!idxByColTuple.has(key)) idxByColTuple.set(key, []);
            idxByColTuple.get(key)!.push(i);
        }

        const expandByValue = cfg.showValueNamesInColumns === true;

        // Index measure columns by label.
        const measureByLabel = new Map<string, powerbi.DataViewValueColumn>();
        for (const v of values) {
            const src: any = v?.source;
            const name = String((src?.displayName || src?.queryName || "")).trim();
            if (!name) continue;
            const roles: any = src?.roles || {};
            const label = name;
            if (!measureByLabel.has(label)) measureByLabel.set(label, v);
        }

        // Index category arrays by display label.
        const categoryByLabel = new Map<string, any[]>();
        for (const c of categories) {
            const src: any = c?.source;
            const name = String((src?.displayName || src?.queryName || "")).trim();
            if (!name) continue;
            if (!categoryByLabel.has(name)) categoryByLabel.set(name, (c.values ?? []) as any[]);
        }

        const safeCodePart = (v: any): string => {
            const s = String(v ?? "");
            // Avoid accidental subtotal markers and our delimiter patterns.
            return s.replace(/\|\|/g, "|").replace(/:/g, ";");
        };

        const columns: BucketedColumnKey[] = [];
        const tupleKeys = Array.from(idxByColTuple.keys()).sort((a, b) => a.localeCompare(b));

        // Column keys for tuple-only mode.
        const colByKey = new Map<string, BucketedColumnKey>();
        const colKeyByTupleKey = new Map<string, string>();

        // Column keys for expanded-by-value mode.
        const colByExpandedKey = new Map<string, BucketedColumnKey>();

        const ensureTupleColumn = (colLevels: string[], tupleKey: string): BucketedColumnKey => {
            const leafLabel = colLevels.length > 0 ? colLevels[colLevels.length - 1] : "Value";
            const parentLevels = colLevels.length > 1 ? colLevels.slice(0, colLevels.length - 1) : [];
            const key = parentLevels.length > 0 ? `${parentLevels.join("||")}||${leafLabel}` : leafLabel;
            const existing = colByKey.get(key);
            if (existing) {
                colKeyByTupleKey.set(tupleKey, existing.key);
                return existing;
            }
            const col: BucketedColumnKey = {
                bucket: "values",
                measure: undefined,
                key,
                levels: [...parentLevels, leafLabel],
                columnLevels: parentLevels,
                leafLabel
            };
            colByKey.set(key, col);
            colKeyByTupleKey.set(tupleKey, col.key);
            columns.push(col);
            return col;
        };

        const ensureExpandedColumn = (colLevels: string[], tupleKey: string, fieldLabel: string): BucketedColumnKey => {
            const leafLabel = String(fieldLabel || "").trim() || "Value";
            const key = `${tupleKey}||__v:${leafLabel}`;
            const existing = colByExpandedKey.get(key);
            if (existing) return existing;
            const col: BucketedColumnKey = {
                bucket: "values",
                measure: undefined,
                key,
                levels: [...colLevels, leafLabel],
                columnLevels: colLevels.slice(),
                leafLabel
            };
            colByExpandedKey.set(key, col);
            columns.push(col);
            return col;
        };

        if (!expandByValue) {
            // Build columns as unique Column tuples only.
            for (const tupleKey of tupleKeys) {
                const colLevels = tupleKey ? tupleKey.split("||") : [];
                ensureTupleColumn(colLevels, tupleKey);
            }
        } else {
            // Build columns as (column tuple) x (selected value fields), showing value names in headers.
            const fieldOrder: string[] = [];
            const fieldSet = new Set<string>();

            for (const pp of (cfg.parents || [])) {
                for (const m of (Array.isArray((pp as any).values) ? (pp as any).values : [])) {
                    const f = String(m?.field || "").trim();
                    if (!f || fieldSet.has(f)) continue;
                    fieldSet.add(f);
                    fieldOrder.push(f);
                }
            }
            for (const ch of (cfg.children || [])) {
                for (const m of (Array.isArray((ch as any).values) ? (ch as any).values : [])) {
                    const f = String(m?.field || "").trim();
                    if (!f || fieldSet.has(f)) continue;
                    fieldSet.add(f);
                    fieldOrder.push(f);
                }
            }

            for (const tupleKey of tupleKeys) {
                const colLevels = tupleKey ? tupleKey.split("||") : [];
                for (const f of fieldOrder) {
                    ensureExpandedColumn(colLevels, tupleKey, f);
                }
            }
        }

        // Raw field maps for conditional formatting "Field value".
        const cfFieldRaw = new Map<string, any>();
        const rowFieldRaw = new Map<string, Map<string, any>>();
        const colFieldRaw = new Map<string, Map<string, any>>();
        const valuesMap = new Map<string, Map<string, number | null>>();
        const valuesRawMap = new Map<string, Map<string, any>>();

        const setCfFieldRaw = (rowCode: string, colLevels: string[], fieldLabel: string, v: any) => {
            const ll = (fieldLabel || "").trim();
            if (!ll) return;
            const key = `${rowCode}||${colLevels.join("||")}||${ll}`;
            cfFieldRaw.set(key, v);
        };
        const setRowFieldRaw = (rowCode: string, fieldLabel: string, v: any) => {
            const ll = (fieldLabel || "").trim();
            if (!ll) return;
            if (!rowFieldRaw.has(rowCode)) rowFieldRaw.set(rowCode, new Map());
            rowFieldRaw.get(rowCode)!.set(ll, v);
        };
        const setColFieldRaw = (colKey: string, fieldLabel: string, v: any) => {
            const ll = (fieldLabel || "").trim();
            if (!ll) return;
            if (!colFieldRaw.has(colKey)) colFieldRaw.set(colKey, new Map());
            colFieldRaw.get(colKey)!.set(ll, v);
        };
        const setCell = (code: string, colKey: string, v: number | null) => {
            if (!valuesMap.has(code)) valuesMap.set(code, new Map());
            valuesMap.get(code)!.set(colKey, v);
        };
        const setCellRaw = (code: string, colKey: string, v: any) => {
            if (!valuesRawMap.has(code)) valuesRawMap.set(code, new Map());
            valuesRawMap.get(code)!.set(colKey, v);
        };

        // Attach column group values to rendered column keys.
        // We pick any representative row index for that tuple.
        if (!expandByValue) {
            for (const tupleKey of tupleKeys) {
                const idx = (idxByColTuple.get(tupleKey) || [])[0];
                if (idx == null) continue;
                const colKey = colKeyByTupleKey.get(tupleKey);
                if (!colKey) continue;
                for (let ci = 0; ci < columnCats.length; ci++) {
                    const fieldName = visibleColumnPairs[ci]?.label || "";
                    if (!fieldName) continue;
                    setColFieldRaw(colKey, fieldName, (columnCatValues[ci] || [])[idx]);
                }
            }
        } else {
            for (const col of columns) {
                const parts = String(col.key || "").split("||__v:");
                const tupleKey = parts[0] || "";
                const idx = (idxByColTuple.get(tupleKey) || [])[0];
                if (idx == null) continue;
                for (let ci = 0; ci < columnCats.length; ci++) {
                    const fieldName = visibleColumnPairs[ci]?.label || "";
                    if (!fieldName) continue;
                    setColFieldRaw(col.key, fieldName, (columnCatValues[ci] || [])[idx]);
                }
            }
        }

        // Create data label/parent maps so we can reuse the existing row building flow.
        const dataLabelByCode = new Map<string, string>();
        const dataParentByCode = new Map<string, string | null>();
        const dataRowLevelByCode = new Map<string, number>();
        const dataGroupKeyByCode = new Map<string, string | null>();

        const customStyleByCode = new Map<string, RowStyleOverride>();
        const customFontFamilyByCode = new Map<string, string>();
        const customFontSizeByCode = new Map<string, number>();

        const parentCodeByNo = new Map<string, string>();
        for (const p of (cfg.parents || [])) {
            const no = String(p.parentNo || "").trim();
            if (!no) continue;
            const code = `CTP:${no}`;
            parentCodeByNo.set(no, code);
            dataLabelByCode.set(code, String(p.parentName || no).trim() || no);
            dataParentByCode.set(code, null);
            dataRowLevelByCode.set(code, 0);
            dataGroupKeyByCode.set(code, null);
            setRowFieldRaw(code, "Parent No", no);
            setRowFieldRaw(code, "Parent Name", String(p.parentName || "").trim());

            const fmt = p.format;
            if (fmt) {
                const st: RowStyleOverride = {};
                if (fmt.color) st.color = String(fmt.color).trim();
                if (fmt.bold != null) st.bold = fmt.bold === true;
                if (Object.keys(st).length > 0) customStyleByCode.set(code, st);

                const ff = String(fmt.fontFamily || "").trim();
                if (ff) customFontFamilyByCode.set(code, ff);
                if (fmt.fontSize != null && Number.isFinite(fmt.fontSize as any)) customFontSizeByCode.set(code, Number(fmt.fontSize));
            }

            // Optional parent-level value mappings
            const pMappings = (Array.isArray((p as any).values) ? ((p as any).values as any[]) : [])
                .map(m => ({ field: String(m?.field || "").trim(), aggregation: (m?.aggregation as CTAggregation) || "none" }))
                .filter(m => !!m.field);
            const effectivePMappings = expandByValue ? pMappings : (pMappings.length > 0 ? [pMappings[0]] : []);
            for (const mm of effectivePMappings) {
                const fieldLabel = mm.field;
                const isMeasure = this.customTableFieldIsMeasure.get(fieldLabel) === true;
                const agg = (isMeasure ? "none" : (mm.aggregation || "none"));

                for (const tupleKey of tupleKeys) {
                    const colLevels = tupleKey ? tupleKey.split("||") : [];
                    const colKey = expandByValue
                        ? `${tupleKey}||__v:${String(fieldLabel || "").trim() || "Value"}`
                        : (colKeyByTupleKey.get(tupleKey) || "");
                    if (!colKey) continue;
                    const idxs = idxByColTuple.get(tupleKey) || [];

                    let rawAny: any = null;
                    let num: number | null = null;

                    if (isMeasure) {
                        const vcol = measureByLabel.get(fieldLabel);
                        if (vcol) {
                            if (idxs.length === 1) {
                                rawAny = (vcol.values || [])[idxs[0]];
                            } else {
                                let sum = 0;
                                let has = false;
                                for (const ii of idxs) {
                                    const vv = (vcol.values || [])[ii];
                                    const n = toNumber(vv);
                                    if (!Number.isFinite(n)) continue;
                                    has = true;
                                    sum += n;
                                }
                                rawAny = has ? sum : null;
                            }
                            const n = toNumber(rawAny);
                            num = Number.isFinite(n) ? n : null;
                        }
                    } else {
                        const arr = categoryByLabel.get(fieldLabel);
                        if (arr) {
                            const nums: number[] = [];
                            let count = 0;
                            for (const ii of idxs) {
                                const vv = (arr || [])[ii];
                                if (vv == null || (typeof vv === "string" && vv.trim() === "")) {
                                    // skip
                                } else {
                                    count++;
                                }
                                const n = toNumber(vv);
                                if (Number.isFinite(n)) nums.push(n);
                            }

                            if (agg === "count") {
                                rawAny = count;
                                num = count;
                            } else if (agg === "none") {
                                rawAny = (arr || [])[idxs[0] ?? 0];
                                const n = toNumber(rawAny);
                                num = Number.isFinite(n) ? n : null;
                            } else if (nums.length > 0) {
                                if (agg === "sum") num = nums.reduce((s, x) => s + x, 0);
                                else if (agg === "avg") num = nums.reduce((s, x) => s + x, 0) / nums.length;
                                else if (agg === "min") num = Math.min(...nums);
                                else if (agg === "max") num = Math.max(...nums);
                                else num = null;
                                rawAny = num;
                            }
                        }
                    }

                    if (rawAny != null || num != null) {
                        setCellRaw(code, colKey, rawAny);
                        setCell(code, colKey, num);
                    }
                    setCfFieldRaw(code, colLevels, fieldLabel, rawAny);
                }
            }
        }

        for (const c of (cfg.children || [])) {
            const autoChildField = String((c as any).childNameFromField || "").trim();
            if (autoChildField) {
                const parentNo = String(c.setParentNo || "").trim();
                const parentCode = parentNo ? (parentCodeByNo.get(parentNo) || null) : null;
                const childArr = categoryByLabel.get(autoChildField) || null;
                const parentMatchField = String((c as any).parentMatchField || "").trim();
                const parentArr = parentMatchField ? (categoryByLabel.get(parentMatchField) || null) : null;
                if (!childArr) continue;

                const parentNoNorm = parentNo ? norm(parentNo) : "";
                const distinct = new Set<string>();
                for (let i = 0; i < rowCount; i++) {
                    if (parentArr && parentNoNorm) {
                        if (norm((parentArr as any[])[i]) !== parentNoNorm) continue;
                    }
                    const cv = norm((childArr as any[])[i]);
                    if (!cv) continue;
                    distinct.add(cv);
                }
                const childValues = Array.from(distinct.values()).sort((a, b) => a.localeCompare(b));
                for (const childValue of childValues) {
                    const code = `CTA:${safeCodePart(c.id)}:${safeCodePart(parentNo)}:${safeCodePart(autoChildField)}:${safeCodePart(childValue)}`;
                    dataLabelByCode.set(code, childValue);
                    dataParentByCode.set(code, parentCode);
                    dataRowLevelByCode.set(code, parentCode ? 1 : 0);
                    dataGroupKeyByCode.set(code, null);
                    setRowFieldRaw(code, "Set Parent No", parentNo);
                    setRowFieldRaw(code, "Child Name", childValue);
                    setRowFieldRaw(code, "Child Name Field", autoChildField);
                    if (parentMatchField) setRowFieldRaw(code, "Parent Match Field", parentMatchField);

                    const fmt = c.format;
                    if (fmt) {
                        const st: RowStyleOverride = {};
                        if (fmt.color) st.color = String(fmt.color).trim();
                        if (fmt.bold != null) st.bold = fmt.bold === true;
                        if (Object.keys(st).length > 0) customStyleByCode.set(code, st);

                        const ff = String(fmt.fontFamily || "").trim();
                        if (ff) customFontFamilyByCode.set(code, ff);
                        if (fmt.fontSize != null && Number.isFinite(fmt.fontSize as any)) customFontSizeByCode.set(code, Number(fmt.fontSize));
                    }

                    const mappings = (Array.isArray(c.values) ? c.values : [])
                        .map(m => ({ field: String(m.field || "").trim(), aggregation: (m.aggregation as CTAggregation) || "none" }))
                        .filter(m => !!m.field);
                    const effectiveMappings = expandByValue ? mappings : (mappings.length > 0 ? [mappings[0]] : []);

                    for (const mm of effectiveMappings) {
                        const fieldLabel = mm.field;
                        const isMeasure = this.customTableFieldIsMeasure.get(fieldLabel) === true;
                        const agg = (isMeasure ? "none" : (mm.aggregation || "none"));

                        for (const tupleKey of tupleKeys) {
                            const colLevels = tupleKey ? tupleKey.split("||") : [];
                            const colKey = expandByValue
                                ? `${tupleKey}||__v:${String(fieldLabel || "").trim() || "Value"}`
                                : (colKeyByTupleKey.get(tupleKey) || "");
                            if (!colKey) continue;

                            const idxsBase = idxByColTuple.get(tupleKey) || [];
                            const idxs: number[] = [];
                            for (const ii of idxsBase) {
                                if (norm((childArr as any[])[ii]) !== childValue) continue;
                                if (parentArr && parentNoNorm) {
                                    if (norm((parentArr as any[])[ii]) !== parentNoNorm) continue;
                                }
                                idxs.push(ii);
                            }
                            if (idxs.length === 0) {
                                setCfFieldRaw(code, colLevels, fieldLabel, null);
                                continue;
                            }

                            let rawAny: any = null;
                            let num: number | null = null;

                            if (isMeasure) {
                                const vcol = measureByLabel.get(fieldLabel);
                                if (vcol) {
                                    if (idxs.length === 1) {
                                        rawAny = (vcol.values || [])[idxs[0]];
                                    } else {
                                        let sum = 0;
                                        let has = false;
                                        for (const jj of idxs) {
                                            const vv = (vcol.values || [])[jj];
                                            const n = toNumber(vv);
                                            if (!Number.isFinite(n)) continue;
                                            has = true;
                                            sum += n;
                                        }
                                        rawAny = has ? sum : null;
                                    }
                                    const n = toNumber(rawAny);
                                    num = Number.isFinite(n) ? n : null;
                                }
                            } else {
                                const arr = categoryByLabel.get(fieldLabel);
                                if (arr) {
                                    const nums: number[] = [];
                                    let count = 0;
                                    for (const jj of idxs) {
                                        const vv = (arr || [])[jj];
                                        if (vv == null || (typeof vv === "string" && vv.trim() === "")) {
                                            // skip
                                        } else {
                                            count++;
                                        }
                                        const n = toNumber(vv);
                                        if (Number.isFinite(n)) nums.push(n);
                                    }

                                    if (agg === "count") {
                                        rawAny = count;
                                        num = count;
                                    } else if (agg === "none") {
                                        rawAny = (arr || [])[idxs[0] ?? 0];
                                        const n = toNumber(rawAny);
                                        num = Number.isFinite(n) ? n : null;
                                    } else if (nums.length > 0) {
                                        if (agg === "sum") num = nums.reduce((s, x) => s + x, 0);
                                        else if (agg === "avg") num = nums.reduce((s, x) => s + x, 0) / nums.length;
                                        else if (agg === "min") num = Math.min(...nums);
                                        else if (agg === "max") num = Math.max(...nums);
                                        else num = null;
                                        rawAny = num;
                                    }
                                }
                            }

                            if (rawAny != null || num != null) {
                                setCellRaw(code, colKey, rawAny);
                                setCell(code, colKey, num);
                            }
                            setCfFieldRaw(code, colLevels, fieldLabel, rawAny);
                        }
                    }
                }
                continue;
            }

            const code = `CTC:${String(c.id || "").trim() || cryptoId("ctc")}`;
            const parentNo = String(c.setParentNo || "").trim();
            const parentCode = parentNo ? (parentCodeByNo.get(parentNo) || null) : null;
            dataLabelByCode.set(code, String(c.childName || "").trim() || "(Blank)");
            dataParentByCode.set(code, parentCode);
            dataRowLevelByCode.set(code, parentCode ? 1 : 0);
            dataGroupKeyByCode.set(code, null);
            setRowFieldRaw(code, "Set Parent No", parentNo);
            setRowFieldRaw(code, "Child Name", String(c.childName || "").trim());

            const fmt = c.format;
            if (fmt) {
                const st: RowStyleOverride = {};
                if (fmt.color) st.color = String(fmt.color).trim();
                if (fmt.bold != null) st.bold = fmt.bold === true;
                if (Object.keys(st).length > 0) customStyleByCode.set(code, st);

                const ff = String(fmt.fontFamily || "").trim();
                if (ff) customFontFamilyByCode.set(code, ff);
                if (fmt.fontSize != null && Number.isFinite(fmt.fontSize as any)) customFontSizeByCode.set(code, Number(fmt.fontSize));
            }

            const mappings = (Array.isArray(c.values) ? c.values : [])
                .map(m => ({ field: String(m.field || "").trim(), aggregation: (m.aggregation as CTAggregation) || "none" }))
                .filter(m => !!m.field);

            const effectiveMappings = expandByValue ? mappings : (mappings.length > 0 ? [mappings[0]] : []);

            for (const mm of effectiveMappings) {
                const fieldLabel = mm.field;
                const isMeasure = this.customTableFieldIsMeasure.get(fieldLabel) === true;
                const agg = (isMeasure ? "none" : (mm.aggregation || "none"));

                for (const tupleKey of tupleKeys) {
                    const colLevels = tupleKey ? tupleKey.split("||") : [];
                    const colKey = expandByValue
                        ? `${tupleKey}||__v:${String(fieldLabel || "").trim() || "Value"}`
                        : (colKeyByTupleKey.get(tupleKey) || "");
                    if (!colKey) continue;
                    const idxs = idxByColTuple.get(tupleKey) || [];

                    let rawAny: any = null;
                    let num: number | null = null;

                    if (isMeasure) {
                        const vcol = measureByLabel.get(fieldLabel);
                        if (vcol) {
                            if (idxs.length === 1) {
                                rawAny = (vcol.values || [])[idxs[0]];
                            } else {
                                let sum = 0;
                                let has = false;
                                for (const ii of idxs) {
                                    const vv = (vcol.values || [])[ii];
                                    const n = toNumber(vv);
                                    if (!Number.isFinite(n)) continue;
                                    has = true;
                                    sum += n;
                                }
                                rawAny = has ? sum : null;
                            }
                            const n = toNumber(rawAny);
                            num = Number.isFinite(n) ? n : null;
                        }
                    } else {
                        const arr = categoryByLabel.get(fieldLabel);
                        if (arr) {
                            const nums: number[] = [];
                            let count = 0;
                            for (const ii of idxs) {
                                const vv = (arr || [])[ii];
                                if (vv == null || (typeof vv === "string" && vv.trim() === "")) {
                                    // skip
                                } else {
                                    count++;
                                }
                                const n = toNumber(vv);
                                if (Number.isFinite(n)) nums.push(n);
                            }

                            if (agg === "count") {
                                rawAny = count;
                                num = count;
                            } else if (agg === "none") {
                                rawAny = (arr || [])[idxs[0] ?? 0];
                                const n = toNumber(rawAny);
                                num = Number.isFinite(n) ? n : null;
                            } else if (nums.length > 0) {
                                if (agg === "sum") num = nums.reduce((s, x) => s + x, 0);
                                else if (agg === "avg") num = nums.reduce((s, x) => s + x, 0) / nums.length;
                                else if (agg === "min") num = Math.min(...nums);
                                else if (agg === "max") num = Math.max(...nums);
                                else num = null;
                                rawAny = num;
                            }
                        }
                    }

                    if (rawAny != null || num != null) {
                        setCellRaw(code, colKey, rawAny);
                        setCell(code, colKey, num);
                    }
                    setCfFieldRaw(code, colLevels, fieldLabel, rawAny);
                }
            }
        }

        // Sort columns by their level tuple.
        columns.sort((a, b) => {
            const len = Math.max(a.levels.length, b.levels.length);
            for (let i = 0; i < len; i++) {
                const av = a.levels[i] ?? "";
                const bv = b.levels[i] ?? "";
                const c = av.localeCompare(bv);
                if (c !== 0) return c;
            }
            return 0;
        });

        // IMPORTANT: Do NOT pre-fill unmapped (row, column) cells.
        // If we pre-fill with null, the rendering layer (blankAsZero) turns them into 0,
        // which looks like an extra "0" column before the actual mapped value.
        // We only ensure the row map exists.
        for (const code of dataLabelByCode.keys()) {
            if (!valuesMap.has(code)) valuesMap.set(code, new Map());
            if (!valuesRawMap.has(code)) valuesRawMap.set(code, new Map());
        }

        // Reuse the existing layout + totals machinery by continuing with the same row construction logic.
        const layoutRows = this.parseLayoutRows(this.formattingSettings.layout.layoutJson.value);
        const formulas = this.parseFormulas(this.formattingSettings.layout.formulasJson.value);

        const rowByCode = new Map<string, RowNode>();
        let orderIndex = 0;
        const addRow = (r: LayoutRow, fallbackLabel?: string, fallbackParent?: string | null): RowNode | null => {
            const code = (r.code || "").trim();
            if (!code) return null;
            const label = (r.label || fallbackLabel || code).trim();
            const parent = r.parent !== undefined ? (r.parent ? r.parent.trim() : null) : (fallbackParent ?? null);
            const type: LayoutRowType = (r.type || "data") as LayoutRowType;
            const formula = (r.formula || formulas.get(code) || undefined);
            const out: RowNode = {
                code,
                label,
                parent,
                type,
                orderIndex: orderIndex++,
                order: r.order,
                formula,
                style: r.style,
                children: [],
                depth: 0,
                rowLevel: undefined,
                groupKey: null
            };
            rowByCode.set(code, out);
            return out;
        };

        const applyCustomOverrides = (row: RowNode) => {
            const code = row.code;
            const st = customStyleByCode.get(code);
            if (st) {
                row.style = { ...(row.style || {}), ...st };
            }
            const ff = customFontFamilyByCode.get(code);
            if (ff) row.fontFamilyOverride = ff;
            const fs = customFontSizeByCode.get(code);
            if (fs != null && Number.isFinite(fs as any)) row.fontSizeOverride = fs;
        };

        for (const r of layoutRows) addRow(r, undefined, null);

        // Add our custom table rows.
        for (const [code, label] of dataLabelByCode.entries()) {
            if (rowByCode.has(code)) {
                const existing = rowByCode.get(code)!;
                if (!existing.parent) existing.parent = dataParentByCode.get(code) ?? null;
                if (!existing.label) existing.label = label;
                applyCustomOverrides(existing);
                continue;
            }
            const created = addRow({ code, type: "data" } as any, label, dataParentByCode.get(code) ?? null);
            if (created) applyCustomOverrides(created);
        }

        // Link children.
        for (const r of rowByCode.values()) {
            if (r.parent && rowByCode.has(r.parent)) {
                rowByCode.get(r.parent)!.children.push(r.code);
            }
        }

        const roots: string[] = [];
        for (const r of rowByCode.values()) {
            if (!r.parent || !rowByCode.has(r.parent)) roots.push(r.code);
        }

        const sortChildren = (codes: string[]) => {
            codes.sort((a, b) => {
                const ra = rowByCode.get(a)!;
                const rb = rowByCode.get(b)!;
                const oa = ra.order ?? ra.orderIndex;
                const ob = rb.order ?? rb.orderIndex;
                if (oa !== ob) return oa - ob;
                return ra.label.localeCompare(rb.label);
            });
        };
        for (const r of rowByCode.values()) sortChildren(r.children);
        sortChildren(roots);

        const flatRows: RowNode[] = [];
        const visit = (code: string, depth: number) => {
            const r = rowByCode.get(code);
            if (!r) return;
            r.depth = depth;
            flatRows.push(r);
            for (const c of r.children) visit(c, depth + 1);
        };
        for (const root of roots) visit(root, 0);

        // Totals settings (same as default model)
        const totalsAny = ((this.formattingSettings as any).totals as any) || null;
        const rowsAny = (this.formattingSettings.rows as any);
        const showGrandTotal = ((totalsAny?.showGrandTotal?.value as boolean | undefined) ?? (rowsAny.showGrandTotal?.value as boolean | undefined)) ?? false;
        const showSubtotals = ((totalsAny?.showSubtotals?.value as boolean | undefined) ?? (rowsAny.showSubtotals?.value as boolean | undefined)) ?? false;
        const showColumnTotal = ((totalsAny?.showColumnTotal?.value as boolean | undefined) ?? (rowsAny.showColumnTotal?.value as boolean | undefined)) ?? false;

        const grandTotalLabel = (((totalsAny?.grandTotalLabel?.value as string | undefined) ?? (rowsAny.grandTotalLabel?.value as string | undefined) ?? "Grand Total").trim()) || "Grand Total";
        const subtotalLabelTemplate = (((totalsAny?.subtotalLabelTemplate?.value as string | undefined) ?? (rowsAny.subtotalLabelTemplate?.value as string | undefined) ?? "Total {label}").trim()) || "Total {label}";
        const formatSubtotalLabel = (label: string): string => {
            const base = String(label ?? "").trim();
            const tpl = subtotalLabelTemplate;
            if (tpl.includes("{label}")) {
                const out = tpl.replace(/\{label\}/g, base);
                return out.trim() || `Total ${base}`;
            }
            return tpl.trim();
        };

        const blankAsZero = this.formattingSettings.rows.blankAsZero.value;
        const autoAggregateParents = this.formattingSettings.rows.autoAggregateParents.value;

        const setComputed = (code: string, colKey: string, v: number | null) => {
            if (!valuesMap.has(code)) valuesMap.set(code, new Map());
            valuesMap.get(code)!.set(colKey, v);
        };

        // Evaluate formulas (reuses same helper in default buildModel)
        const getCell = (code: string, colKey: string): number | null => {
            const v = valuesMap.get(code)?.get(colKey);
            if (v == null) return blankAsZero ? 0 : null;
            return v;
        };
        const evalAst = (ast: any, ctx: { currentColKey: string }): any => {
            switch (ast.type) {
                case "num": return ast.value;
                case "str": return ast.value;
                case "ref": return getCell(ast.value, ctx.currentColKey);
                case "ident": {
                    const upper = String(ast.value).toUpperCase();
                    if (upper === "TRUE") return true;
                    if (upper === "FALSE") return false;
                    return 0;
                }
                case "un": {
                    const v = evalAst(ast.expr, ctx);
                    const op = String(ast.op).toUpperCase();
                    if (op === "+") return toNumber(v);
                    if (op === "-") return -toNumber(v);
                    if (op === "NOT" || op === "!") return !toBool(v);
                    return toNumber(v);
                }
                case "bin": {
                    const l = evalAst(ast.left, ctx);
                    const r = evalAst(ast.right, ctx);
                    const op = String(ast.op).toUpperCase();
                    if (op === "+") return toNumber(l) + toNumber(r);
                    if (op === "-") return toNumber(l) - toNumber(r);
                    if (op === "*") return toNumber(l) * toNumber(r);
                    if (op === "/") return toNumber(l) / toNumber(r);
                    if (op === "^") return Math.pow(toNumber(l), toNumber(r));
                    if (op === ">") return toNumber(l) > toNumber(r);
                    if (op === "<") return toNumber(l) < toNumber(r);
                    if (op === ">=") return toNumber(l) >= toNumber(r);
                    if (op === "<=") return toNumber(l) <= toNumber(r);
                    if (op === "==") return toNumber(l) === toNumber(r);
                    if (op === "!=") return toNumber(l) !== toNumber(r);
                    if (op === "AND" || op === "&&") return toBool(l) && toBool(r);
                    if (op === "OR" || op === "||") return toBool(l) || toBool(r);
                    return 0;
                }
                case "call": {
                    const name = String(ast.name).toUpperCase();
                    const args = (ast.args || []).map((a: any) => evalAst(a, ctx));
                    const asRowCode = (x: any) => String(x ?? "").trim();
                    if (name === "VALUE") {
                        const code = asRowCode(args[0]);
                        const measure = args.length >= 2 ? asRowCode(args[1]) : "";
                        const period = args.length >= 3 ? asRowCode(args[2]) : "";
                        let colKey = ctx.currentColKey;
                        if (measure || period) {
                            if (period && measure) colKey = `${period}||${measure}`;
                            else if (measure) colKey = measure;
                            else colKey = period;
                        }
                        return getCell(code, colKey);
                    }
                    if (name === "IF") {
                        const cond = toBool(args[0]);
                        return cond ? toNumber(args[1]) : toNumber(args[2]);
                    }
                    if (name === "ABS") return Math.abs(toNumber(args[0]));
                    return 0;
                }
                default:
                    return 0;
            }
        };

        const calcRows = [...rowByCode.values()].filter(r => r.type === "calc" && !!r.formula);
        for (let pass = 0; pass < 6; pass++) {
            let changed = false;
            for (const r of calcRows) {
                let ast: any = null;
                try {
                    ast = new ExprParser(r.formula as string).parse();
                } catch {
                    ast = null;
                }
                for (const col of columns) {
                    if (!ast) {
                        setComputed(r.code, col.key, null);
                        continue;
                    }
                    const out = evalAst(ast, { currentColKey: col.key });
                    const num = toNumber(out);
                    const prev = valuesMap.get(r.code)?.get(col.key) ?? null;
                    const next = Number.isFinite(num) ? num : (blankAsZero ? 0 : null);
                    if (prev !== next) {
                        changed = true;
                        setComputed(r.code, col.key, next);
                    }
                }
            }
            if (!changed) break;
        }

        if (autoAggregateParents) {
            const postOrder = [...flatRows].reverse();
            for (const r of postOrder) {
                if (r.type === "blank" || r.type === "calc" || r.children.length === 0) continue;
                for (const col of columns) {
                    const existing = valuesMap.get(r.code)?.get(col.key);
                    if (existing != null) continue;
                    let sum = 0;
                    let hasAny = false;
                    for (const c of r.children) {
                        const cv = valuesMap.get(c)?.get(col.key);
                        if (cv == null) continue;
                        hasAny = true;
                        sum += cv;
                    }
                    setComputed(r.code, col.key, hasAny ? sum : (blankAsZero ? 0 : null));
                }
            }
        }

        if (showSubtotals) {
            const outRows: RowNode[] = [];
            const pushSubtotalFor = (r: RowNode) => {
                if (r.type === "blank" || r.type === "calc") return;
                if (r.children.length === 0) return;
                const subtotalCode = `${r.code}||__subtotal`;
                const subtotalRow: RowNode = {
                    code: subtotalCode,
                    label: formatSubtotalLabel(r.label),
                    parent: r.parent,
                    type: "calc",
                    orderIndex: (r.orderIndex ?? 0) + 0.5,
                    order: (r.order ?? undefined),
                    formula: undefined,
                    style: undefined,
                    children: [],
                    depth: r.depth,
                    rowLevel: r.rowLevel,
                    groupKey: r.groupKey ?? null,
                    isTotal: true
                };
                if (!valuesMap.has(subtotalCode)) valuesMap.set(subtotalCode, new Map());
                for (const col of columns) {
                    let sum = 0;
                    let hasAny = false;
                    for (const c of r.children) {
                        const cv = valuesMap.get(c)?.get(col.key);
                        if (cv == null) continue;
                        hasAny = true;
                        sum += cv;
                    }
                    valuesMap.get(subtotalCode)!.set(col.key, hasAny ? sum : (blankAsZero ? 0 : null));
                }
                outRows.push(subtotalRow);
            };
            const visitWithSubtotals = (code: string) => {
                const r = rowByCode.get(code);
                if (!r) return;
                outRows.push(r);
                for (const c of r.children) visitWithSubtotals(c);
                pushSubtotalFor(r);
            };
            outRows.splice(0, outRows.length);
            for (const root of roots) visitWithSubtotals(root);
            flatRows.splice(0, flatRows.length, ...outRows);
        }

        if (showGrandTotal) {
            const gtCode = "__grand_total";
            const gtRow: RowNode = {
                code: gtCode,
                label: grandTotalLabel,
                parent: null,
                type: "calc",
                orderIndex: 1_000_000,
                children: [],
                depth: 0,
                rowLevel: undefined,
                groupKey: null,
                isTotal: true
            };
            if (!valuesMap.has(gtCode)) valuesMap.set(gtCode, new Map());
            const sumRows = flatRows.filter(r => r.type !== "blank" && r.children.length === 0 && !r.isTotal);
            for (const col of columns) {
                let sum = 0;
                let hasAny = false;
                for (const r of sumRows) {
                    const v = valuesMap.get(r.code)?.get(col.key);
                    if (v == null) continue;
                    hasAny = true;
                    sum += v;
                }
                valuesMap.get(gtCode)!.set(col.key, hasAny ? sum : (blankAsZero ? 0 : null));
            }
            flatRows.push(gtRow);
        }

        const rowHeaderColumns: string[] = [""];
        return {
            columns,
            rows: flatRows,
            values: valuesMap,
            valuesRaw: valuesRawMap,
            cfFieldRaw,
            rowFieldRaw,
            colFieldRaw,
            rowHeaderColumns,
            hasGroup: false,
            rowFieldCount: 1,
            showColumnTotal
        };
    }

    private parseLayoutRows(raw: string): LayoutRow[] {
        const text = (raw || "").trim();
        if (!text) return [];
        try {
            const parsed = JSON.parse(text);
            if (Array.isArray(parsed)) return parsed as LayoutRow[];
            const obj = parsed as LayoutJson;
            if (obj && Array.isArray(obj.rows)) return obj.rows;
            return [];
        } catch {
            return [];
        }
    }

    private parseFormulas(raw: string): Map<string, string> {
        const text = (raw || "").trim();
        const out = new Map<string, string>();
        if (!text) return out;
        try {
            const parsed = JSON.parse(text);
            if (!parsed || typeof parsed !== "object") return out;
            for (const [k, v] of Object.entries(parsed)) {
                const code = String(k ?? "").trim();
                const expr = String(v ?? "").trim();
                if (code && expr) out.set(code, expr);
            }
        } catch {
            return out;
        }
        return out;
    }

    private escapeHtml(s: string): string {
        return (s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }
}
