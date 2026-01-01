import powerbi from "powerbi-visuals-api";

function formatTooltipValue(value: unknown): string {
    if (value == null) return "";
    if (typeof value === "string") return value;
    if (typeof value === "number") {
        return Number.isFinite(value) ? value.toLocaleString() : "";
    }
    if (typeof value === "boolean") return value ? "True" : "False";
    try {
        return String(value);
    } catch {
        return "";
    }
}

export function buildMatrixTooltipItems(args: {
    rowLabel?: string;
    columnLabel?: string;
    valueText?: string;
    extra?: Array<{ name: string; value: unknown }>;
}): powerbi.extensibility.VisualTooltipDataItem[] {
    const items: powerbi.extensibility.VisualTooltipDataItem[] = [];

    const rowLabel = (args.rowLabel ?? "").toString().trim();
    if (rowLabel) items.push({ displayName: "Row", value: rowLabel });

    const columnLabel = (args.columnLabel ?? "").toString().trim();
    if (columnLabel) items.push({ displayName: "Column", value: columnLabel });

    const valueText = (args.valueText ?? "").toString().trim();
    if (valueText) items.push({ displayName: "Value", value: valueText });

    const extra = Array.isArray(args.extra) ? args.extra : [];
    for (const e of extra) {
        const name = (e?.name ?? "").toString().trim();
        if (!name) continue;
        items.push({ displayName: name, value: formatTooltipValue(e?.value) });
    }

    return items;
}
