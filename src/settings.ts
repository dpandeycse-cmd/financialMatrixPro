"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

const displayUnitOptions: powerbi.IEnumMember[] = [
    { value: "auto", displayName: "Auto" },
    { value: "none", displayName: "None" },
    { value: "thousands", displayName: "Thousands" },
    { value: "millions", displayName: "Millions" },
    { value: "billions", displayName: "Billions" }
];

const tablePresetOptions: powerbi.IEnumMember[] = [
    // Modern attractive finance (gradient header + tinted rows)
    { value: "financeBlueZebra", displayName: "Finance blue (tinted rows)" },
    { value: "financePurpleZebra", displayName: "Finance purple (tinted rows)" },
    { value: "financeEmeraldZebra", displayName: "Finance emerald (tinted rows)" },

    { value: "custom", displayName: "Custom" },
    { value: "clean", displayName: "Clean" },
    { value: "compact", displayName: "Compact" },
    // Curated distinct presets (use the separate 'Zebra rows' toggle when you want alternating rows)
    { value: "minimal", displayName: "Minimal" },
    { value: "ledger", displayName: "Ledger" },
    { value: "statement", displayName: "Statement" },
    { value: "spreadsheet", displayName: "Spreadsheet" },
    { value: "boxed", displayName: "Boxed" },

    { value: "card", displayName: "Card" },
    { value: "underlineHeader", displayName: "Underline header" },

    { value: "premiumSoft", displayName: "Premium soft cards" },
    { value: "premiumAccent", displayName: "Premium accent stripe" },
    { value: "premiumCompact", displayName: "Premium compact" }
];

const columnWidthModeOptions: powerbi.IEnumMember[] = [
    { value: "auto", displayName: "Auto" },
    { value: "manual", displayName: "Manual" }
];

const fontFamilyOptions: powerbi.IEnumMember[] = [
    { value: "Segoe UI", displayName: "Segoe UI" },
    { value: "Arial", displayName: "Arial" },
    { value: "Calibri", displayName: "Calibri" }
];

const textAlignOptions: powerbi.IEnumMember[] = [
    { value: "left", displayName: "Left" },
    { value: "center", displayName: "Center" },
    { value: "right", displayName: "Right" }
];

const gridLineStyleOptions: powerbi.IEnumMember[] = [
    { value: "solid", displayName: "Solid" },
    { value: "dashed", displayName: "Dashed" },
    { value: "dotted", displayName: "Dotted" }
];

class TableCardSettings extends FormattingSettingsCard {
    preset = new formattingSettings.ItemDropdown({
        name: "preset",
        displayName: "Preset",
        value: tablePresetOptions.find(x => x.value === "financeBlueZebra") || tablePresetOptions[0],
        items: tablePresetOptions
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background",
        value: { value: "" }
    });

    showGrid = new formattingSettings.ToggleSwitch({
        name: "showGrid",
        displayName: "Show grid",
        value: true
    });

    horizontalGrid = new formattingSettings.ToggleSwitch({ name: "horizontalGrid", displayName: "Horizontal grid", value: true });
    horizontalGridColor = new formattingSettings.ColorPicker({ name: "horizontalGridColor", displayName: "Horizontal grid color", value: { value: "" } });
    horizontalGridThickness = new formattingSettings.NumUpDown({ name: "horizontalGridThickness", displayName: "Horizontal grid thickness", value: 1 });
    horizontalGridStyle = new formattingSettings.ItemDropdown({ name: "horizontalGridStyle", displayName: "Horizontal grid style", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    verticalGrid = new formattingSettings.ToggleSwitch({ name: "verticalGrid", displayName: "Vertical grid", value: true });
    verticalGridColor = new formattingSettings.ColorPicker({ name: "verticalGridColor", displayName: "Vertical grid color", value: { value: "" } });
    verticalGridThickness = new formattingSettings.NumUpDown({ name: "verticalGridThickness", displayName: "Vertical grid thickness", value: 1 });
    verticalGridStyle = new formattingSettings.ItemDropdown({ name: "verticalGridStyle", displayName: "Vertical grid style", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    // Totals grid settings moved to Totals card (keep legacy properties for backward compatibility; not shown in Table pane)
    grandTotalGrid = new formattingSettings.ToggleSwitch({ name: "grandTotalGrid", displayName: "Grand total grid (legacy)", value: false });
    grandTotalGridColor = new formattingSettings.ColorPicker({ name: "grandTotalGridColor", displayName: "Grand total grid color (legacy)", value: { value: "" } });
    grandTotalGridThickness = new formattingSettings.NumUpDown({ name: "grandTotalGridThickness", displayName: "Grand total grid thickness (legacy)", value: 2 });
    grandTotalGridStyle = new formattingSettings.ItemDropdown({ name: "grandTotalGridStyle", displayName: "Grand total grid style (legacy)", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    subtotalGrid = new formattingSettings.ToggleSwitch({ name: "subtotalGrid", displayName: "Subtotal grid (legacy)", value: false });
    subtotalGridColor = new formattingSettings.ColorPicker({ name: "subtotalGridColor", displayName: "Subtotal grid color (legacy)", value: { value: "" } });
    subtotalGridThickness = new formattingSettings.NumUpDown({ name: "subtotalGridThickness", displayName: "Subtotal grid thickness (legacy)", value: 1 });
    subtotalGridStyle = new formattingSettings.ItemDropdown({ name: "subtotalGridStyle", displayName: "Subtotal grid style (legacy)", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    columnTotalGrid = new formattingSettings.ToggleSwitch({ name: "columnTotalGrid", displayName: "Column total grid (legacy)", value: false });
    columnTotalGridColor = new formattingSettings.ColorPicker({ name: "columnTotalGridColor", displayName: "Column total grid color (legacy)", value: { value: "" } });
    columnTotalGridThickness = new formattingSettings.NumUpDown({ name: "columnTotalGridThickness", displayName: "Column total grid thickness (legacy)", value: 2 });
    columnTotalGridStyle = new formattingSettings.ItemDropdown({ name: "columnTotalGridStyle", displayName: "Column total grid style (legacy)", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    zebra = new formattingSettings.ToggleSwitch({
        name: "zebra",
        displayName: "Zebra rows",
        value: false
    });

    columnWidthMode = new formattingSettings.ItemDropdown({
        name: "columnWidthMode",
        displayName: "Column width",
        value: columnWidthModeOptions[0],
        items: columnWidthModeOptions
    });

    columnWidthPx = new formattingSettings.NumUpDown({
        name: "columnWidthPx",
        displayName: "Column width (px)",
        value: 120
    });

    rowHeaderWidthPx = new formattingSettings.NumUpDown({
        name: "rowHeaderWidthPx",
        displayName: "Row header width (px)",
        value: 220
    });

    wrapText = new formattingSettings.ToggleSwitch({
        name: "wrapText",
        displayName: "Wrap text (no hide)",
        value: true
    });

    freezeHeader = new formattingSettings.ToggleSwitch({
        name: "freezeHeader",
        displayName: "Freeze header while scrolling",
        value: false
    });

    freezeRowHeaders = new formattingSettings.ToggleSwitch({
        name: "freezeRowHeaders",
        displayName: "Freeze row headers while scrolling",
        value: false
    });

    name: string = "table";
    displayName: string = "Table";
    slices: Array<FormattingSettingsSlice> = [
        this.preset,
        this.backgroundColor,
        this.showGrid,
        this.horizontalGrid,
        this.horizontalGridColor,
        this.horizontalGridThickness,
        this.horizontalGridStyle,
        this.verticalGrid,
        this.verticalGridColor,
        this.verticalGridThickness,
        this.verticalGridStyle,
        this.zebra,
        this.columnWidthMode,
        this.columnWidthPx,
        this.rowHeaderWidthPx,
        this.wrapText,
        this.freezeHeader,
        this.freezeRowHeaders
    ];
}

class HeaderCardSettings extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show header", value: true });
    columnHierarchy = new formattingSettings.ToggleSwitch({ name: "columnHierarchy", displayName: "Column hierarchy", value: true });
    backgroundColor = new formattingSettings.ColorPicker({ name: "backgroundColor", displayName: "Header background", value: { value: "" } });
    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Header font color", value: { value: "" } });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Header font size", value: 12 });
    fontFamily = new formattingSettings.ItemDropdown({ name: "fontFamily", displayName: "Header font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    textAlign = new formattingSettings.ItemDropdown({ name: "textAlign", displayName: "Header alignment", value: textAlignOptions[1], items: textAlignOptions });

    name: string = "header";
    displayName: string = "Header";
    slices: Array<FormattingSettingsSlice> = [this.show, this.columnHierarchy, this.backgroundColor, this.fontColor, this.fontSize, this.fontFamily, this.textAlign];
}

class RowsCardSettings extends FormattingSettingsCard {
    rowHeight = new formattingSettings.NumUpDown({ name: "rowHeight", displayName: "Row height (px)", value: 26 });
    hierarchyView = new formattingSettings.ToggleSwitch({ name: "hierarchyView", displayName: "Hierarchy view", value: false });
    labelPaddingLeft = new formattingSettings.NumUpDown({ name: "labelPaddingLeft", displayName: "Label left padding (px)", value: 8 });
    indentSize = new formattingSettings.NumUpDown({ name: "indentSize", displayName: "Indent size (px)", value: 16 });
    autoAggregateParents = new formattingSettings.ToggleSwitch({ name: "autoAggregateParents", displayName: "Auto aggregate parents", value: true });
    blankAsZero = new formattingSettings.ToggleSwitch({ name: "blankAsZero", displayName: "Treat blanks as zero", value: true });

    parentBold = new formattingSettings.ToggleSwitch({ name: "parentBold", displayName: "Parent bold", value: true });
    parentFontColor = new formattingSettings.ColorPicker({ name: "parentFontColor", displayName: "Parent font color", value: { value: "" } });
    parentBackground = new formattingSettings.ColorPicker({ name: "parentBackground", displayName: "Parent background", value: { value: "" } });
    parentFontSize = new formattingSettings.NumUpDown({ name: "parentFontSize", displayName: "Parent font size", value: 12 });
    parentFontFamily = new formattingSettings.ItemDropdown({ name: "parentFontFamily", displayName: "Parent font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    parentTextAlign = new formattingSettings.ItemDropdown({ name: "parentTextAlign", displayName: "Parent alignment", value: textAlignOptions[0], items: textAlignOptions });

    childFontColor = new formattingSettings.ColorPicker({ name: "childFontColor", displayName: "Child font color", value: { value: "" } });
    childBackground = new formattingSettings.ColorPicker({ name: "childBackground", displayName: "Child background", value: { value: "" } });
    childFontSize = new formattingSettings.NumUpDown({ name: "childFontSize", displayName: "Child font size", value: 12 });
    childFontFamily = new formattingSettings.ItemDropdown({ name: "childFontFamily", displayName: "Child font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    childTextAlign = new formattingSettings.ItemDropdown({ name: "childTextAlign", displayName: "Child alignment", value: textAlignOptions[0], items: textAlignOptions });

    // Totals controls moved to Totals card (keep legacy fields for backward compatibility; not shown in Rows pane)
    showGrandTotal = new formattingSettings.ToggleSwitch({ name: "showGrandTotal", displayName: "Row level total (Grand total) (legacy)", value: false });
    showSubtotals = new formattingSettings.ToggleSwitch({ name: "showSubtotals", displayName: "Sub row level total (Subtotals) (legacy)", value: false });
    showColumnTotal = new formattingSettings.ToggleSwitch({ name: "showColumnTotal", displayName: "Column level total (legacy)", value: false });
    grandTotalLabel = new formattingSettings.TextInput({ name: "grandTotalLabel", displayName: "Grand total label (legacy)", value: "Grand Total", placeholder: "Grand Total" });
    subtotalLabelTemplate = new formattingSettings.TextInput({ name: "subtotalLabelTemplate", displayName: "Subtotal label template (legacy)", value: "Total {label}", placeholder: "Total {label}" });
    columnTotalLabel = new formattingSettings.TextInput({ name: "columnTotalLabel", displayName: "Column total label (legacy)", value: "Column total", placeholder: "Column total" });

    // Legacy (kept for backward compatibility; not shown in pane)
    totalsBold = new formattingSettings.ToggleSwitch({ name: "totalsBold", displayName: "Totals bold (legacy)", value: true });
    totalsFontColor = new formattingSettings.ColorPicker({ name: "totalsFontColor", displayName: "Totals font color (legacy)", value: { value: "" } });
    totalsBackground = new formattingSettings.ColorPicker({ name: "totalsBackground", displayName: "Totals background (legacy)", value: { value: "" } });
    totalsFontFamily = new formattingSettings.ItemDropdown({ name: "totalsFontFamily", displayName: "Totals font (legacy)", value: fontFamilyOptions[0], items: fontFamilyOptions });
    totalsFontSize = new formattingSettings.NumUpDown({ name: "totalsFontSize", displayName: "Totals font size (legacy)", value: 12 });
    totalsLabelAlign = new formattingSettings.ItemDropdown({ name: "totalsLabelAlign", displayName: "Totals label alignment (legacy)", value: textAlignOptions[0], items: textAlignOptions });
    totalsValueAlign = new formattingSettings.ItemDropdown({ name: "totalsValueAlign", displayName: "Totals value alignment (legacy)", value: textAlignOptions[2], items: textAlignOptions });

    // Separate totals styling moved to Totals card (keep legacy fields for backward compatibility; not shown in Rows pane)
    grandTotalBold = new formattingSettings.ToggleSwitch({ name: "grandTotalBold", displayName: "Grand total bold (legacy)", value: true });
    grandTotalFontColor = new formattingSettings.ColorPicker({ name: "grandTotalFontColor", displayName: "Grand total font color", value: { value: "" } });
    grandTotalBackground = new formattingSettings.ColorPicker({ name: "grandTotalBackground", displayName: "Grand total background", value: { value: "" } });
    grandTotalFontFamily = new formattingSettings.ItemDropdown({ name: "grandTotalFontFamily", displayName: "Grand total font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    grandTotalFontSize = new formattingSettings.NumUpDown({ name: "grandTotalFontSize", displayName: "Grand total font size", value: 12 });
    grandTotalLabelAlign = new formattingSettings.ItemDropdown({ name: "grandTotalLabelAlign", displayName: "Grand total label alignment", value: textAlignOptions[0], items: textAlignOptions });
    grandTotalValueAlign = new formattingSettings.ItemDropdown({ name: "grandTotalValueAlign", displayName: "Grand total value alignment", value: textAlignOptions[2], items: textAlignOptions });

    subtotalBold = new formattingSettings.ToggleSwitch({ name: "subtotalBold", displayName: "Subtotal bold", value: true });
    subtotalFontColor = new formattingSettings.ColorPicker({ name: "subtotalFontColor", displayName: "Subtotal font color", value: { value: "" } });
    subtotalBackground = new formattingSettings.ColorPicker({ name: "subtotalBackground", displayName: "Subtotal background", value: { value: "" } });
    subtotalFontFamily = new formattingSettings.ItemDropdown({ name: "subtotalFontFamily", displayName: "Subtotal font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    subtotalFontSize = new formattingSettings.NumUpDown({ name: "subtotalFontSize", displayName: "Subtotal font size", value: 12 });
    subtotalLabelAlign = new formattingSettings.ItemDropdown({ name: "subtotalLabelAlign", displayName: "Subtotal label alignment", value: textAlignOptions[0], items: textAlignOptions });
    subtotalValueAlign = new formattingSettings.ItemDropdown({ name: "subtotalValueAlign", displayName: "Subtotal value alignment", value: textAlignOptions[2], items: textAlignOptions });

    columnTotalBold = new formattingSettings.ToggleSwitch({ name: "columnTotalBold", displayName: "Column total bold", value: true });
    columnTotalFontColor = new formattingSettings.ColorPicker({ name: "columnTotalFontColor", displayName: "Column total font color", value: { value: "" } });
    columnTotalBackground = new formattingSettings.ColorPicker({ name: "columnTotalBackground", displayName: "Column total background", value: { value: "" } });
    columnTotalFontFamily = new formattingSettings.ItemDropdown({ name: "columnTotalFontFamily", displayName: "Column total font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    columnTotalFontSize = new formattingSettings.NumUpDown({ name: "columnTotalFontSize", displayName: "Column total font size", value: 12 });
    columnTotalLabelAlign = new formattingSettings.ItemDropdown({ name: "columnTotalLabelAlign", displayName: "Column total label alignment", value: textAlignOptions[0], items: textAlignOptions });
    columnTotalValueAlign = new formattingSettings.ItemDropdown({ name: "columnTotalValueAlign", displayName: "Column total value alignment", value: textAlignOptions[2], items: textAlignOptions });

    name: string = "rows";
    displayName: string = "Rows";
    slices: Array<FormattingSettingsSlice> = [
        this.rowHeight,
        this.hierarchyView,
        this.labelPaddingLeft,
        this.indentSize,
        this.autoAggregateParents,
        this.blankAsZero,
        this.parentBold,
        this.parentFontColor,
        this.parentBackground,
        this.parentFontSize,
        this.parentFontFamily,
        this.parentTextAlign,
        this.childFontColor,
        this.childBackground,
        this.childFontSize,
        this.childFontFamily,
        this.childTextAlign
    ];
}

class TotalsCardSettings extends FormattingSettingsCard {
    showGrandTotal = new formattingSettings.ToggleSwitch({ name: "showGrandTotal", displayName: "Show grand total", value: false });
    showSubtotals = new formattingSettings.ToggleSwitch({ name: "showSubtotals", displayName: "Show subtotals", value: false });
    showColumnTotal = new formattingSettings.ToggleSwitch({ name: "showColumnTotal", displayName: "Show column total", value: false });

    grandTotalLabel = new formattingSettings.TextInput({ name: "grandTotalLabel", displayName: "Grand total label", value: "Grand Total", placeholder: "Grand Total" });
    subtotalLabelTemplate = new formattingSettings.TextInput({ name: "subtotalLabelTemplate", displayName: "Subtotal label template", value: "Total {label}", placeholder: "Total {label}" });
    columnTotalLabel = new formattingSettings.TextInput({ name: "columnTotalLabel", displayName: "Column total label", value: "Column total", placeholder: "Column total" });

    // Totals grids should work even if main grid is off.
    grandTotalGrid = new formattingSettings.ToggleSwitch({ name: "grandTotalGrid", displayName: "Grand total grid", value: false });
    grandTotalGridColor = new formattingSettings.ColorPicker({ name: "grandTotalGridColor", displayName: "Grand total grid color", value: { value: "" } });
    grandTotalGridThickness = new formattingSettings.NumUpDown({ name: "grandTotalGridThickness", displayName: "Grand total grid thickness", value: 2 });
    grandTotalGridStyle = new formattingSettings.ItemDropdown({ name: "grandTotalGridStyle", displayName: "Grand total grid style", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    subtotalGrid = new formattingSettings.ToggleSwitch({ name: "subtotalGrid", displayName: "Subtotal grid", value: false });
    subtotalGridColor = new formattingSettings.ColorPicker({ name: "subtotalGridColor", displayName: "Subtotal grid color", value: { value: "" } });
    subtotalGridThickness = new formattingSettings.NumUpDown({ name: "subtotalGridThickness", displayName: "Subtotal grid thickness", value: 1 });
    subtotalGridStyle = new formattingSettings.ItemDropdown({ name: "subtotalGridStyle", displayName: "Subtotal grid style", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    columnTotalGrid = new formattingSettings.ToggleSwitch({ name: "columnTotalGrid", displayName: "Column total grid", value: false });
    columnTotalGridColor = new formattingSettings.ColorPicker({ name: "columnTotalGridColor", displayName: "Column total grid color", value: { value: "" } });
    columnTotalGridThickness = new formattingSettings.NumUpDown({ name: "columnTotalGridThickness", displayName: "Column total grid thickness", value: 2 });
    columnTotalGridStyle = new formattingSettings.ItemDropdown({ name: "columnTotalGridStyle", displayName: "Column total grid style", value: gridLineStyleOptions[0], items: gridLineStyleOptions });

    // Totals typography/colors
    grandTotalBold = new formattingSettings.ToggleSwitch({ name: "grandTotalBold", displayName: "Grand total bold", value: true });
    grandTotalFontColor = new formattingSettings.ColorPicker({ name: "grandTotalFontColor", displayName: "Grand total font color", value: { value: "" } });
    grandTotalBackground = new formattingSettings.ColorPicker({ name: "grandTotalBackground", displayName: "Grand total background", value: { value: "" } });
    grandTotalFontFamily = new formattingSettings.ItemDropdown({ name: "grandTotalFontFamily", displayName: "Grand total font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    grandTotalFontSize = new formattingSettings.NumUpDown({ name: "grandTotalFontSize", displayName: "Grand total font size", value: 12 });
    grandTotalLabelAlign = new formattingSettings.ItemDropdown({ name: "grandTotalLabelAlign", displayName: "Grand total label alignment", value: textAlignOptions[0], items: textAlignOptions });
    grandTotalValueAlign = new formattingSettings.ItemDropdown({ name: "grandTotalValueAlign", displayName: "Grand total value alignment", value: textAlignOptions[2], items: textAlignOptions });

    subtotalBold = new formattingSettings.ToggleSwitch({ name: "subtotalBold", displayName: "Subtotal bold", value: true });
    subtotalFontColor = new formattingSettings.ColorPicker({ name: "subtotalFontColor", displayName: "Subtotal font color", value: { value: "" } });
    subtotalBackground = new formattingSettings.ColorPicker({ name: "subtotalBackground", displayName: "Subtotal background", value: { value: "" } });
    subtotalFontFamily = new formattingSettings.ItemDropdown({ name: "subtotalFontFamily", displayName: "Subtotal font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    subtotalFontSize = new formattingSettings.NumUpDown({ name: "subtotalFontSize", displayName: "Subtotal font size", value: 12 });
    subtotalLabelAlign = new formattingSettings.ItemDropdown({ name: "subtotalLabelAlign", displayName: "Subtotal label alignment", value: textAlignOptions[0], items: textAlignOptions });
    subtotalValueAlign = new formattingSettings.ItemDropdown({ name: "subtotalValueAlign", displayName: "Subtotal value alignment", value: textAlignOptions[2], items: textAlignOptions });

    columnTotalBold = new formattingSettings.ToggleSwitch({ name: "columnTotalBold", displayName: "Column total bold", value: true });
    columnTotalFontColor = new formattingSettings.ColorPicker({ name: "columnTotalFontColor", displayName: "Column total font color", value: { value: "" } });
    columnTotalBackground = new formattingSettings.ColorPicker({ name: "columnTotalBackground", displayName: "Column total background", value: { value: "" } });
    columnTotalFontFamily = new formattingSettings.ItemDropdown({ name: "columnTotalFontFamily", displayName: "Column total font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    columnTotalFontSize = new formattingSettings.NumUpDown({ name: "columnTotalFontSize", displayName: "Column total font size", value: 12 });
    columnTotalLabelAlign = new formattingSettings.ItemDropdown({ name: "columnTotalLabelAlign", displayName: "Column total label alignment", value: textAlignOptions[0], items: textAlignOptions });
    columnTotalValueAlign = new formattingSettings.ItemDropdown({ name: "columnTotalValueAlign", displayName: "Column total value alignment", value: textAlignOptions[2], items: textAlignOptions });

    name: string = "totals";
    displayName: string = "Totals";
    slices: Array<FormattingSettingsSlice> = [
        this.showGrandTotal,
        this.showSubtotals,
        this.showColumnTotal,
        this.grandTotalLabel,
        this.subtotalLabelTemplate,
        this.columnTotalLabel,
        this.grandTotalGrid,
        this.grandTotalGridColor,
        this.grandTotalGridThickness,
        this.grandTotalGridStyle,
        this.subtotalGrid,
        this.subtotalGridColor,
        this.subtotalGridThickness,
        this.subtotalGridStyle,
        this.columnTotalGrid,
        this.columnTotalGridColor,
        this.columnTotalGridThickness,
        this.columnTotalGridStyle,
        this.grandTotalBold,
        this.grandTotalFontColor,
        this.grandTotalBackground,
        this.grandTotalFontFamily,
        this.grandTotalFontSize,
        this.grandTotalLabelAlign,
        this.grandTotalValueAlign,
        this.subtotalBold,
        this.subtotalFontColor,
        this.subtotalBackground,
        this.subtotalFontFamily,
        this.subtotalFontSize,
        this.subtotalLabelAlign,
        this.subtotalValueAlign,
        this.columnTotalBold,
        this.columnTotalFontColor,
        this.columnTotalBackground,
        this.columnTotalFontFamily,
        this.columnTotalFontSize,
        this.columnTotalLabelAlign,
        this.columnTotalValueAlign
    ];
}

class NumbersCardSettings extends FormattingSettingsCard {
    displayUnits = new formattingSettings.ItemDropdown({
        name: "displayUnits",
        displayName: "Display units",
        value: displayUnitOptions[0],
        items: displayUnitOptions
    });

    decimals = new formattingSettings.NumUpDown({ name: "decimals", displayName: "Decimal places", value: 0 });

    valueFontFamily = new formattingSettings.ItemDropdown({ name: "valueFontFamily", displayName: "Value font", value: fontFamilyOptions[0], items: fontFamilyOptions });
    valueFontSize = new formattingSettings.NumUpDown({ name: "valueFontSize", displayName: "Value font size", value: 12 });
    valueTextAlign = new formattingSettings.ItemDropdown({ name: "valueTextAlign", displayName: "Value alignment", value: textAlignOptions[2], items: textAlignOptions });
    valueFontColor = new formattingSettings.ColorPicker({ name: "valueFontColor", displayName: "Value font color", value: { value: "" } });
    valueBold = new formattingSettings.ToggleSwitch({ name: "valueBold", displayName: "Value bold", value: false });

    name: string = "numbers";
    displayName: string = "Values";
    slices: Array<FormattingSettingsSlice> = [this.displayUnits, this.decimals, this.valueFontFamily, this.valueFontSize, this.valueTextAlign, this.valueFontColor, this.valueBold];
}

class LayoutCardSettings extends FormattingSettingsCard {
    layoutJson = new formattingSettings.TextInput({
        name: "layoutJson",
        displayName: "Layout JSON",
        value: "",
        placeholder: "[{ \"code\": \"BS\", \"label\": \"Balance Sheet\", \"type\": \"blank\" }, { \"code\": \"ASSETS\", \"label\": \"Assets\" }]"
    });

    formulasJson = new formattingSettings.TextInput({
        name: "formulasJson",
        displayName: "Formulas JSON",
        value: "",
        placeholder: "{ \"GROSS_PROFIT\": \"VALUE(\\\"REV\\\") - VALUE(\\\"COGS\\\")\" }"
    });

    name: string = "layout";
    displayName: string = "Layout";
    slices: Array<FormattingSettingsSlice> = [this.layoutJson, this.formulasJson];
}

class ConditionalFormattingCardSettings extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({ name: "enabled", displayName: "Enable", value: false });
    showEditor = new formattingSettings.ToggleSwitch({ name: "showEditor", displayName: "Open editor", value: false });

    name: string = "conditionalFormatting";
    displayName: string = "Conditional Formatting";
    slices: Array<FormattingSettingsSlice> = [this.enabled, this.showEditor];
}

class CustomTableCardSettings extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({ name: "enabled", displayName: "Custom table", value: false });
    showEditor = new formattingSettings.ToggleSwitch({ name: "showEditor", displayName: "Open table editor", value: false });

    name: string = "customTable";
    displayName: string = "Custom Table";
    slices: Array<FormattingSettingsSlice> = [this.enabled, this.showEditor];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    table = new TableCardSettings();
    header = new HeaderCardSettings();
    rows = new RowsCardSettings();
    totals = new TotalsCardSettings();
    numbers = new NumbersCardSettings();
    layout = new LayoutCardSettings();
    conditionalFormatting = new ConditionalFormattingCardSettings();
    customTable = new CustomTableCardSettings();

    cards = [this.table, this.header, this.rows, this.totals, this.numbers, this.layout, this.conditionalFormatting, this.customTable];
}
