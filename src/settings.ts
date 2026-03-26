import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { defaultNodePaletteMode } from "./palettes";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;
import Slice = formattingSettings.Slice;
import ToggleSwitch = formattingSettings.ToggleSwitch;
import NumUpDown = formattingSettings.NumUpDown;
import ColorPicker = formattingSettings.ColorPicker;
import ReadOnlyText = formattingSettings.ReadOnlyText;
import AutoDropdown = formattingSettings.AutoDropdown;

const nodeWidthOptions: powerbi.visuals.NumUpDownFormat = {
    minValue: {
        type: powerbi.visuals.ValidatorType.Min,
        value: 8,
    },
    maxValue: {
        type: powerbi.visuals.ValidatorType.Max,
        value: 48,
    },
};

const nodePaddingOptions: powerbi.visuals.NumUpDownFormat = {
    minValue: {
        type: powerbi.visuals.ValidatorType.Min,
        value: 4,
    },
    maxValue: {
        type: powerbi.visuals.ValidatorType.Max,
        value: 80,
    },
};

const fontSizeOptions: powerbi.visuals.NumUpDownFormat = {
    minValue: {
        type: powerbi.visuals.ValidatorType.Min,
        value: 8,
    },
    maxValue: {
        type: powerbi.visuals.ValidatorType.Max,
        value: 28,
    },
};

const decimalPlacesOptions: powerbi.visuals.NumUpDownFormat = {
    minValue: {
        type: powerbi.visuals.ValidatorType.Min,
        value: 0,
    },
    maxValue: {
        type: powerbi.visuals.ValidatorType.Max,
        value: 6,
    },
};

class LayoutCardSettings extends Card {
    nodeWidth = new NumUpDown({
        name: "nodeWidth",
        displayName: "Node width",
        value: 18,
        options: nodeWidthOptions,
    });

    nodePadding = new NumUpDown({
        name: "nodePadding",
        displayName: "Node spacing",
        value: 18,
        options: nodePaddingOptions,
    });

    name: string = "layout";
    displayName: string = "Layout";
    slices: Slice[] = [this.nodeWidth, this.nodePadding];
}

class LabelsCardSettings extends Card {
    show = new ToggleSwitch({
        name: "show",
        displayName: undefined,
        value: true,
    });

    showText = new ToggleSwitch({
        name: "showText",
        displayName: "Show text",
        value: true,
    });

    showValue = new ToggleSwitch({
        name: "showValue",
        displayName: "Show value",
        value: false,
    });

    fontSize = new NumUpDown({
        name: "fontSize",
        displayName: "Font size",
        value: 12,
        options: fontSizeOptions,
    });

    fill = new ColorPicker({
        name: "fill",
        displayName: "Color",
        value: { value: "#1f2933" },
    });

    topLevelSlice: ToggleSwitch = this.show;
    name: string = "labels";
    displayName: string = "Node Labels";
    slices: Slice[] = [this.showText, this.showValue, this.fontSize, this.fill];
}

class LinkLabelsCardSettings extends Card {
    show = new ToggleSwitch({
        name: "show",
        displayName: undefined,
        value: false,
    });

    showText = new ToggleSwitch({
        name: "showText",
        displayName: "Show text",
        value: true,
    });

    showValue = new ToggleSwitch({
        name: "showValue",
        displayName: "Show value",
        value: true,
    });

    fontSize = new NumUpDown({
        name: "fontSize",
        displayName: "Font size",
        value: 10,
        options: fontSizeOptions,
    });

    fill = new ColorPicker({
        name: "fill",
        displayName: "Color",
        value: { value: "#475569" },
    });

    topLevelSlice: ToggleSwitch = this.show;
    name: string = "linkLabels";
    displayName: string = "Edge Labels";
    slices: Slice[] = [this.showText, this.showValue, this.fontSize, this.fill];
}

class ValueFormattingCardSettings extends Card {
    showAsPercentage = new ToggleSwitch({
        name: "showAsPercentage",
        displayName: "Show as percentage",
        value: false,
    });

    decimalPlaces = new NumUpDown({
        name: "decimalPlaces",
        displayName: "Decimal places",
        value: 2,
        options: decimalPlacesOptions,
    });

    name: string = "valueFormatting";
    displayName: string = "Value Format";
    slices: Slice[] = [this.showAsPercentage, this.decimalPlaces];
}

class NodeFormatCardSettings extends Card {
    targetName = new ReadOnlyText({
        name: "targetName",
        displayName: "Editing",
        value: "All nodes",
    });

    useThemePalette = new ToggleSwitch({
        name: "useThemePalette",
        displayName: "Use report palette",
        value: true,
    });

    fill = new ColorPicker({
        name: "fill",
        displayName: "Color",
        value: { value: "#4e79a7" },
    });

    paletteMode = new AutoDropdown({
        name: "paletteMode",
        displayName: "Palette",
        value: defaultNodePaletteMode,
        visible: false,
    });

    name: string = "defaultNodes";
    displayName: string = "Node Format";
    slices: Slice[] = [this.targetName, this.useThemePalette, this.paletteMode, this.fill];

    sync(
        targetName: string,
        style: { color: string; opacity: number },
        useThemePalette: boolean,
        showStyleControls: boolean,
        showFillControl: boolean,
    ): void {
        this.targetName.value = targetName;
        this.fill.value = { value: style.color };
        this.paletteMode.visible = !useThemePalette;
        this.fill.visible = showStyleControls && showFillControl;
    }
}

class PersistedStateCardSettings extends Card {
    selectedNodeName = new ReadOnlyText({
        name: "selectedNodeName",
        displayName: "Selected node name",
        value: "",
        visible: false,
    });

    selectedLinkName = new ReadOnlyText({
        name: "selectedLinkName",
        displayName: "Selected link name",
        value: "",
        visible: false,
    });

    nodeStyleMap = new ReadOnlyText({
        name: "nodeStyleMap",
        displayName: "Node style map",
        value: "{}",
        visible: false,
    });

    linkStyleMap = new ReadOnlyText({
        name: "linkStyleMap",
        displayName: "Link style map",
        value: "{}",
        visible: false,
    });

    nodePositionMap = new ReadOnlyText({
        name: "nodePositionMap",
        displayName: "Node position map",
        value: "{}",
        visible: false,
    });

    nodeGlobalStyle = new ReadOnlyText({
        name: "nodeGlobalStyle",
        displayName: "Node global style",
        value: "{\"color\":\"#4e79a7\",\"opacity\":90}",
        visible: false,
    });

    nodeFormatSyncState = new ReadOnlyText({
        name: "nodeFormatSyncState",
        displayName: "Node format sync state",
        value: "",
        visible: false,
    });

    linkFormatSyncState = new ReadOnlyText({
        name: "linkFormatSyncState",
        displayName: "Link format sync state",
        value: "",
        visible: false,
    });

    name: string = "persistedState";
    displayName: string = "Persisted State";
    visible: boolean = false;
    slices: Slice[] = [
        this.selectedNodeName,
        this.selectedLinkName,
        this.nodeStyleMap,
        this.linkStyleMap,
        this.nodePositionMap,
        this.nodeGlobalStyle,
        this.nodeFormatSyncState,
        this.linkFormatSyncState,
    ];
}

export class VisualFormattingSettingsModel extends Model {
    layout = new LayoutCardSettings();
    labels = new LabelsCardSettings();
    linkLabels = new LinkLabelsCardSettings();
    valueFormatting = new ValueFormattingCardSettings();
    defaultNodes = new NodeFormatCardSettings();
    persistedState = new PersistedStateCardSettings();

    cards: Card[] = [
        this.layout,
        this.labels,
        this.linkLabels,
        this.valueFormatting,
        this.defaultNodes,
        this.persistedState,
    ];

    prepareFormattingState(options: {
        nodeTargetName: string;
        nodeTargetStyle: { color: string; opacity: number };
        useThemePalette: boolean;
        showNodeStyleControls: boolean;
        showNodeFillControl: boolean;
    }): void {
        this.defaultNodes.sync(
            options.nodeTargetName,
            options.nodeTargetStyle,
            options.useThemePalette,
            options.showNodeStyleControls,
            options.showNodeFillControl,
        );
    }
}
