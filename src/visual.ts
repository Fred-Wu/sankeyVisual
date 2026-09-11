/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";

import powerbi from "powerbi-visuals-api";
import { sankey, sankeyLinkHorizontal } from "d3-sankey";
import type { SankeyGraph, SankeyLayout } from "d3-sankey";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import "./../style/visual.less";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import IViewport = powerbi.IViewport;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;

import {
    AppliedStyleContext,
    buildSankeyData,
    PersistedStyle,
    SankeyVisualData,
    StyleMap,
    parseStyleMap,
    serializeStyleMap,
} from "./data";
import { defaultNodePaletteMode, getNodePaletteColors, isNodePaletteMode, NodePaletteMode } from "./palettes";
import { VisualFormattingSettingsModel } from "./settings";
import { alignByInputOrder } from "./layout";

type VisualNode = {
    id: string;
    label: string;
    layerIndex: number;
    color: string;
    opacity: number;
    incomingValue: number;
    outgoingValue: number;
    selectionIds: ISelectionId[];
    contextMenuSelectionId?: ISelectionId;
    value?: number;
    x0?: number;
    x1?: number;
    y0?: number;
    y1?: number;
};

type VisualLink = {
    id: string;
    key: string;
    label: string;
    source: string | VisualNode;
    target: string | VisualNode;
    value: number;
    color: string;
    opacity: number;
    selectionIds: ISelectionId[];
    contextMenuSelectionId?: ISelectionId;
    width?: number;
    y0?: number;
    y1?: number;
};

const svgNamespace = "http://www.w3.org/2000/svg";
const layerSuffixPattern = /\s+\(L\d+\)$/;
const dragStartThreshold = 1;
const nodePositionPrecision = 10000;

type DragBounds = {
    left: number;
    top: number;
    right: number;
    bottom: number;
};

type Point = {
    x: number;
    y: number;
};

type NodePosition = {
    x?: number;
    y?: number;
};

type ZoomState = {
    scale: number;
    translateX: number;
    translateY: number;
};

type SyncedSelectionState = {
    key: string;
    controlStyle: PersistedStyle;
    targetStyle: PersistedStyle;
};

type FormatTargetState = {
    key: string;
    label: string;
    style: PersistedStyle;
    showStyleControls: boolean;
    showFillControl: boolean;
};

type NodePositionMap = Record<string, NodePosition>;
type GraphNode = SankeyGraph<VisualNode, VisualLink>["nodes"][number];
type GraphLink = SankeyGraph<VisualNode, VisualLink>["links"][number];
type SankeyGenerator = SankeyLayout<SankeyGraph<VisualNode, VisualLink>, VisualNode, VisualLink>;
type RgbColor = { r: number; g: number; b: number };
type TooltipItem = {
    displayName: string;
    value: string;
    color?: string;
    header?: string;
    opacity?: string;
};

const allNodesFormatKey = "__all_nodes__";
const noNodeSelectionFormatKey = "__no_node_selection__";
const defaultLinkOpacity = 45;
const minZoomScale = 1;
const maxZoomScale = 5;
const wheelZoomSensitivity = 0.0015;
const defaultNodeStyle: PersistedStyle = {
    color: "#4e79a7",
    opacity: 90,
};

function createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T): SVGElementTagNameMap[T] {
    return document.createElementNS(svgNamespace, tagName);
}

function setAttributes(element: Element, attributes: Record<string, string | number>): void {
    Object.keys(attributes).forEach((key: string) => {
        element.setAttribute(key, String(attributes[key]));
    });
}

function clamp(value: number, minValue: number, maxValue: number): number {
    return Math.max(minValue, Math.min(maxValue, value));
}

function isFiniteNumber(value: number | undefined): value is number {
    return typeof value === "number" && Number.isFinite(value);
}

function parseColor(color: string): RgbColor | undefined {
    const hexMatch = color.trim().match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (hexMatch) {
        const hexValue = hexMatch[1];
        const expandedHex = hexValue.length === 3
            ? hexValue.split("").map((value: string) => `${value}${value}`).join("")
            : hexValue;

        return {
            r: Number.parseInt(expandedHex.slice(0, 2), 16),
            g: Number.parseInt(expandedHex.slice(2, 4), 16),
            b: Number.parseInt(expandedHex.slice(4, 6), 16),
        };
    }

    const rgbMatch = color.trim().match(/^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)(?:\s*,\s*[\d.]+\s*)?\)$/i);
    if (rgbMatch) {
        return {
            r: clamp(Math.round(Number(rgbMatch[1])), 0, 255),
            g: clamp(Math.round(Number(rgbMatch[2])), 0, 255),
            b: clamp(Math.round(Number(rgbMatch[3])), 0, 255),
        };
    }

    return undefined;
}

function toRgbString(color: RgbColor): string {
    return `rgb(${color.r}, ${color.g}, ${color.b})`;
}

function mixWithWhite(color: string, intensity: number): string {
    const parsedColor = parseColor(color);
    if (!parsedColor) {
        return color;
    }

    const mixRatio = clamp(intensity, 0, 1);
    return toRgbString({
        r: Math.round(parsedColor.r + ((255 - parsedColor.r) * mixRatio)),
        g: Math.round(parsedColor.g + ((255 - parsedColor.g) * mixRatio)),
        b: Math.round(parsedColor.b + ((255 - parsedColor.b) * mixRatio)),
    });
}

function clampNodeToBounds(node: GraphNode, bounds: DragBounds): void {
    const nodeWidth = Math.max(1, (node.x1 || bounds.left) - (node.x0 || bounds.left));
    const nodeHeight = Math.max(1, (node.y1 || bounds.top) - (node.y0 || bounds.top));
    const nextX0 = clamp(node.x0 || bounds.left, bounds.left, bounds.right - nodeWidth);
    const nextY0 = clamp(node.y0 || bounds.top, bounds.top, bounds.bottom - nodeHeight);

    node.x0 = nextX0;
    node.x1 = nextX0 + nodeWidth;
    node.y0 = nextY0;
    node.y1 = nextY0 + nodeHeight;
}

export class Visual implements IVisual {
    private readonly host: IVisualHost;
    private readonly allowInteractions: boolean;
    private readonly events: IVisualEventService;
    private readonly selectionManager: ISelectionManager;
    private readonly root: HTMLDivElement;
    private readonly surface: HTMLDivElement;
    private readonly svg: SVGSVGElement;
    private readonly overlay: HTMLDivElement;
    private readonly landingPage: HTMLDivElement;
    private readonly warning: HTMLDivElement;
    private readonly formattingSettingsService: FormattingSettingsService;

    private formattingSettings: VisualFormattingSettingsModel;
    private modelValueFormatter: valueFormatter.IValueFormatter;
    private percentageValueFormatter: valueFormatter.IValueFormatter;
    private suppressClickSelection: boolean = false;
    private hoveredLinkKey?: string;
    private focusedNodeId?: string;
    private pendingSelectedNodeName?: string;
    private pendingClearSelection: boolean = false;
    private currentViewport?: IViewport;
    private currentGraphViewport?: SVGGElement;
    private zoomState: ZoomState = {
        scale: minZoomScale,
        translateX: 0,
        translateY: 0,
    };

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.allowInteractions = this.host.hostCapabilities?.allowInteractions !== false;
        this.events = options.host.eventService;
        this.selectionManager = this.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.modelValueFormatter = valueFormatter.create({
            cultureSelector: this.host.locale,
        });
        this.percentageValueFormatter = valueFormatter.create({
            format: "0%",
            precision: 2,
            cultureSelector: this.host.locale,
        });

        this.root = document.createElement("div");
        this.root.className = "sankey-root";
        this.root.classList.toggle("is-interactive", this.allowInteractions);

        this.surface = document.createElement("div");
        this.surface.className = "sankey-surface";

        this.svg = createSvgElement("svg");
        this.svg.classList.add("sankey-svg");

        this.overlay = document.createElement("div");
        this.overlay.className = "sankey-overlay";

        this.landingPage = document.createElement("div");
        this.landingPage.className = "sankey-landing";

        const landingTitle = document.createElement("div");
        landingTitle.className = "sankey-landing-title";
        landingTitle.textContent = "Build a Sankey from source and target pairs";

        const landingBody = document.createElement("div");
        landingBody.className = "sankey-landing-body";
        landingBody.textContent = "Add one or more Source->Target Pair(s) fields and a single Values measure.";

        const landingHint = document.createElement("div");
        landingHint.className = "sankey-landing-hint";
        landingHint.textContent = "Examples: A -> B, A -> B, C, or A -> B; B -> D";

        this.landingPage.appendChild(landingTitle);
        this.landingPage.appendChild(landingBody);
        this.landingPage.appendChild(landingHint);

        this.warning = document.createElement("div");
        this.warning.className = "sankey-warning";

        this.overlay.appendChild(this.landingPage);
        this.overlay.appendChild(this.warning);
        this.svg.addEventListener("click", () => {
            if (!this.allowInteractions) {
                return;
            }

            if (this.suppressClickSelection) {
                this.suppressClickSelection = false;
                return;
            }

            this.clearSelection();
        });
        this.svg.addEventListener("contextmenu", (event: MouseEvent) => {
            event.preventDefault();
            this.showContextMenu(undefined, event);
        });
        this.svg.addEventListener("wheel", (event: WheelEvent) => {
            this.handleWheelZoom(event);
        }, { passive: false });
        this.svg.addEventListener("pointerdown", (event: PointerEvent) => {
            if (event.target !== this.svg) {
                return;
            }

            this.startViewportPan(event);
        });

        this.surface.appendChild(this.svg);
        this.surface.appendChild(this.overlay);
        this.root.appendChild(this.surface);
        options.element.appendChild(this.root);
    }

    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);

        try {
            const dataView = options.dataViews?.[0];

            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                dataView,
            );
            this.configureValueFormatters(dataView);

            const nodeStyleMap = parseStyleMap(this.formattingSettings.persistedState.nodeStyleMap.value);
            const linkStyleMap = parseStyleMap(this.formattingSettings.persistedState.linkStyleMap.value);
            const nodePositionMap = this.parseNodePositionMap(this.formattingSettings.persistedState.nodePositionMap.value);
            const nodeGlobalStyle = this.parsePersistedStyle(
                this.formattingSettings.persistedState.nodeGlobalStyle.value,
                defaultNodeStyle,
            );
            const nodeFormatSyncState = this.parseSelectionSyncState(this.formattingSettings.persistedState.nodeFormatSyncState.value);
            const linkFormatSyncState = this.parseSelectionSyncState(this.formattingSettings.persistedState.linkFormatSyncState.value);
            const persistedSelectedNodeName = this.formattingSettings.persistedState.selectedNodeName.value.trim();
            const selectedLinkLabel = this.formattingSettings.persistedState.selectedLinkName.value.trim();
            const hasAssignedFields = Boolean(dataView?.metadata?.columns?.length);

            const styleContext = this.buildAppliedStyleContext(nodeStyleMap, nodeGlobalStyle);
            const visualData = buildSankeyData(
                dataView,
                this.host.colorPalette,
                styleContext,
                this.host.createSelectionIdBuilder.bind(this.host),
                this.host.locale,
            );

            const activeNodeTargetName = this.getActiveNodeTargetName(
                visualData,
                persistedSelectedNodeName,
            );
            const nodeTarget = this.getNodeFormatTarget(
                visualData,
                nodeGlobalStyle,
                activeNodeTargetName,
            );
            const configuredNodeFormatStyle = this.getConfiguredNodeFormatStyle(nodeTarget.style);
            const displayedNodeFormatStyle = this.getDisplayedFormatStyle(
                nodeTarget,
                configuredNodeFormatStyle,
                nodeFormatSyncState,
            );

            this.syncPersistedStyles(
                nodeStyleMap,
                linkStyleMap,
                nodeGlobalStyle,
                nodeFormatSyncState,
                linkFormatSyncState,
                nodeTarget,
                configuredNodeFormatStyle,
                selectedLinkLabel,
            );

            this.formattingSettings.prepareFormattingState({
                nodeTargetName: nodeTarget.label,
                nodeTargetStyle: displayedNodeFormatStyle,
                useThemePalette: this.formattingSettings.defaultNodes.useThemePalette.value,
                showNodeStyleControls: nodeTarget.showStyleControls,
                showNodeFillControl: nodeTarget.showFillControl,
            });

            this.render(
                options.viewport,
                visualData,
                nodePositionMap,
                hasAssignedFields,
                activeNodeTargetName,
            );

            this.events.renderingFinished(options);
        } catch (error) {
            this.host.eventService.renderingFailed(options, String(error));
            throw error;
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        const formattingModel = this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
        this.addPersistedStateResetDescriptors(formattingModel);
        return formattingModel;
    }

    private buildAppliedStyleContext(
        nodeStyleMap: StyleMap,
        nodeGlobalStyle: PersistedStyle,
    ): AppliedStyleContext {
        const colorPalette = this.host.colorPalette;

        return {
            defaultNodeColor: nodeGlobalStyle.color,
            defaultNodeOpacity: nodeGlobalStyle.opacity,
            useThemePalette: this.formattingSettings.defaultNodes.useThemePalette.value,
            nodePaletteColors: getNodePaletteColors(this.getSelectedPaletteMode()),
            defaultLinkColor: nodeGlobalStyle.color,
            defaultLinkOpacity,
            inheritLinkColor: true,
            nodeStyles: nodeStyleMap,
            linkStyles: {},
            isHighContrast: colorPalette.isHighContrast,
            highContrastForeground: colorPalette.foreground.value,
            highContrastBackground: colorPalette.background.value,
        };
    }

    private configureValueFormatters(dataView: powerbi.DataView | undefined): void {
        const valueColumn = dataView?.metadata?.columns?.find(
            (column: powerbi.DataViewMetadataColumn) => column.roles?.values,
        );
        const decimalPlaces = Math.max(
            0,
            Math.min(6, Math.round(this.formattingSettings.valueFormatting.decimalPlaces.value)),
        );

        this.modelValueFormatter = valueFormatter.create({
            format: valueColumn?.format,
            columnType: valueColumn?.type,
            cultureSelector: this.host.locale,
        });
        this.percentageValueFormatter = valueFormatter.create({
            format: "0%",
            precision: decimalPlaces,
            cultureSelector: this.host.locale,
        });
    }

    private syncPersistedStyles(
        currentNodeStyles: StyleMap,
        currentLinkStyles: StyleMap,
        nodeGlobalStyle: PersistedStyle,
        nodeFormatSyncState: SyncedSelectionState | undefined,
        linkFormatSyncState: SyncedSelectionState | undefined,
        nodeTarget: FormatTargetState,
        configuredNodeFormatStyle: PersistedStyle,
        selectedLinkLabel: string,
    ): void {
        const merge: powerbi.VisualObjectInstance[] = [];

        this.syncNodeFormatState(
            merge,
            currentNodeStyles,
            nodeGlobalStyle,
            nodeFormatSyncState,
            nodeTarget,
            configuredNodeFormatStyle,
        );
        this.cleanupLegacyLinkState(
            merge,
            currentLinkStyles,
            selectedLinkLabel,
            linkFormatSyncState,
        );

        if (merge.length > 0) {
            this.host.persistProperties({ merge });
        }
    }

    private syncNodeFormatState(
        merge: powerbi.VisualObjectInstance[],
        currentNodeStyles: StyleMap,
        nodeGlobalStyle: PersistedStyle,
        syncState: SyncedSelectionState | undefined,
        target: FormatTargetState,
        configuredStyle: PersistedStyle,
    ): void {
        if (!target.showStyleControls) {
            if (
                !syncState
                || syncState.key !== target.key
                || !this.stylesMatch(syncState.controlStyle, configuredStyle)
                || !this.stylesMatch(syncState.targetStyle, target.style)
            ) {
                this.queueNodeFormatSyncState(merge, target.key, configuredStyle, target.style);
            }
            return;
        }

        if (!syncState || syncState.key !== target.key) {
            this.queueVisibleNodeFormatStyle(merge, configuredStyle, target.style);
            this.queueNodeFormatSyncState(merge, target.key, configuredStyle, target.style);
            return;
        }

        const controlUnchanged = this.stylesMatch(configuredStyle, syncState.controlStyle);
        const hostCaughtUpToTarget = this.stylesMatch(configuredStyle, syncState.targetStyle);

        if (controlUnchanged) {
            if (!this.stylesMatch(syncState.targetStyle, target.style)) {
                this.queueVisibleNodeFormatStyle(merge, configuredStyle, target.style);
                this.queueNodeFormatSyncState(merge, target.key, configuredStyle, target.style);
            }
            return;
        }

        if (hostCaughtUpToTarget) {
            this.queueNodeFormatSyncState(merge, target.key, configuredStyle, target.style);
            return;
        }

        if (target.key === allNodesFormatKey) {
            if (!this.stylesMatch(nodeGlobalStyle, configuredStyle)) {
                merge.push({
                    objectName: "persistedState",
                    selector: undefined,
                    properties: {
                        nodeGlobalStyle: this.serializePersistedStyle(configuredStyle),
                    },
                });
            }

            this.queueVisibleNodeFormatStyle(merge, configuredStyle, configuredStyle);
            this.queueNodeFormatSyncState(merge, target.key, configuredStyle, configuredStyle);
            return;
        }

        const existingStyle = currentNodeStyles[target.key];
        if (!existingStyle || !this.stylesMatch(existingStyle, configuredStyle)) {
            merge.push({
                objectName: "persistedState",
                selector: undefined,
                properties: {
                    nodeStyleMap: serializeStyleMap({
                        ...currentNodeStyles,
                        [target.key]: configuredStyle,
                    }),
                },
            });
        }

        this.queueVisibleNodeFormatStyle(merge, configuredStyle, configuredStyle);
        this.queueNodeFormatSyncState(merge, target.key, configuredStyle, configuredStyle);
    }

    private cleanupLegacyLinkState(
        merge: powerbi.VisualObjectInstance[],
        currentLinkStyles: StyleMap,
        selectedLinkLabel: string,
        linkFormatSyncState: SyncedSelectionState | undefined,
    ): void {
        const properties: Record<string, string> = {};

        if (selectedLinkLabel) {
            properties.selectedLinkName = "";
        }

        if (Object.keys(currentLinkStyles).length > 0) {
            properties.linkStyleMap = "{}";
        }

        if (linkFormatSyncState) {
            properties.linkFormatSyncState = "";
        }

        if (Object.keys(properties).length === 0) {
            return;
        }
        merge.push({
            objectName: "persistedState",
            selector: undefined,
            properties: {
                ...properties,
            },
        });
    }

    private queueNodeFormatSyncState(
        merge: powerbi.VisualObjectInstance[],
        key: string,
        controlStyle: PersistedStyle,
        targetStyle: PersistedStyle,
    ): void {
        merge.push({
            objectName: "persistedState",
            selector: undefined,
            properties: {
                nodeFormatSyncState: this.serializeSelectionSyncState({ key, controlStyle, targetStyle }),
            },
        });
    }

    private queueVisibleNodeFormatStyle(
        merge: powerbi.VisualObjectInstance[],
        currentStyle: PersistedStyle,
        nextStyle: PersistedStyle,
    ): void {
        if (this.stylesMatch(currentStyle, nextStyle)) {
            return;
        }

        merge.push({
            objectName: "defaultNodes",
            selector: undefined,
            properties: {
                fill: {
                    solid: {
                        color: nextStyle.color,
                    },
                },
            },
        });
    }

    private render(
        viewport: IViewport,
        visualData: SankeyVisualData,
        nodePositionMap: NodePositionMap,
        hasAssignedFields: boolean,
        selectedNodeName?: string,
    ): void {
        this.currentViewport = viewport;
        this.surface.style.width = `${viewport.width}px`;
        this.surface.style.height = `${viewport.height}px`;
        this.hoveredLinkKey = undefined;
        this.focusedNodeId = undefined;
        this.hideTooltip();
        this.hideLandingPage();
        this.svg.replaceChildren();
        this.currentGraphViewport = undefined;

        setAttributes(this.svg, {
            width: viewport.width,
            height: viewport.height,
            viewBox: `0 0 ${viewport.width} ${viewport.height}`,
            role: "img",
            "aria-label": "Sankey diagram",
        });
        this.normalizeZoomState(viewport);
        this.applyGraphViewportTransform();

        if (viewport.width < 180 || viewport.height < 140) {
            this.renderMessage("Resize the visual to render the Sankey diagram.");
            return;
        }

        if (!hasAssignedFields) {
            this.showLandingPage();
            return;
        }

        if (visualData.nodes.length === 0 || visualData.links.length === 0) {
            this.renderMessage("No valid Sankey flows were found. Use Source->Target Pair(s) values like A -> B and positive numeric Values.");
            return;
        }

        this.renderWarnings(visualData.warnings);

        const margin = {
            top: 20,
            right: 80,
            bottom: 32,
            left: 80,
        };

        const layoutWidth = Math.max(1, viewport.width - margin.left - margin.right);
        const layoutHeight = Math.max(1, viewport.height - margin.top - margin.bottom);
        const layoutTop = margin.top;
        const layoutBottom = margin.top + layoutHeight;
        const dragBounds: DragBounds = {
            left: margin.left,
            top: layoutTop,
            right: margin.left + layoutWidth,
            bottom: layoutBottom,
        };

        const sankeyLayout: SankeyGenerator = sankey<VisualNode, VisualLink>()
            .nodeId((node: VisualNode) => node.id)
            .nodeAlign(alignByInputOrder)
            .nodeSort(null)
            .linkSort(null)
            .nodeWidth(this.formattingSettings.layout.nodeWidth.value)
            .nodePadding(this.formattingSettings.layout.nodePadding.value)
            .extent([
                [dragBounds.left, dragBounds.top],
                [dragBounds.right, dragBounds.bottom],
            ]);

        const sankeyGraph = sankeyLayout({
            nodes: visualData.nodes.map((node: VisualNode) => ({ ...node })),
            links: visualData.links.map((link: VisualLink) => ({ ...link })),
        });

        this.applyPersistedNodePositions(sankeyGraph, nodePositionMap, dragBounds);
        sankeyLayout.update(sankeyGraph);

        const graphViewport = createSvgElement("g");
        graphViewport.classList.add("sankey-graph-viewport");
        this.svg.appendChild(graphViewport);
        this.currentGraphViewport = graphViewport;
        this.applyGraphViewportTransform();

        const linkGroup = createSvgElement("g");
        linkGroup.classList.add("sankey-links");
        graphViewport.appendChild(linkGroup);

        const linkLabelGroup = createSvgElement("g");
        linkLabelGroup.classList.add("sankey-link-labels");
        graphViewport.appendChild(linkLabelGroup);

        const selectionGroup = createSvgElement("g");
        selectionGroup.classList.add("sankey-node-selection");
        graphViewport.appendChild(selectionGroup);

        const nodeGroup = createSvgElement("g");
        nodeGroup.classList.add("sankey-nodes");
        graphViewport.appendChild(nodeGroup);

        const labelGroup = createSvgElement("g");
        labelGroup.classList.add("sankey-labels");
        graphViewport.appendChild(labelGroup);

        const linkPathGenerator = sankeyLinkHorizontal<VisualNode, VisualLink>();
        const linkElements: Array<{ link: GraphLink; path: SVGPathElement; label?: SVGTextElement }> = [];
        const nodeElements: Array<{ node: GraphNode; outline: SVGRectElement; rect: SVGRectElement; label?: SVGTextElement }> = [];

        const refreshGraph = (): void => {
            linkElements.forEach(({ link, path, label }) => {
                const isHovered = this.hoveredLinkKey === link.key;
                setAttributes(path, {
                    d: linkPathGenerator(link) || "",
                    stroke: isHovered ? this.getHoveredLinkColor(link.color) : link.color,
                    "stroke-opacity": this.host.colorPalette.isHighContrast
                        ? 1
                        : isHovered
                            ? Math.min(0.92, Math.max(link.opacity, 0.72))
                            : link.opacity,
                    "stroke-width": Math.max(1, link.width || 1),
                });

                if (label) {
                    setAttributes(label, this.getLinkLabelAttributes(link));
                }
            });

            nodeElements.forEach(({ node, outline, rect, label }) => {
                const x0 = node.x0 || 0;
                const x1 = node.x1 || 0;
                const y0 = node.y0 || 0;
                const y1 = node.y1 || 0;
                const isSelected = selectedNodeName === node.id;
                const isFocused = this.focusedNodeId === node.id;

                setAttributes(outline, {
                    x: x0 - 3,
                    y: y0 - 3,
                    width: Math.max(1, x1 - x0) + 6,
                    height: Math.max(4, y1 - y0) + 6,
                    rx: 7,
                    ry: 7,
                    fill: "none",
                    stroke: this.host.colorPalette.isHighContrast
                        ? this.host.colorPalette.foreground.value
                        : "#0f172a",
                    "stroke-width": (isSelected || isFocused) ? 2 : 0,
                    opacity: (isSelected || isFocused) ? 1 : 0,
                });

                setAttributes(rect, {
                    x: x0,
                    y: y0,
                    width: Math.max(1, x1 - x0),
                    height: Math.max(4, y1 - y0),
                    fill: node.color,
                    "fill-opacity": node.opacity,
                    stroke: this.host.colorPalette.isHighContrast
                        ? this.host.colorPalette.foreground.value
                        : "rgba(15, 23, 42, 0.18)",
                    "stroke-width": 1,
                });

                if (label) {
                    setAttributes(label, this.getNodeLabelAttributes(node, viewport.width));
                }
            });
        };

        sankeyGraph.links.forEach((link: GraphLink) => {
            const path = createSvgElement("path");

            path.classList.add("sankey-link");

            setAttributes(path, {
                fill: "none",
                "stroke-linecap": "butt",
                tabindex: 0,
                focusable: "true",
                role: this.allowInteractions ? "button" : "img",
                "aria-label": this.getLinkAriaLabel(link),
            });

            const activateLink = (multiSelect: boolean): void => {
                this.selectLink(link, multiSelect);
            };

            path.addEventListener("click", (event: MouseEvent) => {
                if (!this.allowInteractions) {
                    return;
                }

                event.stopPropagation();
                activateLink(this.isMultiSelect(event));
            });

            path.addEventListener("contextmenu", (event: MouseEvent) => {
                event.preventDefault();
                event.stopPropagation();
                this.showContextMenu(link.contextMenuSelectionId, event);
            });

            path.addEventListener("mouseenter", (event: MouseEvent) => {
                this.hoveredLinkKey = link.key;
                this.showLinkTooltip(link, event.clientX, event.clientY);
                refreshGraph();
            });

            path.addEventListener("mousemove", (event: MouseEvent) => {
                this.moveTooltip(link.selectionIds, event.clientX, event.clientY);
            });

            path.addEventListener("mouseleave", () => {
                if (this.hoveredLinkKey === link.key) {
                    this.hoveredLinkKey = undefined;
                    this.hideTooltip();
                    refreshGraph();
                }
            });

            path.addEventListener("focus", () => {
                this.hoveredLinkKey = link.key;
                this.showLinkTooltipAtElement(link, path);
                refreshGraph();
            });

            path.addEventListener("blur", () => {
                if (this.hoveredLinkKey === link.key) {
                    this.hoveredLinkKey = undefined;
                    this.hideTooltip();
                    refreshGraph();
                }
            });

            path.addEventListener("keydown", (event: KeyboardEvent) => {
                if (this.allowInteractions && this.isActivationKey(event)) {
                    event.preventDefault();
                    event.stopPropagation();
                    activateLink(this.isMultiSelect(event));
                    return;
                }

                if (this.isContextMenuKey(event)) {
                    event.preventDefault();
                    event.stopPropagation();
                    this.showContextMenuAtElement(link.contextMenuSelectionId, path);
                }
            });

            linkGroup.appendChild(path);

            const linkLabelText = this.getLinkLabelText(link);
            let label: SVGTextElement | undefined;

            if (linkLabelText) {
                label = createSvgElement("text");
                label.classList.add("sankey-link-label");
                label.textContent = linkLabelText;
                linkLabelGroup.appendChild(label);
            }

            linkElements.push({ link, path, label });
        });

        sankeyGraph.nodes.forEach((node: GraphNode) => {
            const isSelected = selectedNodeName === node.id;

            const outline = createSvgElement("rect");
            outline.classList.add("sankey-node-outline");
            selectionGroup.appendChild(outline);

            const rect = createSvgElement("rect");
            rect.classList.add("sankey-node");

            setAttributes(rect, {
                rx: 4,
                ry: 4,
                tabindex: 0,
                focusable: "true",
                role: this.allowInteractions ? "button" : "img",
                "aria-label": this.getNodeAriaLabel(node),
            });

            const activateNode = (multiSelect: boolean): void => {
                if (isSelected && !multiSelect) {
                    this.clearSelection();
                    return;
                }

                this.selectNode(node, multiSelect);
            };

            rect.addEventListener("pointerdown", (event: PointerEvent) => {
                this.startNodeDrag(
                    event,
                    rect,
                    node,
                    sankeyGraph,
                    sankeyLayout,
                    dragBounds,
                    refreshGraph,
                );
            });

            rect.addEventListener("click", (event: MouseEvent) => {
                if (!this.allowInteractions) {
                    return;
                }

                event.stopPropagation();
                if (this.suppressClickSelection) {
                    this.suppressClickSelection = false;
                    return;
                }

                activateNode(this.isMultiSelect(event));
            });

            rect.addEventListener("contextmenu", (event: MouseEvent) => {
                event.preventDefault();
                event.stopPropagation();
                this.showContextMenu(node.contextMenuSelectionId, event);
            });

            rect.addEventListener("mouseenter", (event: MouseEvent) => {
                this.showNodeTooltip(node, event.clientX, event.clientY);
            });

            rect.addEventListener("mousemove", (event: MouseEvent) => {
                this.moveTooltip(node.selectionIds, event.clientX, event.clientY);
            });

            rect.addEventListener("mouseleave", () => {
                this.hideTooltip();
            });

            rect.addEventListener("focus", () => {
                this.focusedNodeId = node.id;
                this.showNodeTooltipAtElement(node, rect);
                refreshGraph();
            });

            rect.addEventListener("blur", () => {
                if (this.focusedNodeId === node.id) {
                    this.focusedNodeId = undefined;
                    this.hideTooltip();
                    refreshGraph();
                }
            });

            rect.addEventListener("keydown", (event: KeyboardEvent) => {
                if (this.allowInteractions && this.isActivationKey(event)) {
                    event.preventDefault();
                    event.stopPropagation();
                    activateNode(this.isMultiSelect(event));
                    return;
                }

                if (this.isContextMenuKey(event)) {
                    event.preventDefault();
                    event.stopPropagation();
                    this.showContextMenuAtElement(node.contextMenuSelectionId, rect);
                }
            });

            nodeGroup.appendChild(rect);

            const nodeLabelText = this.getNodeLabelText(node);
            let label: SVGTextElement | undefined;

            if (nodeLabelText) {
                label = createSvgElement("text");

                label.classList.add("sankey-label");
                label.textContent = nodeLabelText;
                labelGroup.appendChild(label);
            }

            nodeElements.push({ node, outline, rect, label });
        });

        refreshGraph();
    }

    private renderWarnings(warnings: string[]): void {
        if (warnings.length === 0) {
            this.warning.textContent = "";
            this.warning.classList.remove("is-visible");
            return;
        }

        const displayedWarnings = warnings.slice(0, 2);
        const suffix = warnings.length > displayedWarnings.length ? ` (+${warnings.length - displayedWarnings.length} more)` : "";
        this.warning.textContent = `${displayedWarnings.join(" ")}${suffix}`;
        this.warning.classList.add("is-visible");
    }

    private showLandingPage(): void {
        this.warning.textContent = "";
        this.warning.classList.remove("is-visible");
        this.landingPage.classList.add("is-visible");
    }

    private hideLandingPage(): void {
        this.landingPage.classList.remove("is-visible");
    }

    private renderMessage(message: string): void {
        this.warning.textContent = "";
        this.warning.classList.remove("is-visible");

        const messageGroup = createSvgElement("g");
        messageGroup.classList.add("sankey-empty");

        const background = createSvgElement("rect");
        setAttributes(background, {
            x: 16,
            y: 16,
            width: Math.max(120, this.svg.viewBox.baseVal.width - 32),
            height: Math.max(80, this.svg.viewBox.baseVal.height - 48),
            rx: 12,
            ry: 12,
            fill: this.host.colorPalette.isHighContrast
                ? this.host.colorPalette.background.value
                : "rgba(248, 250, 252, 0.92)",
            stroke: this.host.colorPalette.isHighContrast
                ? this.host.colorPalette.foreground.value
                : "rgba(148, 163, 184, 0.35)",
        });

        const text = createSvgElement("text");
        text.textContent = message;
        setAttributes(text, {
            x: this.svg.viewBox.baseVal.width / 2,
            y: this.svg.viewBox.baseVal.height / 2,
            "text-anchor": "middle",
            "dominant-baseline": "middle",
            "font-size": 13,
            fill: this.host.colorPalette.isHighContrast
                ? this.host.colorPalette.foreground.value
                : "#334155",
        });

        messageGroup.appendChild(background);
        messageGroup.appendChild(text);
        this.svg.appendChild(messageGroup);
    }

    private handleWheelZoom(event: WheelEvent): void {
        if (!this.allowInteractions || !this.currentGraphViewport || !this.currentViewport) {
            return;
        }

        if (!(event.ctrlKey || event.metaKey)) {
            return;
        }

        event.preventDefault();
        event.stopPropagation();
        this.hideTooltip();

        const nextScale = clamp(
            this.zoomState.scale * Math.exp(-event.deltaY * wheelZoomSensitivity),
            minZoomScale,
            maxZoomScale,
        );
        this.zoomTo(nextScale, event.clientX, event.clientY);
    }

    private startViewportPan(event: PointerEvent): void {
        if (
            !this.allowInteractions
            || event.button !== 0
            || !this.currentViewport
            || !this.currentGraphViewport
            || this.zoomState.scale <= minZoomScale
        ) {
            return;
        }

        event.preventDefault();
        this.hideTooltip();

        const startClientX = event.clientX;
        const startClientY = event.clientY;
        const startTranslateX = this.zoomState.translateX;
        const startTranslateY = this.zoomState.translateY;
        let moved = false;

        this.svg.classList.add("is-panning");
        this.svg.setPointerCapture(event.pointerId);

        const handlePointerMove = (moveEvent: PointerEvent): void => {
            const deltaX = moveEvent.clientX - startClientX;
            const deltaY = moveEvent.clientY - startClientY;

            if (!moved && Math.max(Math.abs(deltaX), Math.abs(deltaY)) < dragStartThreshold) {
                return;
            }

            moved = true;
            this.zoomState.translateX = startTranslateX + deltaX;
            this.zoomState.translateY = startTranslateY + deltaY;
            this.normalizeZoomState(this.currentViewport!);
            this.applyGraphViewportTransform();
        };

        const finishPan = (): void => {
            this.svg.classList.remove("is-panning");

            if (this.svg.hasPointerCapture(event.pointerId)) {
                this.svg.releasePointerCapture(event.pointerId);
            }

            this.svg.removeEventListener("pointermove", handlePointerMove);
            this.svg.removeEventListener("pointerup", handlePointerUp);
            this.svg.removeEventListener("pointercancel", handlePointerCancel);

            if (moved) {
                this.suppressClickSelection = true;
            }
        };

        const handlePointerUp = (): void => {
            finishPan();
        };

        const handlePointerCancel = (): void => {
            finishPan();
        };

        this.svg.addEventListener("pointermove", handlePointerMove);
        this.svg.addEventListener("pointerup", handlePointerUp);
        this.svg.addEventListener("pointercancel", handlePointerCancel);
    }

    private getSvgPointFromClient(clientX: number, clientY: number): Point {
        const svgBounds = this.svg.getBoundingClientRect();
        return {
            x: clientX - svgBounds.left,
            y: clientY - svgBounds.top,
        };
    }

    private getGraphPointFromClient(clientX: number, clientY: number): Point {
        const svgPoint = this.getSvgPointFromClient(clientX, clientY);
        return {
            x: (svgPoint.x - this.zoomState.translateX) / this.zoomState.scale,
            y: (svgPoint.y - this.zoomState.translateY) / this.zoomState.scale,
        };
    }

    private zoomTo(nextScale: number, clientX: number, clientY: number): void {
        if (!this.currentGraphViewport || !this.currentViewport) {
            return;
        }

        const normalizedScale = clamp(nextScale, minZoomScale, maxZoomScale);
        if (Math.abs(normalizedScale - this.zoomState.scale) < 0.001) {
            return;
        }

        const svgPoint = this.getSvgPointFromClient(clientX, clientY);
        const graphX = (svgPoint.x - this.zoomState.translateX) / this.zoomState.scale;
        const graphY = (svgPoint.y - this.zoomState.translateY) / this.zoomState.scale;

        this.zoomState.scale = normalizedScale;
        this.zoomState.translateX = svgPoint.x - (graphX * normalizedScale);
        this.zoomState.translateY = svgPoint.y - (graphY * normalizedScale);
        this.normalizeZoomState(this.currentViewport);
        this.applyGraphViewportTransform();
    }

    private normalizeZoomState(viewport: IViewport): void {
        this.zoomState.scale = clamp(this.zoomState.scale, minZoomScale, maxZoomScale);

        if (this.zoomState.scale <= minZoomScale + 0.001) {
            this.zoomState.scale = minZoomScale;
            this.zoomState.translateX = 0;
            this.zoomState.translateY = 0;
            return;
        }

        const minTranslateX = viewport.width - (viewport.width * this.zoomState.scale);
        const minTranslateY = viewport.height - (viewport.height * this.zoomState.scale);

        this.zoomState.translateX = clamp(this.zoomState.translateX, minTranslateX, 0);
        this.zoomState.translateY = clamp(this.zoomState.translateY, minTranslateY, 0);
    }

    private applyGraphViewportTransform(): void {
        if (this.currentGraphViewport) {
            setAttributes(this.currentGraphViewport, {
                transform: `matrix(${this.zoomState.scale} 0 0 ${this.zoomState.scale} ${this.zoomState.translateX} ${this.zoomState.translateY})`,
            });
        }

        const isPannable = this.allowInteractions && Boolean(this.currentGraphViewport) && this.zoomState.scale > minZoomScale + 0.001;
        this.svg.classList.toggle("is-pannable", isPannable);
        if (!isPannable) {
            this.svg.classList.remove("is-panning");
        }
    }

    private startNodeDrag(
        event: PointerEvent,
        rect: SVGRectElement,
        node: GraphNode,
        graph: SankeyGraph<VisualNode, VisualLink>,
        sankeyLayout: SankeyGenerator,
        dragBounds: DragBounds,
        refreshGraph: () => void,
    ): void {
        if (!this.allowInteractions || event.button !== 0) {
            return;
        }

        event.preventDefault();
        event.stopPropagation();
        this.hideTooltip();

        const startClientY = event.clientY;
        const startClientX = event.clientX;
        const startGraphPoint = this.getGraphPointFromClient(startClientX, startClientY);
        const pointerOffsetX = startGraphPoint.x - (node.x0 || dragBounds.left);
        const pointerOffsetY = startGraphPoint.y - (node.y0 || dragBounds.top);
        let moved = false;
        let animationFrameId: number | undefined;

        rect.classList.add("is-dragging");
        rect.setPointerCapture(event.pointerId);

        const renderDragFrame = (): void => {
            animationFrameId = undefined;
            sankeyLayout.update(graph);
            refreshGraph();
        };

        const scheduleDragFrame = (): void => {
            if (animationFrameId !== undefined) {
                return;
            }

            animationFrameId = window.requestAnimationFrame(renderDragFrame);
        };

        const handlePointerMove = (moveEvent: PointerEvent): void => {
            const deltaX = moveEvent.clientX - startClientX;
            const deltaY = moveEvent.clientY - startClientY;

            if (!moved && Math.max(Math.abs(deltaX), Math.abs(deltaY)) < dragStartThreshold) {
                return;
            }

            moved = true;

            const nodeWidth = Math.max(1, (node.x1 || 0) - (node.x0 || 0));
            const nodeHeight = Math.max(1, (node.y1 || 0) - (node.y0 || 0));
            const pointerPoint = this.getGraphPointFromClient(moveEvent.clientX, moveEvent.clientY);
            const desiredX0 = clamp(pointerPoint.x - pointerOffsetX, dragBounds.left, dragBounds.right - nodeWidth);
            const desiredY0 = clamp(pointerPoint.y - pointerOffsetY, dragBounds.top, dragBounds.bottom - nodeHeight);

            node.x0 = desiredX0;
            node.x1 = desiredX0 + nodeWidth;
            node.y0 = desiredY0;
            node.y1 = desiredY0 + nodeHeight;
            scheduleDragFrame();
        };

        const finishDrag = (): void => {
            rect.classList.remove("is-dragging");

            if (rect.hasPointerCapture(event.pointerId)) {
                rect.releasePointerCapture(event.pointerId);
            }

            rect.removeEventListener("pointermove", handlePointerMove);
            rect.removeEventListener("pointerup", handlePointerUp);
            rect.removeEventListener("pointercancel", handlePointerCancel);

            if (animationFrameId !== undefined) {
                window.cancelAnimationFrame(animationFrameId);
                animationFrameId = undefined;
            }

            if (!moved) {
                return;
            }

            clampNodeToBounds(node, dragBounds);
            sankeyLayout.update(graph);
            refreshGraph();
            this.suppressClickSelection = true;
            this.persistNodePositions(graph, dragBounds);
        };

        const handlePointerUp = (): void => {
            finishDrag();
        };

        const handlePointerCancel = (): void => {
            finishDrag();
        };

        rect.addEventListener("pointermove", handlePointerMove);
        rect.addEventListener("pointerup", handlePointerUp);
        rect.addEventListener("pointercancel", handlePointerCancel);
    }

    private applyPersistedNodePositions(
        graph: SankeyGraph<VisualNode, VisualLink>,
        nodePositionMap: NodePositionMap,
        dragBounds: DragBounds,
    ): void {
        const layoutWidth = Math.max(1, dragBounds.right - dragBounds.left);
        const layoutHeight = Math.max(1, dragBounds.bottom - dragBounds.top);

        graph.nodes.forEach((currentNode: GraphNode) => {
            const storedPosition = nodePositionMap[currentNode.id];
            if (!storedPosition) {
                return;
            }

            const nodeWidth = Math.max(1, (currentNode.x1 || 0) - (currentNode.x0 || 0));
            const nodeHeight = Math.max(1, (currentNode.y1 || 0) - (currentNode.y0 || 0));

            if (isFiniteNumber(storedPosition.x)) {
                const availableWidth = Math.max(1, layoutWidth - nodeWidth);
                const desiredX0 = dragBounds.left + clamp(storedPosition.x, 0, 1) * availableWidth;
                currentNode.x0 = desiredX0;
                currentNode.x1 = desiredX0 + nodeWidth;
            }

            if (isFiniteNumber(storedPosition.y)) {
                const availableHeight = Math.max(1, layoutHeight - nodeHeight);
                const desiredY0 = dragBounds.top + clamp(storedPosition.y, 0, 1) * availableHeight;
                currentNode.y0 = desiredY0;
                currentNode.y1 = desiredY0 + nodeHeight;
            }

            clampNodeToBounds(currentNode, dragBounds);
        });
    }

    private persistNodePositions(graph: SankeyGraph<VisualNode, VisualLink>, dragBounds: DragBounds): void {
        const existingPositionMap = this.parseNodePositionMap(this.formattingSettings.persistedState.nodePositionMap.value);
        const nextPositionMap: NodePositionMap = {
            ...existingPositionMap,
            ...this.buildNodePositionMap(graph, dragBounds),
        };
        const currentSerialized = this.serializeNodePositionMap(existingPositionMap);
        const nextSerialized = this.serializeNodePositionMap(nextPositionMap);

        if (nextSerialized === currentSerialized) {
            return;
        }

        this.host.persistProperties({
            merge: [
                {
                    objectName: "persistedState",
                    selector: undefined,
                    properties: {
                        nodePositionMap: nextSerialized,
                    },
                },
            ],
        });
    }

    private buildNodePositionMap(graph: SankeyGraph<VisualNode, VisualLink>, dragBounds: DragBounds): NodePositionMap {
        const layoutWidth = Math.max(1, dragBounds.right - dragBounds.left);
        const layoutHeight = Math.max(1, dragBounds.bottom - dragBounds.top);
        const positionMap: NodePositionMap = {};

        graph.nodes.forEach((node: GraphNode) => {
            const nodeWidth = Math.max(1, (node.x1 || 0) - (node.x0 || 0));
            const nodeHeight = Math.max(1, (node.y1 || 0) - (node.y0 || 0));
            const availableWidth = Math.max(1, layoutWidth - nodeWidth);
            const availableHeight = Math.max(1, layoutHeight - nodeHeight);

            positionMap[node.id] = {
                x: clamp(((node.x0 || dragBounds.left) - dragBounds.left) / availableWidth, 0, 1),
                y: clamp(((node.y0 || dragBounds.top) - dragBounds.top) / availableHeight, 0, 1),
            };
        });

        return positionMap;
    }

    private parseNodePositionMap(jsonValue: string | undefined): NodePositionMap {
        if (!jsonValue) {
            return {};
        }

        try {
            const parsed = JSON.parse(jsonValue) as Record<string, unknown>;
            const nodePositionMap: NodePositionMap = {};

            Object.keys(parsed || {}).forEach((key: string) => {
                const value = parsed[key];

                if (typeof value === "number" && Number.isFinite(value)) {
                    nodePositionMap[key] = { y: clamp(value, 0, 1) };
                    return;
                }

                if (typeof value === "object" && value !== null) {
                    const position = value as Partial<NodePosition>;
                    const normalizedPosition: NodePosition = {};

                    if (isFiniteNumber(position.x)) {
                        normalizedPosition.x = clamp(position.x, 0, 1);
                    }

                    if (isFiniteNumber(position.y)) {
                        normalizedPosition.y = clamp(position.y, 0, 1);
                    }

                    if (normalizedPosition.x !== undefined || normalizedPosition.y !== undefined) {
                        nodePositionMap[key] = normalizedPosition;
                    }
                }
            });

            return nodePositionMap;
        } catch {
            return {};
        }
    }

    private serializeNodePositionMap(nodePositionMap: NodePositionMap): string {
        const sortedKeys = Object.keys(nodePositionMap).sort((left: string, right: string) => left.localeCompare(right));
        const normalized: NodePositionMap = {};

        sortedKeys.forEach((key: string) => {
            const currentPosition = nodePositionMap[key];
            const nextPosition: NodePosition = {};

            if (isFiniteNumber(currentPosition?.x)) {
                nextPosition.x = Math.round(clamp(currentPosition.x, 0, 1) * nodePositionPrecision) / nodePositionPrecision;
            }

            if (isFiniteNumber(currentPosition?.y)) {
                nextPosition.y = Math.round(clamp(currentPosition.y, 0, 1) * nodePositionPrecision) / nodePositionPrecision;
            }

            normalized[key] = nextPosition;
        });

        return JSON.stringify(normalized);
    }

    private parseSelectionSyncState(jsonValue: string | undefined): SyncedSelectionState | undefined {
        if (!jsonValue) {
            return undefined;
        }

        try {
            const parsed = JSON.parse(jsonValue) as Partial<SyncedSelectionState> & { style?: Partial<PersistedStyle> };
            const key = typeof parsed.key === "string" ? parsed.key.trim() : "";
            const controlStyle = this.parseSyncStyle(parsed.controlStyle ?? parsed.targetStyle ?? parsed.style);
            const targetStyle = this.parseSyncStyle(parsed.targetStyle ?? parsed.controlStyle ?? parsed.style);

            if (!key || !controlStyle || !targetStyle) {
                return undefined;
            }

            return {
                key,
                controlStyle,
                targetStyle,
            };
        } catch {
            return undefined;
        }
    }

    private serializeSelectionSyncState(syncState: SyncedSelectionState): string {
        return JSON.stringify({
            key: syncState.key,
            controlStyle: {
                color: syncState.controlStyle.color,
                opacity: Math.round(Math.max(0, Math.min(100, syncState.controlStyle.opacity))),
            },
            targetStyle: {
                color: syncState.targetStyle.color,
                opacity: Math.round(Math.max(0, Math.min(100, syncState.targetStyle.opacity))),
            },
        });
    }

    private parseSyncStyle(style: Partial<PersistedStyle> | undefined): PersistedStyle | undefined {
        const color = typeof style?.color === "string" ? style.color.trim() : "";
        const opacity = typeof style?.opacity === "number"
            ? Math.round(Math.max(0, Math.min(100, style.opacity)))
            : undefined;

        if (!color || opacity === undefined) {
            return undefined;
        }

        return { color, opacity };
    }

    private toPersistedStyle(color: string, opacity: number): PersistedStyle {
        return {
            color,
            opacity: Math.round(Math.max(0, Math.min(1, opacity)) * 100),
        };
    }

    private parsePersistedStyle(jsonValue: string | undefined, fallback: PersistedStyle): PersistedStyle {
        if (!jsonValue) {
            return fallback;
        }

        try {
            const parsed = JSON.parse(jsonValue) as Partial<PersistedStyle>;
            const color = typeof parsed.color === "string" && parsed.color.trim().length > 0
                ? parsed.color.trim()
                : fallback.color;
            const opacity = typeof parsed.opacity === "number" && Number.isFinite(parsed.opacity)
                ? Math.round(Math.max(0, Math.min(100, parsed.opacity)))
                : fallback.opacity;

            return { color, opacity };
        } catch {
            return fallback;
        }
    }

    private serializePersistedStyle(style: PersistedStyle): string {
        return JSON.stringify({
            color: style.color,
            opacity: Math.round(Math.max(0, Math.min(100, style.opacity))),
        });
    }

    private getConfiguredNodeFormatStyle(targetStyle: PersistedStyle): PersistedStyle {
        return {
            color: this.formattingSettings.defaultNodes.fill.value.value,
            opacity: targetStyle.opacity,
        };
    }

    private getSelectedPaletteMode(): NodePaletteMode {
        const selectedPaletteMode = this.formattingSettings.defaultNodes.paletteMode.value;
        return typeof selectedPaletteMode === "string" && isNodePaletteMode(selectedPaletteMode)
            ? selectedPaletteMode
            : defaultNodePaletteMode;
    }

    private getDisplayedFormatStyle(
        target: FormatTargetState,
        configuredStyle: PersistedStyle,
        syncState: SyncedSelectionState | undefined,
    ): PersistedStyle {
        if (!target.showStyleControls) {
            return target.style;
        }

        if (!syncState || syncState.key !== target.key) {
            return target.style;
        }

        if (!this.stylesMatch(configuredStyle, syncState.controlStyle)) {
            return configuredStyle;
        }

        return target.style;
    }

    private getActiveNodeTargetName(
        visualData: SankeyVisualData,
        persistedSelectedNodeName: string,
    ): string | undefined {
        if (this.pendingClearSelection) {
            if (persistedSelectedNodeName.length === 0) {
                this.pendingClearSelection = false;
            }

            return undefined;
        }

        if (this.pendingSelectedNodeName) {
            if (persistedSelectedNodeName === this.pendingSelectedNodeName) {
                this.pendingSelectedNodeName = undefined;
            }

            if (visualData.nodes.some((node: VisualNode) => node.id === this.pendingSelectedNodeName)) {
                return this.pendingSelectedNodeName;
            }

            this.pendingSelectedNodeName = undefined;
        }

        if (
            persistedSelectedNodeName.length > 0
            && visualData.nodes.some((node: VisualNode) => node.id === persistedSelectedNodeName)
        ) {
            return persistedSelectedNodeName;
        }

        return undefined;
    }

    private getNodeFormatTarget(
        visualData: SankeyVisualData,
        nodeGlobalStyle: PersistedStyle,
        selectedNodeName?: string,
    ): FormatTargetState {
        if (selectedNodeName) {
            const selectedNode = visualData.nodes.find((node: VisualNode) => node.id === selectedNodeName);
            if (selectedNode) {
                return {
                    key: selectedNode.id,
                    label: `Selected node: ${selectedNode.label}`,
                    style: this.toPersistedStyle(selectedNode.color, selectedNode.opacity),
                    showStyleControls: true,
                    showFillControl: true,
                };
            }
        }

        if (
            !this.formattingSettings.defaultNodes.useThemePalette.value
            && this.getSelectedPaletteMode() === defaultNodePaletteMode
        ) {
            return {
                key: allNodesFormatKey,
                label: "All nodes",
                style: nodeGlobalStyle,
                showStyleControls: true,
                showFillControl: true,
            };
        }

        if (!this.formattingSettings.defaultNodes.useThemePalette.value) {
            return {
                key: noNodeSelectionFormatKey,
                label: "All nodes palette",
                style: nodeGlobalStyle,
                showStyleControls: false,
                showFillControl: false,
            };
        }

        return {
            key: noNodeSelectionFormatKey,
            label: "Select a node to override",
            style: nodeGlobalStyle,
            showStyleControls: false,
            showFillControl: false,
        };
    }

    private getNodeLabelText(node: GraphNode): string | undefined {
        if (!this.formattingSettings.labels.show.value) {
            return undefined;
        }

        const showText = this.formattingSettings.labels.showText.value;
        const showValue = this.formattingSettings.labels.showValue.value;

        if (!showText && !showValue) {
            return undefined;
        }

        if (showText && showValue) {
            return `${node.label} (${this.formatConfiguredValue(this.getNodeDisplayValue(node))})`;
        }

        if (showText) {
            return node.label;
        }

        return this.formatConfiguredValue(this.getNodeDisplayValue(node));
    }

    private getLinkLabelText(link: GraphLink): string | undefined {
        if (!this.formattingSettings.linkLabels.show.value) {
            return undefined;
        }

        const showText = this.formattingSettings.linkLabels.showText.value;
        const showValue = this.formattingSettings.linkLabels.showValue.value;
        const showAsPercentage = this.formattingSettings.valueFormatting.showAsPercentage.value;

        if (!showText && !showValue) {
            return undefined;
        }

        const sourceLabel = this.getEndpointLabel(link.source);
        const targetLabel = this.getEndpointLabel(link.target);
        const linkText = `${sourceLabel} -> ${targetLabel}`;
        const formattedValue = this.formatConfiguredValue(link.value, showAsPercentage);

        if (showText && showValue) {
            return `${linkText} (${formattedValue})`;
        }

        if (showText) {
            return linkText;
        }

        return formattedValue;
    }

    private formatConfiguredValue(
        value: number,
        showAsPercentage: boolean = this.formattingSettings.valueFormatting.showAsPercentage.value,
    ): string {
        return showAsPercentage
            ? this.percentageValueFormatter.format(value)
            : this.modelValueFormatter.format(value);
    }

    private getHoveredLinkColor(color: string): string {
        return this.host.colorPalette.isHighContrast ? color : mixWithWhite(color, 0.28);
    }

    private getNodeTooltipItems(node: GraphNode): TooltipItem[] {
        return [
            {
                header: node.label,
                displayName: "Value",
                value: this.formatConfiguredValue(this.getNodeDisplayValue(node)),
                color: node.color,
            },
            {
                displayName: "Incoming",
                value: this.formatConfiguredValue(node.incomingValue),
            },
            {
                displayName: "Outgoing",
                value: this.formatConfiguredValue(node.outgoingValue),
            },
        ];
    }

    private getLinkTooltipItems(link: GraphLink): TooltipItem[] {
        const sourceLabel = this.getEndpointLabel(link.source);
        const targetLabel = this.getEndpointLabel(link.target);
        return [
            {
                header: `${sourceLabel} -> ${targetLabel}`,
                displayName: "Value",
                value: this.formatConfiguredValue(link.value),
                color: this.getHoveredLinkColor(link.color),
            },
        ];
    }

    private showNodeTooltip(node: GraphNode, clientX: number, clientY: number): void {
        this.showTooltip(this.getNodeTooltipItems(node), node.selectionIds, clientX, clientY);
    }

    private showNodeTooltipAtElement(node: GraphNode, element: SVGGraphicsElement): void {
        const [x, y] = this.getElementCenterCoordinates(element);
        this.showNodeTooltip(node, x, y);
    }

    private showLinkTooltip(link: GraphLink, clientX: number, clientY: number): void {
        this.showTooltip(this.getLinkTooltipItems(link), link.selectionIds, clientX, clientY);
    }

    private showLinkTooltipAtElement(link: GraphLink, element: SVGGraphicsElement): void {
        const [x, y] = this.getElementCenterCoordinates(element);
        this.showLinkTooltip(link, x, y);
    }

    private showTooltip(dataItems: TooltipItem[], identities: ISelectionId[], clientX: number, clientY: number): void {
        if (!this.host.tooltipService.enabled()) {
            return;
        }

        this.host.tooltipService.show({
            coordinates: [clientX, clientY],
            isTouchEvent: false,
            dataItems,
            identities,
        });
    }

    private moveTooltip(identities: ISelectionId[], clientX: number, clientY: number): void {
        if (!this.host.tooltipService.enabled()) {
            return;
        }

        this.host.tooltipService.move({
            coordinates: [clientX, clientY],
            isTouchEvent: false,
            identities,
        });
    }

    private hideTooltip(): void {
        if (!this.host.tooltipService.enabled()) {
            return;
        }

        this.host.tooltipService.hide({
            isTouchEvent: false,
            immediately: false,
        });
    }

    private getNodeLabelAttributes(node: GraphNode, viewportWidth: number): Record<string, string | number> {
        const x0 = node.x0 || 0;
        const x1 = node.x1 || 0;
        const y0 = node.y0 || 0;
        const y1 = node.y1 || 0;
        const anchorOnLeft = x0 < viewportWidth / 2;
        const labelX = anchorOnLeft ? x1 + 8 : x0 - 8;
        const labelAnchor = anchorOnLeft ? "start" : "end";

        return {
            x: labelX,
            y: (y0 + y1) / 2,
            "text-anchor": labelAnchor,
            "dominant-baseline": "middle",
            "font-size": this.formattingSettings.labels.fontSize.value,
            fill: this.host.colorPalette.isHighContrast
                ? this.host.colorPalette.foreground.value
                : this.formattingSettings.labels.fill.value.value,
        };
    }

    private getLinkLabelAttributes(link: GraphLink): Record<string, string | number> {
        const sourceNode = typeof link.source === "string" ? undefined : link.source;
        const targetNode = typeof link.target === "string" ? undefined : link.target;
        const sourceX = sourceNode?.x1 || 0;
        const targetX = targetNode?.x0 || sourceX;
        const centerX = sourceX + ((targetX - sourceX) / 2);
        const centerY = ((link.y0 || 0) + (link.y1 || 0)) / 2;

        return {
            x: centerX,
            y: centerY,
            "text-anchor": "middle",
            "dominant-baseline": "middle",
            "font-size": this.formattingSettings.linkLabels.fontSize.value,
            fill: this.host.colorPalette.isHighContrast
                ? this.host.colorPalette.foreground.value
                : this.formattingSettings.linkLabels.fill.value.value,
        };
    }

    private getNodeDisplayValue(node: GraphNode): number {
        if (isFiniteNumber(node.value)) {
            return node.value;
        }

        return Math.max(node.incomingValue, node.outgoingValue);
    }

    private getEndpointLabel(nodeReference: string | VisualNode): string {
        if (typeof nodeReference === "string") {
            return nodeReference.replace(layerSuffixPattern, "");
        }

        return nodeReference.label;
    }

    private getNodeAriaLabel(node: GraphNode): string {
        return `Node ${node.label}, value ${this.formatConfiguredValue(this.getNodeDisplayValue(node))}`;
    }

    private getLinkAriaLabel(link: GraphLink): string {
        return `Flow ${this.getEndpointLabel(link.source)} to ${this.getEndpointLabel(link.target)}, value ${this.formatConfiguredValue(link.value)}`;
    }

    private getElementCenterCoordinates(element: SVGGraphicsElement): [number, number] {
        const bounds = element.getBoundingClientRect();
        return [
            bounds.left + (bounds.width / 2),
            bounds.top + (bounds.height / 2),
        ];
    }

    private isActivationKey(event: KeyboardEvent): boolean {
        return event.key === "Enter" || event.key === " ";
    }

    private isContextMenuKey(event: KeyboardEvent): boolean {
        return event.key === "ContextMenu" || (event.shiftKey && event.key === "F10");
    }

    private isMultiSelect(event: MouseEvent | KeyboardEvent): boolean {
        return event.ctrlKey || event.metaKey;
    }

    private showContextMenu(selectionId: ISelectionId | undefined, event: MouseEvent): void {
        this.showContextMenuAtCoordinates(selectionId, event.clientX, event.clientY);
    }

    private showContextMenuAtElement(selectionId: ISelectionId | undefined, element: SVGGraphicsElement): void {
        const [x, y] = this.getElementCenterCoordinates(element);
        this.showContextMenuAtCoordinates(selectionId, x, y);
    }

    private showContextMenuAtCoordinates(selectionId: ISelectionId | undefined, x: number, y: number): void {
        void this.selectionManager.showContextMenu(
            selectionId || ({} as ISelectionId),
            {
                x,
                y,
            },
        );
    }

    private selectVisualDataPoint(selectionIds: ISelectionId[], multiSelect: boolean): void {
        if (selectionIds.length === 0) {
            return;
        }

        void this.selectionManager.select(selectionIds, multiSelect);
    }

    private persistSelectedNodeName(nodeId?: string): void {
        this.pendingSelectedNodeName = nodeId;
        this.pendingClearSelection = !nodeId;
        this.host.persistProperties({
            merge: [
                {
                    objectName: "persistedState",
                    selector: undefined,
                    properties: {
                        selectedNodeName: nodeId || "",
                        selectedLinkName: "",
                        ...(nodeId ? {} : {
                            nodeFormatSyncState: "",
                            linkFormatSyncState: "",
                        }),
                    },
                },
            ],
        });
    }

    private selectNode(node: VisualNode, multiSelect: boolean): void {
        if (!this.allowInteractions) {
            return;
        }

        this.persistSelectedNodeName(node.id);
        this.selectVisualDataPoint(node.selectionIds, multiSelect);
    }

    private selectLink(link: VisualLink, multiSelect: boolean): void {
        if (!this.allowInteractions) {
            return;
        }

        this.persistSelectedNodeName(undefined);
        this.selectVisualDataPoint(link.selectionIds, multiSelect);
    }

    private clearSelection(): void {
        if (!this.allowInteractions) {
            return;
        }

        this.persistSelectedNodeName(undefined);
        void this.selectionManager.clear();
    }

    private stylesMatch(left: PersistedStyle, right: PersistedStyle): boolean {
        return left.color === right.color && left.opacity === right.opacity;
    }

    private addPersistedStateResetDescriptors(formattingModel: powerbi.visuals.FormattingModel): void {
        const defaultNodesCard = formattingModel.cards.find(
            (card: powerbi.visuals.FormattingCard | powerbi.visuals.FormattingCardPlaceholder): card is powerbi.visuals.FormattingCard =>
                "groups" in card && card.displayName === this.formattingSettings.defaultNodes.displayName,
        );

        if (!defaultNodesCard) {
            return;
        }

        defaultNodesCard.revertToDefaultDescriptors = [
            ...(defaultNodesCard.revertToDefaultDescriptors || []),
            ...this.getPersistedStateResetDescriptors(),
        ];
    }

    private getPersistedStateResetDescriptors(): powerbi.visuals.FormattingDescriptor[] {
        return [
            "selectedNodeName",
            "selectedLinkName",
            "nodeStyleMap",
            "linkStyleMap",
            "nodePositionMap",
            "nodeGlobalStyle",
            "nodeFormatSyncState",
            "linkFormatSyncState",
        ].map((propertyName: string) => ({
            objectName: "persistedState",
            propertyName,
        }));
    }
}
