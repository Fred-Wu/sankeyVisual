import powerbi from "powerbi-visuals-api";

import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewTable = powerbi.DataViewTable;
import IColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;
import ISelectionId = powerbi.visuals.ISelectionId;
import ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;

export interface PersistedStyle {
    color: string;
    opacity: number;
}

export interface StyleMap {
    [key: string]: PersistedStyle;
}

export interface AppliedStyleContext {
    defaultNodeColor: string;
    defaultNodeOpacity: number;
    useThemePalette: boolean;
    nodePaletteColors?: string[];
    defaultLinkColor: string;
    defaultLinkOpacity: number;
    inheritLinkColor: boolean;
    nodeStyles: StyleMap;
    linkStyles: StyleMap;
    isHighContrast: boolean;
    highContrastForeground: string;
    highContrastBackground: string;
}

export interface SankeyNodeDatum {
    id: string;
    label: string;
    layerIndex: number;
    color: string;
    opacity: number;
    incomingValue: number;
    outgoingValue: number;
    selectionIds: ISelectionId[];
    contextMenuSelectionId?: ISelectionId;
}

export interface SankeyLinkDatum {
    id: string;
    key: string;
    label: string;
    source: string;
    target: string;
    value: number;
    color: string;
    opacity: number;
    selectionIds: ISelectionId[];
    contextMenuSelectionId?: ISelectionId;
}

export interface SankeyVisualData {
    nodes: SankeyNodeDatum[];
    links: SankeyLinkDatum[];
    warnings: string[];
}

export type SelectionIdBuilderFactory = () => ISelectionIdBuilder;

interface PendingNode {
    label: string;
    layerIndex: number;
    incomingValue: number;
    outgoingValue: number;
    selectionIds: Map<string, ISelectionId>;
}

interface PendingLink {
    key: string;
    source: string;
    target: string;
    value: number;
    selectionIds: Map<string, ISelectionId>;
}

interface SankeyAccumulator {
    nodeOrder: string[];
    nodes: Map<string, PendingNode>;
    links: Map<string, PendingLink>;
    warnings: string[];
}

const supportedArrowTokens: string[] = ["->", "=>", "→"];

function parsePairCell(input: string): Array<{ source: string; target: string }> {
    const parts = input
        .split(/[\r\n;]+/)
        .map((value: string) => value.trim())
        .filter(Boolean);

    const pairs: Array<{ source: string; target: string }> = [];

    parts.forEach((part: string) => {
        const arrowToken = supportedArrowTokens.find((token: string) => part.includes(token));
        if (!arrowToken) {
            return;
        }

        const [rawSource, rawTargets] = part.split(arrowToken, 2);
        const source = rawSource?.trim();

        if (!source || !rawTargets) {
            return;
        }

        rawTargets
            .split(/[|,]+/)
            .map((value: string) => value.trim())
            .filter(Boolean)
            .forEach((target: string) => {
                if (target !== source) {
                    pairs.push({ source, target });
                }
            });
    });

    return pairs;
}

function buildNodeId(label: string, layerIndex: number): string {
    return `${label} (L${layerIndex + 1})`;
}

function addNodeIfMissing(
    accumulator: SankeyAccumulator,
    nodeId: string,
    label: string,
    layerIndex: number,
): void {
    if (!accumulator.nodes.has(nodeId)) {
        accumulator.nodes.set(nodeId, {
            label,
            layerIndex,
            incomingValue: 0,
            outgoingValue: 0,
            selectionIds: new Map<string, ISelectionId>(),
        });
        accumulator.nodeOrder.push(nodeId);
    }
}

function createAccumulator(): SankeyAccumulator {
    return {
        nodeOrder: [],
        nodes: new Map<string, PendingNode>(),
        links: new Map<string, PendingLink>(),
        warnings: [],
    };
}

function addSelectionId(selectionIds: Map<string, ISelectionId>, selectionId?: ISelectionId): void {
    if (!selectionId) {
        return;
    }

    selectionIds.set(selectionId.getKey(), selectionId);
}

function getSelectionArray(selectionIds: Map<string, ISelectionId>): ISelectionId[] {
    return Array.from(selectionIds.values());
}

function buildCategoricalRowSelectionId(
    selectionIdBuilderFactory: SelectionIdBuilderFactory | undefined,
    categoryColumns: DataViewCategoryColumn[],
    rowIndex: number,
): ISelectionId | undefined {
    if (!selectionIdBuilderFactory) {
        return undefined;
    }

    const availableColumns = categoryColumns.filter((column: DataViewCategoryColumn) => rowIndex < column.values.length);
    if (availableColumns.length === 0) {
        return undefined;
    }

    const selectionIdBuilder = selectionIdBuilderFactory();
    availableColumns.forEach((column: DataViewCategoryColumn) => {
        selectionIdBuilder.withCategory(column, rowIndex);
    });

    return selectionIdBuilder.createSelectionId();
}

function buildTableRowSelectionId(
    selectionIdBuilderFactory: SelectionIdBuilderFactory | undefined,
    table: DataViewTable,
    rowIndex: number,
): ISelectionId | undefined {
    if (!selectionIdBuilderFactory) {
        return undefined;
    }

    return selectionIdBuilderFactory()
        .withTable(table, rowIndex)
        .createSelectionId();
}

function addLink(
    accumulator: SankeyAccumulator,
    sourceId: string,
    sourceLabel: string,
    sourceLayerIndex: number,
    targetId: string,
    targetLabel: string,
    targetLayerIndex: number,
    value: number,
    selectionId?: ISelectionId,
): void {
    if (!sourceId || !targetId || sourceId === targetId) {
        return;
    }

    addNodeIfMissing(accumulator, sourceId, sourceLabel, sourceLayerIndex);
    addNodeIfMissing(accumulator, targetId, targetLabel, targetLayerIndex);

    const sourceNode = accumulator.nodes.get(sourceId);
    const targetNode = accumulator.nodes.get(targetId);
    if (!sourceNode || !targetNode) {
        return;
    }

    addSelectionId(sourceNode.selectionIds, selectionId);
    addSelectionId(targetNode.selectionIds, selectionId);

    const linkKey = `${sourceId}|||${targetId}`;
    const currentLink = accumulator.links.get(linkKey);
    if (currentLink) {
        currentLink.value += value;
        addSelectionId(currentLink.selectionIds, selectionId);
    } else {
        const selectionIds = new Map<string, ISelectionId>();
        addSelectionId(selectionIds, selectionId);
        accumulator.links.set(linkKey, {
            key: linkKey,
            source: sourceId,
            target: targetId,
            value,
            selectionIds,
        });
    }

    sourceNode.outgoingValue += value;
    targetNode.incomingValue += value;
}

function addInlinePairRow(
    accumulator: SankeyAccumulator,
    pairText: string,
    value: number,
    rowIndex: number,
    selectionId?: ISelectionId,
): void {
    const trimmedText = pairText.trim();
    if (!trimmedText) {
        accumulator.warnings.push(`Row ${rowIndex + 1} has an empty Source->Target Pair(s) value and was skipped.`);
        return;
    }

    const parsedPairs = parsePairCell(trimmedText);
    if (parsedPairs.length === 0) {
        accumulator.warnings.push(`Row ${rowIndex + 1} could not be parsed. Use formats like "A -> B" or "A -> B, C".`);
        return;
    }

    parsedPairs.forEach(({ source, target }) => addLink(
        accumulator,
        buildNodeId(source, 0),
        source,
        0,
        buildNodeId(target, 1),
        target,
        1,
        value,
        selectionId,
    ));
}

function addLayeredPairRow(
    accumulator: SankeyAccumulator,
    pairValues: string[],
    value: number,
    rowIndex: number,
    selectionId?: ISelectionId,
): void {
    let addedEdge = false;

    for (let index = 0; index < pairValues.length - 1; index++) {
        const source = pairValues[index]?.trim() || "";
        const target = pairValues[index + 1]?.trim() || "";

        if (!source || !target || source === target) {
            continue;
        }

        addLink(
            accumulator,
            buildNodeId(source, index),
            source,
            index,
            buildNodeId(target, index + 1),
            target,
            index + 1,
            value,
            selectionId,
        );
        addedEdge = true;
    }

    if (!addedEdge) {
        accumulator.warnings.push(`Row ${rowIndex + 1} needs at least two populated category values in Source->Target Pair(s).`);
    }
}

function getNodeBaseColor(nodeId: string, paletteColor: string, styles: AppliedStyleContext): string {
    if (styles.isHighContrast) {
        return styles.highContrastBackground;
    }

    if (styles.nodeStyles[nodeId]) {
        return styles.nodeStyles[nodeId].color;
    }

    if (styles.useThemePalette || styles.nodePaletteColors?.length) {
        return paletteColor;
    }

    return styles.defaultNodeColor;
}

function getNodeOpacity(nodeId: string, styles: AppliedStyleContext): number {
    if (styles.isHighContrast) {
        return 1;
    }

    if (styles.nodeStyles[nodeId]) {
        return clampOpacity(styles.nodeStyles[nodeId].opacity);
    }

    return clampOpacity(styles.defaultNodeOpacity);
}

function getLinkColor(link: PendingLink, nodeColors: Map<string, string>, styles: AppliedStyleContext): string {
    if (styles.isHighContrast) {
        return styles.highContrastForeground;
    }

    if (styles.inheritLinkColor && styles.linkStyles[link.key]) {
        return styles.linkStyles[link.key].color;
    }

    if (styles.inheritLinkColor) {
        return nodeColors.get(link.source) || styles.defaultLinkColor;
    }

    return styles.defaultLinkColor;
}

function getLinkOpacity(linkKey: string, styles: AppliedStyleContext): number {
    if (styles.isHighContrast) {
        return 0.9;
    }

    if (styles.inheritLinkColor && styles.linkStyles[linkKey]) {
        return clampOpacity(styles.linkStyles[linkKey].opacity);
    }

    return clampOpacity(styles.defaultLinkOpacity);
}

export function clampOpacity(value: number): number {
    if (!Number.isFinite(value)) {
        return 1;
    }

    return Math.max(0, Math.min(100, value)) / 100;
}

export function parseStyleMap(jsonValue: string | undefined): StyleMap {
    if (!jsonValue) {
        return {};
    }

    try {
        const parsed = JSON.parse(jsonValue) as Record<string, Partial<PersistedStyle>>;
        const styleMap: StyleMap = {};

        Object.keys(parsed || {}).forEach((key: string) => {
            const style = parsed[key];
            const color = typeof style?.color === "string" && style.color.trim().length > 0
                ? style.color.trim()
                : undefined;
            const opacity = typeof style?.opacity === "number" && Number.isFinite(style.opacity)
                ? Math.max(0, Math.min(100, style.opacity))
                : undefined;

            if (color && opacity !== undefined) {
                styleMap[key] = { color, opacity };
            }
        });

        return styleMap;
    } catch {
        return {};
    }
}

export function serializeStyleMap(styleMap: StyleMap): string {
    const sortedKeys = Object.keys(styleMap).sort((left: string, right: string) => left.localeCompare(right));
    const normalized: StyleMap = {};

    sortedKeys.forEach((key: string) => {
        normalized[key] = {
            color: styleMap[key].color,
            opacity: Math.round(Math.max(0, Math.min(100, styleMap[key].opacity))),
        };
    });

    return JSON.stringify(normalized);
}

function finalizeSankeyData(accumulator: SankeyAccumulator, palette: IColorPalette, styles: AppliedStyleContext): SankeyVisualData {
    const paletteColors = new Map<string, string>();
    if (styles.useThemePalette && !styles.isHighContrast) {
        accumulator.nodeOrder.forEach((nodeId: string) => {
            paletteColors.set(nodeId, palette.getColor(nodeId).value);
        });
    } else if (!styles.isHighContrast && styles.nodePaletteColors?.length) {
        accumulator.nodeOrder.forEach((nodeId: string, index: number) => {
            paletteColors.set(nodeId, styles.nodePaletteColors![index % styles.nodePaletteColors!.length]);
        });
    }

    const nodeColors = new Map<string, string>();
    const nodes: SankeyNodeDatum[] = accumulator.nodeOrder.map((nodeId: string) => {
        const nodeInfo = accumulator.nodes.get(nodeId) || {
            label: nodeId,
            layerIndex: 0,
            incomingValue: 0,
            outgoingValue: 0,
            selectionIds: new Map<string, ISelectionId>(),
        };
        const paletteColor = paletteColors.get(nodeId) || styles.defaultNodeColor;
        const color = getNodeBaseColor(nodeId, paletteColor, styles);
        const selectionIds = getSelectionArray(nodeInfo.selectionIds);
        nodeColors.set(nodeId, color);

        return {
            id: nodeId,
            label: nodeInfo.label,
            layerIndex: nodeInfo.layerIndex,
            color,
            opacity: getNodeOpacity(nodeId, styles),
            incomingValue: nodeInfo.incomingValue,
            outgoingValue: nodeInfo.outgoingValue,
            selectionIds,
            contextMenuSelectionId: selectionIds[0],
        };
    });

    const links: SankeyLinkDatum[] = Array.from(accumulator.links.values()).map((link: PendingLink) => {
        const sourceNode = accumulator.nodes.get(link.source);
        const targetNode = accumulator.nodes.get(link.target);
        const selectionIds = getSelectionArray(link.selectionIds);

        return {
            id: link.key,
            key: link.key,
            label: `${sourceNode?.label || link.source} (L${(sourceNode?.layerIndex ?? 0) + 1}) -> ${targetNode?.label || link.target} (L${(targetNode?.layerIndex ?? 0) + 1})`,
            source: link.source,
            target: link.target,
            value: link.value,
            color: getLinkColor(link, nodeColors, styles),
            opacity: getLinkOpacity(link.key, styles),
            selectionIds,
            contextMenuSelectionId: selectionIds[0],
        };
    });

    return {
        nodes,
        links,
        warnings: accumulator.warnings,
    };
}

function buildSankeyDataFromTable(
    table: DataViewTable | undefined,
    palette: IColorPalette,
    styles: AppliedStyleContext,
    selectionIdBuilderFactory?: SelectionIdBuilderFactory,
): SankeyVisualData | undefined {
    if (!table?.columns?.length) {
        return undefined;
    }

    const pairIndexes: number[] = [];
    let valueIndex: number | undefined;

    table.columns.forEach((column: powerbi.DataViewMetadataColumn, index: number) => {
        if (column.roles?.pairs) {
            pairIndexes.push(index);
        }

        if (column.roles?.values && valueIndex === undefined) {
            valueIndex = index;
        }
    });

    if (pairIndexes.length === 0 || valueIndex === undefined) {
        return undefined;
    }

    const accumulator = createAccumulator();

    table.rows?.forEach((row: powerbi.DataViewTableRow, rowIndex: number) => {
        const rawValue = Number(row[valueIndex]);
        if (!Number.isFinite(rawValue) || rawValue <= 0) {
            accumulator.warnings.push(`Row ${rowIndex + 1} has a non-positive Values entry and was skipped.`);
            return;
        }

        const selectionId = buildTableRowSelectionId(selectionIdBuilderFactory, table, rowIndex);

        if (pairIndexes.length === 1) {
            addInlinePairRow(accumulator, String(row[pairIndexes[0]] ?? ""), rawValue, rowIndex, selectionId);
            return;
        }

        addLayeredPairRow(
            accumulator,
            pairIndexes.map((index: number) => String(row[index] ?? "")),
            rawValue,
            rowIndex,
            selectionId,
        );
    });

    return finalizeSankeyData(accumulator, palette, styles);
}

function buildSankeyDataFromCategorical(
    dataView: DataView | undefined,
    palette: IColorPalette,
    styles: AppliedStyleContext,
    selectionIdBuilderFactory?: SelectionIdBuilderFactory,
): SankeyVisualData | undefined {
    const categoryColumns: DataViewCategoryColumn[] = (dataView?.categorical?.categories || [])
        .filter((column: DataViewCategoryColumn) => column.source.roles?.pairs);
    const valueColumn = (dataView?.categorical?.values || [])
        .find((column: powerbi.DataViewValueColumn) => column.source.roles?.values)
        || dataView?.categorical?.values?.[0];

    if (categoryColumns.length === 0 || !valueColumn) {
        return undefined;
    }

    const rowCount = Math.max(...categoryColumns.map((column: DataViewCategoryColumn) => column.values.length), valueColumn.values?.length || 0);
    const accumulator = createAccumulator();

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
        const rawValue = Number(valueColumn.values?.[rowIndex]);
        if (!Number.isFinite(rawValue) || rawValue <= 0) {
            accumulator.warnings.push(`Row ${rowIndex + 1} has a non-positive Values entry and was skipped.`);
            continue;
        }

        const selectionId = buildCategoricalRowSelectionId(selectionIdBuilderFactory, categoryColumns, rowIndex);

        if (categoryColumns.length === 1) {
            addInlinePairRow(accumulator, String(categoryColumns[0].values[rowIndex] ?? ""), rawValue, rowIndex, selectionId);
            continue;
        }

        addLayeredPairRow(
            accumulator,
            categoryColumns.map((column: DataViewCategoryColumn) => String(column.values[rowIndex] ?? "")),
            rawValue,
            rowIndex,
            selectionId,
        );
    }

    return finalizeSankeyData(accumulator, palette, styles);
}

export function buildSankeyData(
    dataView: DataView | undefined,
    palette: IColorPalette,
    styles: AppliedStyleContext,
    selectionIdBuilderFactory?: SelectionIdBuilderFactory,
): SankeyVisualData {
    const categoricalData = buildSankeyDataFromCategorical(dataView, palette, styles, selectionIdBuilderFactory);
    if (categoricalData) {
        return categoricalData;
    }

    const tableData = buildSankeyDataFromTable(dataView?.table, palette, styles, selectionIdBuilderFactory);
    if (tableData) {
        return tableData;
    }

    const categoryColumn: DataViewCategoryColumn | undefined = dataView?.categorical?.categories?.[0];
    const valueColumn = dataView?.categorical?.values?.[0];

    if (!categoryColumn || !valueColumn) {
        return {
            nodes: [],
            links: [],
            warnings: [],
        };
    }

    const accumulator = createAccumulator();

    categoryColumn.values.forEach((rawPair: powerbi.PrimitiveValue, index: number) => {
        const rawValue = Number(valueColumn.values?.[index]);

        if (!Number.isFinite(rawValue) || rawValue <= 0) {
            accumulator.warnings.push(`Row ${index + 1} has a non-positive Values entry and was skipped.`);
            return;
        }

        addInlinePairRow(
            accumulator,
            String(rawPair ?? ""),
            rawValue,
            index,
            buildCategoricalRowSelectionId(selectionIdBuilderFactory, [categoryColumn], index),
        );
    });

    return finalizeSankeyData(accumulator, palette, styles);
}
