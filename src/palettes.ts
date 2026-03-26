export const defaultNodePaletteMode = "singleColor" as const;

export const nodePaletteColors = {
    tableau: ["#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F", "#EDC948", "#B07AA1", "#FF9DA7"],
    okabeIto: ["#0072B2", "#E69F00", "#009E73", "#D55E00", "#CC79A7", "#56B4E9", "#F0E442", "#000000"],
    softModern: ["#5B8FF9", "#61DDAA", "#65789B", "#F6BD16", "#7262FD", "#78D3F8", "#9661BC", "#F6903D"],
    mutedExecutive: ["#355070", "#6D597A", "#B56576", "#E56B6F", "#EAAC8B", "#99C1B9", "#6B9080", "#A4C3B2"],
    ibmCarbon: ["#0F62FE", "#198038", "#8A3FFC", "#FF832B", "#DA1E28", "#009D9A", "#6F6F6F", "#B28600"],
} as const;

export type CustomNodePaletteMode = keyof typeof nodePaletteColors;
export type NodePaletteMode = typeof defaultNodePaletteMode | CustomNodePaletteMode;

export function isNodePaletteMode(value: string | undefined): value is NodePaletteMode {
    return value === defaultNodePaletteMode || value in nodePaletteColors;
}

export function getNodePaletteColors(mode: NodePaletteMode): string[] | undefined {
    if (mode === defaultNodePaletteMode) {
        return undefined;
    }

    return [...nodePaletteColors[mode]];
}
