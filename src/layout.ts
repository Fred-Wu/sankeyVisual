interface LayeredNode {
    layerIndex: number;
    depth?: number;
    height?: number;
}

export function alignByInputOrder(node: LayeredNode, columnCount: number): number {
    // Missing categories or pruned links can leave fewer layout columns than
    // input fields. Reserve enough columns for every ancestor and descendant
    // so a surviving flow cannot collapse into a single column.
    const firstColumn = node.depth ?? 0;
    const lastColumn = columnCount - 1 - (node.height ?? 0);
    return Math.max(firstColumn, Math.min(lastColumn, node.layerIndex));
}
