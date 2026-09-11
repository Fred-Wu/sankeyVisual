const assert = require("node:assert/strict");
const { test } = require("node:test");
const { sankey, sankeyLinkHorizontal } = require("d3-sankey");
const { alignByInputOrder } = require("../src/layout.ts");

function renderLayout(nodes, links) {
    const graph = sankey()
        .nodeId(node => node.id)
        .nodeAlign(alignByInputOrder)
        .nodeSort(null)
        .linkSort(null)
        .extent([[80, 20], [500, 300]])({ nodes, links });

    for (const node of graph.nodes) {
        for (const coordinate of [node.x0, node.x1, node.y0, node.y1]) {
            assert.ok(Number.isFinite(coordinate), `${node.id} has invalid geometry`);
        }
    }
    for (const link of graph.links) {
        assert.ok(link.source.x1 < link.target.x0, "Flows must run left to right");
        assert.ok(Number.isFinite(link.width) && link.width > 0);
        assert.doesNotMatch(sankeyLinkHorizontal()(link), /NaN|Infinity/);
    }
    return graph;
}

test("renders B -> C after the first of three category fields is blank", () => {
    const graph = renderLayout(
        [{ id: "B", layerIndex: 1 }, { id: "C", layerIndex: 2 }],
        [{ source: "B", target: "C", value: 1 }],
    );
    assert.equal(graph.links[0].value, 1);
    assert.equal(graph.nodes[0].x0, 80);
    assert.equal(graph.nodes[1].x1, 500);
});

test("renders disconnected flows separated by blank category fields", () => {
    renderLayout(
        [
            { id: "A", layerIndex: 0 }, { id: "B", layerIndex: 1 },
            { id: "D", layerIndex: 3 }, { id: "E", layerIndex: 4 },
        ],
        [{ source: "A", target: "B", value: 2 }, { source: "D", target: "E", value: 3 }],
    );
});

test("preserves columns for a complete layered flow alongside a partial flow", () => {
    const graph = renderLayout(
        [
            { id: "A", layerIndex: 0 }, { id: "B", layerIndex: 1 },
            { id: "C", layerIndex: 2 }, { id: "X", layerIndex: 1 },
            { id: "Y", layerIndex: 2 },
        ],
        [
            { source: "A", target: "B", value: 2 },
            { source: "B", target: "C", value: 2 },
            { source: "X", target: "Y", value: 3 },
        ],
    );
    assert.equal(graph.nodes[1].x0, graph.nodes[3].x0);
    assert.equal(graph.nodes[2].x0, graph.nodes[4].x0);
});
