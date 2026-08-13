# Sankey Flow

Interactive Power BI Sankey visual focused on a clean interaction model and AppSource-ready implementation.

## Current implementation

- Renders Sankey flows from `Source->Target Pair(s)` and `Values`.
- Supports either:
  - one `pairs` field with inline values such as `A -> B`, `A -> B, C`, and multiple entries split by `;` or new lines
  - multiple `pairs` fields to build layered flows across columns
- Persists node positions after drag and reapplies them on refresh.
- Supports Power BI cross-visual selection through `ISelectionManager` and row-backed `ISelectionId`s.
- Supports Power BI context menus on both data points and empty space.
- Uses the Power BI host tooltip service for nodes and links.
- Supports keyboard focus for nodes and links with `Enter`/`Space` activation and keyboard context menus.
- Shows a landing page when no data fields are bound.
- Shows inline warnings when rows are skipped because values are invalid or pair strings cannot be parsed.
- Requests up to 30,000 rows, aggregates repeated flows, and renders the top 1,000 aggregated flows when necessary to protect report performance.
- Uses the semantic model's measure and category format strings and the Power BI host locale.

## Data roles

- `Source->Target Pair(s)`: one or more categorical fields
- `Values`: one numeric measure

Rows with non-positive values are ignored. For a single pair field, use values like:

```text
A -> B
A -> B, C
A -> B; B -> D
```

## Format pane

### Layout

- `Node width`
- `Node spacing`

### Node Labels

- `Show`
- `Show text`
- `Show value`
- `Font size`
- `Color`

### Edge Labels

- `Show`
- `Show text`
- `Show value`
- `Font size`
- `Color`

### Value Format

- `Show as percentage`
- `Percentage decimal places` (shown when `Show as percentage` is enabled)

When percentage display is disabled, labels, tooltips, and accessibility text use the Values measure's semantic-model format string.

### Node Format

- `Editing`
- `Use report palette`
- `Palette`
- `Color`

Behavior:

- When `Use report palette` is `On`, nodes use the Power BI report palette. Selecting a node exposes a color override for that node.
- When `Use report palette` is `Off` and `Palette` is `Single color`, the `Color` picker controls the global node color.
- When `Use report palette` is `Off` and `Palette` is a preset palette, the visual cycles preset colors across nodes.
- Preset palettes currently available:
  - `Tableau`
  - `Okabe-Ito`
  - `Soft modern`
  - `Muted executive`
  - `IBM Carbon`

## Interaction behavior

- Click a node to select its contributing data rows in Power BI.
- Click a link to select the rows contributing to that connection.
- Use `Ctrl` or `Cmd` while clicking for multi-select.
- Click empty space to clear selection.
- Right-click a node or link to open the Power BI data-point context menu.
- Right-click empty space to open the general Power BI context menu.
- Drag nodes to reposition them inside the plot area.
- Hover or focus nodes and links to see the Power BI host tooltip.
- Press `Tab` to move across links and nodes, `Enter` or `Space` to select, and `Shift+F10` or the context-menu key to open the context menu.

## Styling notes

- Links inherit the source-node color by default.
- Link opacity currently stays on the internal fixed default; there is no visible link opacity control in the format pane.
- Node opacity is still stored internally for compatibility, but the visible node opacity control has been removed.
- The old outer background pane and drag-help overlay message were removed.
- There is no separate visible link formatting card at the moment.

## Development

```bash
npm install
npm start
npm run lint
npm run package
```

`npm run package` runs `pbiviz package --certification-audit`.

## Publication and certification status

The codebase now includes several certification-readiness features:

- `apiVersion` `5.11.0`
- locally pinned `powerbi-visuals-tools` `7.2.1`
- `powerbi-visuals-api` package `5.11.1`
- no declared privileges
- Power BI selection manager integration
- empty-space and data-point context menus
- cross-visual selection support
- Power BI host tooltip integration
- keyboard-focus support through `supportsKeyboardFocus`
- landing-page support through `supportsLandingPage` and `supportsEmptyDataView`
- clean moderate-and-higher `npm audit` result through the current dependency overrides
- successful `pbiviz package --certification-audit`

Still required before Microsoft submission:

- replace the `author.email` value in `pbiviz.json` with a monitored support mailbox before submission
- prepare AppSource submission assets
  - offline sample `.pbix`
  - screenshots
  - privacy policy
  - EULA
  - support site
  - 300x300 marketplace logo
- address remaining recommended submission gaps if targeting a stronger review outcome
  - allow interactions
  - highlight data
  - localization
