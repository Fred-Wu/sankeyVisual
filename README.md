# Certified Sankey Visual

Custom Power BI Sankey visual focused on a clean interaction model and AppSource/certification-oriented implementation.

## Current implementation

- Renders Sankey flows from `Source->Target Pair(s)` and `Values`.
- Supports either:
  - one `pairs` field with inline values such as `A -> B`, `A -> B, C`, and multiple entries split by `;` or new lines
  - multiple `pairs` fields to build layered flows across columns
- Persists node positions after drag and reapplies them on refresh.
- Supports Power BI cross-visual selection through `ISelectionManager` and row-backed `ISelectionId`s.
- Supports Power BI context menus on both data points and empty space.
- Shows inline warnings when rows are skipped because values are invalid or pair strings cannot be parsed.

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
- `Decimal places`

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
- Hover nodes and links to see the visual's custom tooltip.

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

The codebase now includes several certification-oriented pieces:

- `apiVersion` `5.11.0`
- no declared privileges
- Power BI selection manager integration
- empty-space and data-point context menus
- cross-visual selection support
- clean `npm audit` result through the current dependency overrides

Still required before Microsoft submission:

- replace placeholder metadata in `pbiviz.json`
  - `supportUrl`
  - `gitHubUrl`
  - author support email
- publish the real source repository and maintain the `certification` branch
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
  - keyboard navigation
  - landing page
  - localization
  - official tooltip integration
