# Sankey Flow

A custom Power BI visual for showing how values flow between categories. The width of each connection represents its value, making it easy to compare flows such as spending, customer journeys, or process stages.

## Get started

1. Import the `.pbiviz` file into your Power BI report.
2. Add **Sankey Flow** to the report canvas.
3. Add your category fields to **Source->Target Pair(s)** and a numeric measure to **Values**.

## Prepare your data

Use either a single field containing source-to-target pairs:

| Source->Target Pair(s) | Values |
| --- | ---: |
| Sales -> Operations | 80 |
| Sales -> Marketing | 20 |

Or use separate fields for successive stages. Add the category fields in the order you want them to appear:

| Region | Channel | Product | Values |
| --- | --- | --- | ---: |
| East | Online | Product A | 80 |
| West | Retail | Product B | 20 |

Repeated connections are combined. Values must be positive numbers. Blank categories omit their adjacent connections.

## Customize the visual

Use the Format pane to adjust node width and spacing, show or hide node and connection labels, change colors and palettes, and display values as percentages.

Values otherwise follow the measure's Power BI format, such as currency or decimal places.

## Explore your flows

- **Select:** Click a node or connection to filter other visuals. Hold Ctrl/Cmd to select more than one; click empty space to clear.
- **Inspect:** Hover over a node or connection to see its tooltip. Right-click for the Power BI context menu.
- **Arrange:** Drag nodes to reposition them. Positions are saved with the report.
- **Keyboard:** Use Tab to move between nodes and connections, Enter/Space to select, and Shift+F10 for the context menu.

For large datasets, the visual uses up to 30,000 rows and displays at most 1,000 aggregated connections. A warning appears when connections are omitted from the display.

## Support

Report issues or request features on [GitHub Issues](https://github.com/Fred-Wu/sankeyVisual/issues).
