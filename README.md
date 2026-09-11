# Sankey Flow

A custom Power BI visual for showing how values flow between categories. The width of each connection represents its value, making it easy to compare flows such as spending, customer journeys, or process stages.

## Get started

1. Import the `.pbiviz` file into your Power BI report.
2. Add **Sankey Flow** to the report canvas.
3. Add your category fields to **Source->Target Pair(s)** and a numeric measure to **Values**.

## Prepare your data

The visual uses **distinct combinations of the selected group columns**, with **one aggregated numeric value for each combination**:

| Region | Channel | Product | Amount |
| --- | --- | --- | ---: |
| East | Online | Product A | 80 |
| West | Retail | Product B | 20 |
| East | Retail | Product B | 40 |

In the visual's field wells:

- **Source->Target Pair(s):** Add the group columns in stage order: **Region**, **Channel**, then **Product**. Use at least two group columns.
- **Values:** Add one numeric measure, such as **Sum of Amount**.

Power BI groups the data by those columns and evaluates the measure for each distinct combination. The visual then creates connections between adjacent groups. For the first row, it creates **East → Online** and **Online → Product A**, each with a value of **80**.

Use one numeric measure in **Values**. Zero, negative, and invalid values are skipped. Repeated connections are summed. Blank stage values omit their adjacent connections.

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
