# Sankey Flow Certification Next Steps

The current repo packages successfully with `pbiviz package --certification-audit`.

Code-side items now covered:

- schema-clean `pbiviz.json` for `externalJS`
- Power BI host tooltips
- keyboard focus support for Sankey links and nodes
- landing page support when no fields are bound
- locally pinned current Power BI visual tooling
- clean moderate-and-higher dependency audit
- model-aware number and category formatting
- explicit 30,000-row data request with bounded SVG rendering

Remaining code/product gaps:

- `Allow Interactions`
- `Highlight Data`
- `Localizations`

Remaining submission assets and Partner Center work:

- replace the GitHub `noreply` author email in `pbiviz.json` with a monitored support address
- prepare an offline sample `.pbix` that demonstrates:
  - single-column `A -> B` parsing
  - multi-column layered flows
  - positive, zero, and invalid values
  - cross-visual selection
  - context menu behavior
  - keyboard navigation
  - high-contrast rendering
- prepare AppSource screenshots
- prepare privacy policy and EULA URLs
- prepare a support/help page
- prepare a 300x300 marketplace logo

Recommended validation before submission:

- run `npm run lint`
- run `npm run package -- --verbose`
- verify no browser-console errors in Power BI Desktop
- verify tooltip, selection, context menu, drag, resize, and keyboard behavior in Desktop and Service
- verify no web access or other unexpected certification privileges are introduced
