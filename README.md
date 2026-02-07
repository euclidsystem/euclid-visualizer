# Euclid Visualizer

Euclid Visualizer is a Microsoft Excel Office Add-in (task pane) built for the Euclid System. It lets users inspect the complete derivation trace behind financial metrics directly inside Excel. When a calculation graph is loaded, the add-in recursively resolves `$ref` links across multiple JSON files and renders the full computation DAG as a collapsible tree view. Mathematical expressions embedded in the graph — such as formulas and set predicates — are rendered with KaTeX. The add-in also provides a simple cell content viewer for quick inspection of selected cells.
