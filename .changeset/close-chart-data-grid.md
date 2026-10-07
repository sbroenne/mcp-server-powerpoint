---
"powerpointmcp": patch
---

**Adding several charts in a row works again** (#108): chart actions (`add-chart`, `get-chart-data`, `add-series`, `replace-chart-data`) left PowerPoint's chart data window ("Chart in Microsoft PowerPoint" in Excel) open. The next `add-chart` then failed with "The chart data grid is already open". Each chart action now closes that window when it finishes, so the next chart can be added straight away and no stray Excel window is left on screen.
