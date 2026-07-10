---
"powerpointmcp": minor
---

Add `chart(action: "replace-chart-data", ...)` to wholesale-replace an existing chart's
categories, series names, and values in a single call, including changing the category count.
`series_values` is a flat, series-major array (all values for the first series, then the second,
etc.), avoiding the previous delete-shape-and-recreate workaround for editing chart data in place.
