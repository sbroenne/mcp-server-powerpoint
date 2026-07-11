---
template: home.html
title: Home
description: >-
  Control Microsoft PowerPoint with natural language through AI assistants
  like GitHub Copilot and Claude. Automate slides, shapes, tables, charts and
  export-to-verify visual checks — no PowerPoint object-model knowledge
  required.
keywords: "PowerPoint automation, MCP server, AI PowerPoint, slide automation, PowerPoint COM, GitHub Copilot PowerPoint, Claude PowerPoint, PowerPoint CLI, presentation automation"
hide:
  - navigation
  - toc
---

!!! success "Live COM automation, not file parsing"
    Most PowerPoint MCP servers manipulate `.pptx` files offline with libraries
    like `python-pptx`, or use agent-run scripts with LibreOffice-rendered
    thumbnails. This project instead drives a **live, real PowerPoint desktop
    instance** via `Microsoft.Office.Interop.PowerPoint` — the official Primary
    Interop Assembly.

    - **True-fidelity rendering.** PowerPoint itself renders and saves the
      file, so there's zero risk of producing a `.pptx` that PowerPoint can't
      open.
    - **Export-to-verify.** After any visual edit, export the slide (or the
      whole deck) to an image with `export_slide_to_image` /
      `export_all_slides_to_images` and let a vision-capable AI assistant
      *see* the result — catching overlapping shapes, text overflow, and
      layout regressions that text-only automation simply cannot detect.

!!! tip "Also automating spreadsheets?"
    Check out [Excel MCP Server](https://excelmcpserver.dev/) — the sister
    project, built the same way.

## Key features

<div class="grid cards" markdown>

-   :material-view-carousel:{ .lg .middle } __Slides &amp; layouts__

    ---

    Add and delete slides, apply and inspect layouts, and query slide count
    for state discovery.

-   :material-shape-rectangle-plus:{ .lg .middle } __Shapes &amp; text__

    ---

    Add rectangles and text boxes, position and resize shapes, set and read
    rich text with font size, bold and color.

-   :material-table:{ .lg .middle } __Tables &amp; charts__

    ---

    Build tables cell-by-cell and add charts with real data — then read the
    data back to verify.

-   :material-note-text:{ .lg .middle } __Speaker notes__

    ---

    Set and read presenter notes per slide for talk-track generation and
    review.

-   :material-image:{ .lg .middle } __Images__

    ---

    Insert pictures from local files directly onto any slide.

-   :material-palette-swatch:{ .lg .middle } __Templates &amp; themes__

    ---

    Apply a `.potx`/`.pptx` template's masters, theme and layouts while
    preserving existing slide content, and read back the current theme name.

-   :material-format-font:{ .lg .middle } __Slide masters__

    ---

    Set title/body placeholder fonts and background color on the slide
    master — one edit, applied to every slide that inherits it.

-   :material-motion-play:{ .lg .middle } __Animations &amp; transitions__

    ---

    Add entrance/emphasis/exit effects to shapes and set slide transitions,
    then read them back to verify.

-   :material-image-check:{ .lg .middle } __Export-to-verify__

    ---

    Export any slide — or the whole deck — to images for multimodal visual
    verification. The project's core differentiator over text-only
    PowerPoint tooling.

</div>

[See all 18 tools (~98 operations) across 12 domains :material-arrow-right:](features.md){ .md-button .md-button--primary }

## See it in action

Ask your AI assistant in plain language — it drives PowerPoint for you:

!!! example "📝 Build a deck from scratch"
    **You:** "Create a new presentation with a title slide and three content
    slides about our Q3 results, then export it as images so I can see it."

    AI creates the presentation, adds slides with headings and body text, and
    exports PNGs of every slide to verify the result.

!!! example "📊 Tables &amp; charts"
    **You:** "Add a 4x3 table summarizing this data, then add a bar chart next
    to it."

    AI builds the table cell-by-cell and adds a chart shape with the given
    data, then exports the slide to confirm the layout looks right.

!!! example "🎨 Formatting &amp; shapes"
    **You:** "Make the title bold and blue, and move the logo to the
    top-right corner."

    AI applies text formatting through the TextFrame tools and repositions
    the shape, then exports an image to verify nothing overlaps.

!!! example "🗣️ Speaker notes"
    **You:** "Write speaker notes for each slide summarizing the key talking
    point."

    AI reads each slide's content and writes tailored notes via
    `set_notes_text`.

!!! example "🖼️ Visual verification"
    **You:** "Export slide 3 as an image and tell me if the chart overlaps
    the text box."

    AI exports the slide with `export_slide_to_image` and inspects the
    rendered PNG directly — catching issues no text-only tool could see.

## CLI or MCP Server?

This project provides both a **CLI** and an **MCP Server** interface. Choose
based on your use case:

| Interface | Best for | Why |
|-----------|----------|-----|
| **CLI** (`pptcli`) | Coding agents (Copilot, Cursor, Windsurf) | Single tool, no large schemas — better for cost-sensitive, high-throughput automation. |
| **MCP Server** (`mcp-powerpoint`) | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent session. Better for interactive, exploratory workflows. |

[MCP Server docs](mcp-server.md){ .md-button } [CLI docs](cli.md){ .md-button }

## GitHub star history

![GitHub stars over time for PowerPointMcp](assets/images/star-history.svg){ loading=lazy }

Updated daily from GitHub's stargazer data.
