---
title: Apply a Template Without Replacing Content
description: Apply PowerPoint template masters, themes, and layouts to an open deck while preserving its slide content.
---

# Apply a Template Without Replacing Content

Use `apply-template` when a deck has the right content but needs a different theme or master.

## 1. Open and inspect the deck

Open the presentation and read its current theme with
`presentation(action="get-theme-name", ...)`. Render a few representative slides before making
the change so you have a visual baseline.

## 2. Apply the template

```text
presentation(
  action="apply-template",
  sessionId="...",
  templatePath="C:\Templates\brand.potx"
)
```

The template source may be a PowerPoint template or presentation file. Its masters, theme, and
layouts are applied while existing slide content is preserved.

## 3. Check the result

Read the theme name again, inspect slide layouts, and export all slides to images. Template changes
can alter fonts, placeholder sizes, and line wrapping, so every slide needs visual review.

## 4. Save only after review

Close with `save=true` after the rendered deck is acceptable. Close with `save=false` to discard
the template change.
