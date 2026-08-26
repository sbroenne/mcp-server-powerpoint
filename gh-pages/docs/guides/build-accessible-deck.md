---
title: Build an Accessible Presentation
description: Add useful alternative text, title slides clearly, set reading order, and run deterministic accessibility checks in PowerPoint.
---

# Build an Accessible Presentation

The accessibility tool performs deterministic checks. It does not replace human review, but it
catches repeatable problems before delivery.

## Give visual content useful alternative text

Use `shape(action="set-alt-text", ...)` for images, charts, diagrams, and other meaningful visual
content. Describe the information the object communicates, not its appearance alone.

Decorative objects should be handled consistently with the presentation's accessibility policy.

## Keep titles meaningful

Use a title placeholder when the layout provides one and avoid leaving it empty. The accessibility
audit reports empty title placeholders because slide titles help people navigate a deck.

## Set a logical reading order

1. Read the current order with `accessibility(action="get-reading-order", ...)`.
2. Arrange shapes in the order a listener should hear them.
3. Write the order with `accessibility(action="set-reading-order", ...)`.
4. Read it back to confirm the change.

## Audit and visually verify

Run `accessibility(action="audit", ...)` after content changes. Then export each slide to an image
and review contrast, font size, spacing, and information that depends only on color. Those visual
qualities are outside the deterministic audit.
