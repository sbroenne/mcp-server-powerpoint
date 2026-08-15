# PowerPoint Automation Guides

Task-focused walkthroughs for driving PowerPoint from an AI assistant or a script.

Each guide starts from a real goal — "build me a deck", "update last quarter's slides" — and
walks the whole loop, including the visual verification step that separates this approach from
offline `.pptx` libraries.

## Guides

| Guide | Start here if you want to... |
|-------|------------------------------|
| [Build a deck with AI](BUILD-A-DECK-WITH-AI.md) | Generate a complete multi-slide presentation from a prompt or an outline |
| [Edit an existing deck](EDIT-AN-EXISTING-DECK.md) | Update, restructure or fix a `.pptx` file you already have |
| [Export slides to images](EXPORT-SLIDES-TO-IMAGES.md) | Render slides to PNG so an assistant can see and correct its own output |
| [Automate charts and tables](AUTOMATE-CHARTS-AND-TABLES.md) | Put real data on a slide as a native chart or table |
| [Apply a corporate template](APPLY-A-CORPORATE-TEMPLATE.md) | Force generated decks onto your brand's template, theme and fonts |

## Before you start

All five guides assume you have the MCP server or the CLI installed — see
[Installation](https://powerpointmcpserver.dev/installation/). They also assume Windows with the
PowerPoint desktop application, because every operation runs against a live PowerPoint instance
over COM.

Two things are true of every guide here:

- **Indexes are 1-based.** The first slide is `slide_index: 1`, matching PowerPoint's own object
  model. There is no slide `0`.
- **Sessions are reused.** Open or create a presentation once, keep the returned session id, and
  pass it to every subsequent call. Reopening the same file for each operation is the single most
  common mistake, and it is slow — launching PowerPoint costs far more than the work itself.

## Going deeper

The [reference section](https://powerpointmcpserver.dev/reference/) publishes the complete expert
corpus that ships inside the agent skill packages — every tool, action and parameter, plus the
behavioral rules and anti-patterns an assistant is expected to follow.
