---
format: 1920x1080
duration: 60s
message: "Give AI a dependable MCP interface to the real PowerPoint desktop app, with an optional token-efficient CLI"
arc: Demo Loop
audience: developers and AI automation users
mode: collaborative
music: none
narration: warm confident female voice, with burned-in captions and WebVTT
---

## Video direction

Use the remixed Code Editorial system from `frame.md`: white editorial canvas,
black ink, PowerPoint orange as the single voltage accent, dark warm terminal
surfaces, Segoe UI display/body, and Cascadia Mono for commands and metadata.
The dominant film current is leftward. Ordinary seams use a velocity-matched
leftward push; zoom-through is reserved for the two section changes. Reveals
are sequential across each frame instead of arriving all at once, with smooth
long-tail settles and no bounce. Frames 1 and 7 are deliberate held reads.
Live PowerPoint footage performs the middle of the film; cards never breathe or
float merely to fill time. Avoid generic AI gradients, browser chrome, fake
PowerPoint UI, crossfades, front-loaded slideshow motion, and independent
screensaver-like drift.

## Frame 1 — PowerPoint, automated

- scene: The PowerPoint MCP name and orange mark fade in over a restrained dark field.
- voiceover: "PowerPoint is where ideas become decisions."
- duration: 5s
- poster: 3s
- transition_in: cut
- status: animated
- src: compositions/frames/01-powerpoint-automated.html
- type: hook
- persuasion: Outcome-first promise
- beat: confidence
- blueprint: titlecard-reveal
- asset_candidates:
- focal: typography-only
- roles: no media; the title and PowerPoint orange rule are the hero

narrativeRole: Open with the result, not the implementation, and establish the product identity.
keyMessage: PowerPoint automation can feel direct and dependable.

Reproduce: keep the single restrained title reveal and long held read.
Scene 1 (0.0–1.2s): black full frame with only a short PowerPoint-orange rule drawing left-to-right in the upper-left third — asymmetric editorial layout, low density, two depth layers; SVG self-draw (`svg-path-draw`) establishes the current.
Scene 2 (1.2–2.4s): “PowerPoint, automated.” arrives as two waterfall lines from the right, continuing leftward into position — oversized display hierarchy filling most of the upper 70%; waterfall entry (`waterfall-entry`) with a smooth long-tail settle.
Scene 3 (2.4–5.0s): the mono “MCP SERVER · OPTIONAL CLI” label reveals beneath the title, then the composition holds completely still for the read.

## Frame 2 — Start with MCP

- scene: MCP takes primary focus while `pptcli` appears as an optional token-efficient path for scripts and LLMs.
- voiceover: "PowerPoint MCP lets your AI work directly in the real desktop app, not through a file parser."
- duration: 8s
- poster: 5s
- transition_in: zoom-through
- status: animated
- src: compositions/frames/02-two-ways.html
- type: product_intro
- persuasion: Clear primary path
- beat: clarity + control
- blueprint: comparison-split
- asset_candidates: assets/powerpoint-mcp-server.png — captured PowerPoint MCP product artwork
- focal: assets/powerpoint-mcp-server.png
- roles: powerpoint-mcp-server = supporting brand anchor between the two interface cards

narrativeRole: Establish MCP as the main product experience and CLI as a compact alternate interface.
keyMessage: Start with natural-language MCP; use `pptcli` when scripts or LLMs benefit from a more token-efficient command surface.

Adapt: preserve the split-card motion, but make the MCP card clearly dominant and the CLI card compact.
Scene 1 (0.0–1.6s): “Start with MCP.” enters from the right into the upper-left third while a small product mark seats beside it — asymmetric header, medium density; per-word staggered reveal (`dynamic-content-sequencing`).
Scene 2 (1.6–5.3s): the large MCP card enters first, followed by a smaller `pptcli` card labeled as optional, token-efficient control for scripts and LLMs — three depth layers; split-tilt cards (`split-tilt-cards`) preserve the visual signature without implying equal priority.
Scene 3 (5.3–8.0s): PowerPoint-orange connectors draw from both cards into “SAME REAL POWERPOINT AUTOMATION ENGINE,” which locks beneath them and holds without drift — SVG self-draw (`svg-path-draw`).

## Frame 3 — Ask for the deck

- scene: A natural-language request turns into compact MCP tool calls beside the polished PowerPoint result.
- voiceover: "Start with a simple request. Ask for an executive dashboard, a project update, or a complete presentation."
- duration: 9s
- poster: 6s
- transition_in: push-slide LEFT
- status: animated
- src: compositions/frames/03-ask-for-deck.html
- type: feature_showcase
- persuasion: Friction reduction
- beat: curiosity
- blueprint: prompt-type-submit-generate
- asset_candidates: assets/mcp-executive-demo.png — native PowerPoint export of the executive dashboard
- focal: assets/mcp-executive-demo.png
- roles: mcp-executive-demo = polished PowerPoint result that answers the natural-language request

narrativeRole: Show that the workflow begins with intent instead of manual clicking.
keyMessage: Describe the outcome; MCP translates the request into precise PowerPoint operations.

Adapt: keep the type-submit signature and let the real desktop recording answer instead of a reconstructed result panel.
Scene 1 (0.0–2.0s): an executive-dashboard request types into a white editorial prompt card on the left while the completed native PowerPoint slide waits dimmed on the right — asymmetric 45/55 layout, layered depth; type-on with caret (`discrete-text-sequence`).
Scene 2 (2.0–5.7s): the prompt card compresses upward and a warm-navy MCP activity panel opens beneath it; presentation creation and 31 native shape operations resolve line-by-line — panel-to-surface coupling, medium density.
Scene 3 (5.7–9.0s): the final MCP action lands, the polished dashboard brightens, and the workbench nudges left to give the result more room — the machine response is the reveal; nudge curve (`nudge-curve`) gives the output room.

## Frame 4 — Watch the real app respond

- scene: The native dashboard export fills the hero frame over a subtle, blurred real PowerPoint desktop recording.
- voiceover: "Behind the scenes, MCP turns that intent into precise PowerPoint operations: creating slides, placing native shapes, formatting text, and building charts."
- duration: 12s
- poster: 8s
- transition_in: push-slide LEFT
- status: animated
- src: compositions/frames/04-real-app-responds.html
- type: feature_showcase
- persuasion: Show-don't-tell proof
- beat: trust + momentum
- blueprint: device-surface-showcase
- asset_candidates: assets/mcp-executive-demo.png — native PowerPoint export; assets/live-powerpoint-demo.mp4 — subdued authentic desktop-motion texture
- focal: assets/mcp-executive-demo.png
- roles: mcp-executive-demo = primary product proof; live-powerpoint-demo = blurred background motion proving the desktop app is involved

narrativeRole: Prove that the automation drives the installed PowerPoint desktop application rather than rewriting a file offline.
keyMessage: The real PowerPoint app performs and renders every edit.

Reproduce: use the static-tour variant, with the actual PowerPoint window as the persistent hero and no fake controls.
Scene 1 (0.0–2.0s): the polished PowerPoint dashboard fills a framed hero surface while a subtle blurred desktop recording moves behind it; an orange “NATIVE POWERPOINT OUTPUT” label enters and locks to the upper-left.
Scene 2 (2.0–8.7s): a restrained crop camera focuses on the dashboard title and KPI area, then returns to the full slide — camera with intent (`coordinate-target-zoom`) limited to one focus-and-return arc.
Scene 3 (8.7–12.0s): “Real slides. Real PowerPoint.” reveals over a dark lower-third panel, then holds with no further camera movement — per-word staggered reveal (`dynamic-content-sequencing`).

## Frame 5 — Build, then verify

- scene: The polished native slide remains visible while a clean MCP export-and-inspection receipt takes focus.
- voiceover: "The result is a real, editable deck, rendered by PowerPoint itself. The AI exports each slide, inspects the layout, and verifies the result."
- duration: 10s
- poster: 7s
- transition_in: push-slide LEFT
- status: animated
- src: compositions/frames/05-build-verify.html
- type: benefit_highlight
- persuasion: Risk reversal
- beat: reassurance
- blueprint: agent-progress-theater
- asset_candidates: assets/mcp-executive-demo.png — completed native slide with KPIs, roadmap, and chart
- focal: assets/mcp-executive-demo.png
- roles: mcp-executive-demo = left-side native output proof and final-state receipt

narrativeRole: Turn the visible edit into a reliability story by showing that output can be exported and inspected.
keyMessage: PowerPoint creates the result, and export-to-verify makes visual mistakes discoverable.

Adapt: keep the machine-working-to-receipt arc; the live chart creation is the working state and the export checklist is the receipt.
Scene 1 (0.0–3.2s): the live PowerPoint footage occupies the left 60% as the chart arrives; “Build.” appears above it and the poster is held behind as a quiet backup layer — asymmetric 60/40, medium density.
Scene 2 (3.2–6.2s): “Export.” types into the right receipt card while a thin orange progress line fills; the video settles on the completed slide — sequential state theater (`discrete-text-sequence`, `stat-bars-and-fills`).
Scene 3 (6.2–9.0s): three receipt rows arrive one by one — Slide exported, Layout inspected, Result verified — and each outline badge flips to a check exactly as its line lands; checklist mutation (`scale-swap-transition`, `svg-path-draw`).
Scene 4 (9.0–10.0s): “Verify.” becomes the dominant word above the completed receipt and the full frame holds still.

## Frame 6 — Complete presentation control

- scene: A compact feature field assembles around the product: slides, shapes, text, tables, charts, notes, accessibility, and export.
- voiceover: "For scripts and coding agents, pptcli provides the same engine through a compact, token-efficient command line."
- duration: 9s
- poster: 6s
- transition_in: zoom-through
- status: animated
- src: compositions/frames/06-complete-control.html
- type: benefit_highlight
- persuasion: Value stacking
- beat: capability + confidence
- blueprint: grid-card-assemble
- asset_candidates: assets/powerpoint-mcp-server.png — captured PowerPoint MCP product artwork
- focal: assets/powerpoint-mcp-server.png
- roles: powerpoint-mcp-server = supporting product anchor behind the capability field

narrativeRole: Show the breadth behind the demo without turning the video into a long feature list.
keyMessage: The same system covers the full presentation workflow across 16 tool domains.

Reproduce: keep the short-path card assembly and avoid a shared-center burst for the eight-item field.
Scene 1 (0.0–1.8s): “Control the whole deck.” enters at the upper-left while the product mark fixes the opposite corner — asymmetric editorial header, medium density.
Scene 2 (1.8–6.5s): eight capability cards arrive in two waves directly into their slots — Slides, Shapes, Text, Tables, then Charts, Notes, Accessibility, Export — four-column grid, three depth layers; short-path stagger assembly (`center-outward-expansion`) with binary waterfall timing.
Scene 3 (6.5–9.0s): the mono proof line “16 TOOL DOMAINS · 186 OPERATIONS” draws in under the grid, an orange sweep passes once behind the cards, then everything holds — traveling sheen (`ambient-glow-bloom`) used once, not as a loop.

## Frame 7 — Use the interface that fits

- scene: MCP-first positioning and the optional CLI resolve into the PowerPoint MCP lockup and website call to action.
- voiceover: "Sixteen domains. One hundred eighty-six operations. Get started at powerpointmcpserver.dev."
- duration: 7s
- poster: 5s
- transition_in: push-slide LEFT
- status: animated
- src: compositions/frames/07-use-the-interface.html
- type: cta
- persuasion: Clear next step
- beat: motivation
- blueprint: logo-assemble-lockup
- asset_candidates: assets/powerpoint-mcp-server.png — captured PowerPoint MCP product artwork
- focal: assets/powerpoint-mcp-server.png
- roles: powerpoint-mcp-server = central brand lockup

narrativeRole: Close on the product choice and send viewers to the project site.
keyMessage: Bring PowerPoint to your AI with MCP at powerpointmcpserver.dev; `pptcli` remains available as an optional compact interface.

Adapt: keep the assembling brand lockup and use the two interface pills as satellites that resolve into one final call to action.
Scene 1 (0.0–2.0s): the captured product mark assembles at center from a tight cluster while the black field clears around it — centered low-density climax; depth scatter-assemble (`depth-scatter-assemble`) with a smooth settle.
Scene 2 (2.0–4.6s): “Bring PowerPoint to your AI.” reveals beneath the mark, then MCP PRIMARY and OPTIONAL `pptcli` pills arrive from opposite sides and seat below the line — centered hierarchy, three depth layers; per-word stagger plus mirrored short-path entries.
Scene 3 (4.6–7.0s): the interface pills slide together into a single orange call-to-action bar reading `powerpointmcpserver.dev`; the site address completes left-to-right and holds to the final frame — card morph-anchor (`card-morph-anchor`) and clip reveal, no fade-out.
