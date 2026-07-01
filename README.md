# mcp-server-powerpoint

**Work-in-progress.** COM/PIA-based PowerPoint automation, sibling project to
[mcp-server-excel](https://github.com/sbroenne/mcp-server-excel), following the same
architecture: `ComInterop` → `Core` → `Service` → `CLI` / `MCP Server`.

Unlike existing PowerPoint MCP servers (which manipulate `.pptx` offline via `python-pptx`)
or Anthropic's `pptx` agent skill (agent-run scripts, LibreOffice-rendered thumbnails),
this project drives a **live, real PowerPoint desktop instance** via
`Microsoft.Office.Interop.PowerPoint` (the official Primary Interop Assembly) — enabling
true-fidelity rendering, VBA, OLE objects, and interaction with an already-open deck.
Windows + PowerPoint desktop required.

## Status

Early scaffold. Only the presentation lifecycle "create" vertical slice is implemented,
proving the COM batch pattern end-to-end (STA thread + OLE message filter + channel-based
work queue, ported from mcp-server-excel's `ExcelBatch`). See `CONTINUATION.md` for the
full roadmap and what's implemented vs. still to do.

## Build

```powershell
dotnet build
```

## Try it

```powershell
dotnet run --project src\PowerPointMcp.CLI -- create C:\temp\demo.pptx
```

## License

MIT — see [LICENSE](LICENSE).
