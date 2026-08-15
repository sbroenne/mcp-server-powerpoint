---
title: Troubleshooting
description: Fixes for common PowerPoint MCP Server errors — PowerPoint is not installed, unknown sessionId, daemon connection failures, COM HRESULT errors and lingering POWERPNT.EXE processes.
keywords: "PowerPoint MCP error, Microsoft PowerPoint is not installed, Unknown sessionId, Cannot connect to daemon, RPC server unavailable PowerPoint, 0x800A03EC, POWERPNT.EXE not closing, PowerPoint COM error"
---

# Troubleshooting

Concrete errors you may hit, what causes them, and how to fix them. If your
problem is a general "how does this work" question, try the
[FAQ](faq.md) first.

## Setup errors

### `Microsoft PowerPoint is not installed on this system.`

Thrown when the COM layer cannot create a `PowerPoint.Application` object.

**Causes and fixes:**

- **PowerPoint desktop is genuinely not installed.** The full desktop
  application is required. PowerPoint Online, PowerPoint Mobile, LibreOffice
  Impress and Google Slides do not expose COM automation and cannot be used.
- **PowerPoint is installed but not activated.** An unlicensed or
  reduced-functionality install refuses automation requests. Open PowerPoint
  manually and confirm it starts fully.
- **Microsoft Store / UWP build of Office.** Some sandboxed Store builds do not
  register the COM class reliably. Install the Click-to-Run or MSI build of
  Microsoft 365 / Office instead.
- **Architecture mismatch.** A 32-bit Office install with a 64-bit host process
  is normally fine for out-of-process COM, but a broken registration is not.
  Repair Office from **Settings → Apps → Microsoft 365 → Modify → Quick Repair**.

### `mcp-powerpoint` or `pptcli` is not recognised

The .NET global tools directory is not on your `PATH`, or your shell cached the
old `PATH`.

```powershell
# Confirm the tool installed
dotnet tool list --global

# Default global tools location
$env:USERPROFILE\.dotnet\tools
```

Restart your terminal **and** your MCP client after
`dotnet tool install --global` so the updated `PATH` is inherited. If it still
fails, use the absolute path to the executable in your client config, or use the
[standalone executable](installation.md) which needs no .NET runtime.

### The server does not appear in my AI client

Almost always a config problem rather than a server problem:

- **Unescaped backslashes.** JSON needs `C:\\Tools\\mcp-powerpoint.exe`, not
  `C:\Tools\mcp-powerpoint.exe`. This is the single most common cause.
- **Client not fully restarted.** Claude Desktop must be quit from the system
  tray; VS Code needs a window reload; Cursor needs a restart.
- **Wrong config key.** VS Code uses `servers`, Claude Desktop and Cursor use
  `mcpServers`. See the [Claude Desktop](claude-desktop.md),
  [VS Code](vs-code.md) and [Cursor](cursor.md) guides for the exact blocks.

## File and session errors

### `PowerPoint file not found: <path>` / `File not found: <path>`

The path does not resolve on disk. Check that:

- The path is absolute, or relative to the process working directory — which is
  the *client's* working directory, not yours.
- Backslashes are escaped correctly if the path came through JSON.
- The file is not on a disconnected network share or an unsynced OneDrive path.
  A OneDrive file marked *online-only* may need to be downloaded first.

### `Unknown sessionId: <id>`

The session no longer exists in the registry. Sessions are not persistent
identifiers — they live only as long as the host process.

- The session was already closed with
  `presentation(action="close", sessionId=...)`.
- The MCP server or CLI daemon restarted, discarding every session.
- The AI reused a `sessionId` from an earlier conversation.

Re-open the file with `presentation(action="open", filePath=...)` and use the
new `sessionId`. Note that `presentation(action="create", ...)` already leaves
the new session open — reuse that id rather than opening the same file again.

### `isMacroEnabled=true requires a .pptm file path.`

Macro-enabled presentations must use the `.pptm` extension. Either change the
target path to `.pptm`, or drop `isMacroEnabled`.

### Edits fail or the file opens read-only

Close any open instance of the target presentation first — the server needs
exclusive access while it drives PowerPoint. Also check for:

- A stale `~$<name>.pptx` lock file next to the presentation.
- The file being open on another machine from a shared drive.
- A read-only attribute, or a file under a protected/synced folder.

## CLI daemon errors

The CLI (`pptcli`) talks to a background daemon over a named pipe so sessions
survive across invocations. These errors are specific to that channel.

### `Cannot connect to daemon. Is it running?`

The daemon is not running, or it exited. It auto-starts on the first
`session open` / `session create`, so:

```powershell
pptcli service status
pptcli service start
```

If it still fails, an antivirus or endpoint-protection product may be blocking
the named pipe or the child process launch.

### `Daemon connection timed out` / `Daemon request timed out`

The daemon is alive but did not answer in time. Usually one of:

- **PowerPoint is still starting.** A cold PowerPoint launch legitimately takes
  ~90–150 seconds on some machines. Retry rather than killing the daemon.
- **A modal dialog is blocking PowerPoint.** See
  [PowerPoint is stuck on a dialog](#powerpoint-is-stuck-on-a-dialog) below.
- **A long-running operation.** Exporting a large deck to images takes time.

### `Connection to the daemon was lost while waiting for a response.`

The daemon process died mid-request — typically because PowerPoint crashed or
the process was force-killed. Restart it:

```powershell
pptcli service stop
pptcli service start
```

Any sessions it held are gone; re-open your presentations.

### `Failed to start daemon: <message>`

The CLI could not launch the daemon executable. Check that the install is
complete (`dotnet tool list --global`), that the executable is not quarantined
by antivirus, and that your user has permission to start processes from that
directory.

## COM runtime errors

These surface as HRESULT codes on the tool boundary. The MCP server logs the
HRESULT to stderr and returns a structured error rather than crashing.

| HRESULT | Meaning | What to do |
|---|---|---|
| `0x800A03EC` | Generic PowerPoint automation rejection | Usually an invalid argument for the current object state — e.g. an out-of-range index, or an operation the current slide layout does not permit. Check indexes are **1-based** and in range. |
| `0x80010001` (`RPC_E_CALL_REJECTED`) | Call rejected by callee | PowerPoint is busy or showing a modal dialog. Dismiss the dialog; the OLE message filter retries transient rejections automatically. |
| `0x8001010A` (`RPC_E_SERVERCALL_RETRYLATER`) | Server busy | Transient; the message filter retries. Persistent failures mean PowerPoint is blocked on a dialog. |
| `0x800706BA` (`RPC_S_SERVER_UNAVAILABLE`) | RPC server unavailable | The PowerPoint process died underneath the session. Re-open the presentation to get a new session. |
| `0x80004005` (`E_FAIL`) | Unspecified failure | Often a locked file, a corrupt template, or an unsupported operation for the object. Check the stderr log for the failing operation. |

### `Failed to register OLE message filter. HRESULT: 0x…`

The COM apartment could not accept the message filter. This indicates the
thread is not in an STA or another filter is already installed. Restart the
host process; if it recurs, file an issue with the HRESULT value.

### PowerPoint is stuck on a dialog

COM calls block while PowerPoint shows a modal dialog — a font-substitution
prompt, a repair prompt, a "file in use" warning, or an activation nag.

Bring the PowerPoint window to the foreground and dismiss it. To avoid this
class of failure, prefer files that open cleanly by hand first, and avoid
templates that trigger font substitution on your machine.

## Cleanup and processes

### Lingering `POWERPNT.EXE` processes

**This is usually normal, not a bug.** Office's own post-quit cleanup can take
up to roughly **90–200 seconds** after a session is closed. The design
deliberately does not force-kill on the happy path:

- `presentation(action="close", sessionId=...)` removes the session from the
  registry **immediately** and disposes the underlying batch on a background
  task, so your caller never blocks.
- Host shutdown disposes every remaining session.

Wait a few minutes before assuming a leak. If a process genuinely survives a
crash of the host, end it from Task Manager.

### PowerPoint windows appearing on screen

Expected. This project drives the real, visible desktop application — that is
what makes true-fidelity rendering and export-to-verify possible. There is no
headless mode; see the [FAQ](faq.md#can-i-run-it-on-a-server-or-in-ci).

## Still stuck?

Collect the following and
[open an issue](https://github.com/sbroenne/mcp-server-powerpoint/issues):

- Windows version and PowerPoint version (**File → Account → About PowerPoint**)
- How you installed (.NET tool, standalone executable, VS Code extension, MCPB)
- Your MCP client and its version
- The exact error text, plus any HRESULT from the server's stderr log
- The tool call that failed, with arguments (redact confidential file paths)
