---
"powerpointmcp": patch
---

Keep existing PowerPoint sessions available after cancelled commands or expired waits, and show unfinished work and later failures in session lists. Keep service status and shutdown responsive while editing is busy, and leave unsaved presentations open when the service stops instead of terminating PowerPoint. Update dependencies and ensure `service start --pipe-name` uses the requested service.

Track clipboard waits as part of the running request, and keep service controls available when recorded PowerPoint processes cannot be checked. Preserve support for VS Code 1.125; retain the previous video-tool version because the newer version adds dependencies outside the repository's accepted licenses.
