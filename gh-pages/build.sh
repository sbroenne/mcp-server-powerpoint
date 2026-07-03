#!/bin/bash
# Jekyll build script for PowerPoint MCP Server documentation
# This script copies shared content files before building Jekyll
# Used by both local development and GitHub Actions
#
# Unlike mcp-server-excel, this repo does not yet have root-level FEATURES.md/CHANGELOG.md
# or a docs/ folder with INSTALLATION.md/CONTRIBUTING.md/SECURITY.md. Where a source file is
# missing, this script writes a small placeholder into _includes/ so the build never fails.

set -e  # Exit on error

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"

echo "📁 Copying shared content files..."

# Create _includes directory if it doesn't exist
mkdir -p "$SCRIPT_DIR/_includes"

# Copy CHANGELOG.md from root (centralized changelog for all components), if present.
# Strip top H1 block (title + paragraph) and convert remaining H1 to H2.
if [ -f "$ROOT_DIR/CHANGELOG.md" ]; then
    awk '
        BEGIN { inheader=0; headerdone=0 }
        {
            if (headerdone==0 && /^# /) { inheader=1; next }                 # drop H1 title
            if (inheader==1 && /^This changelog/) { next }                   # drop description line
            if (inheader==1 && /^$/) { inheader=0; headerdone=1; next }      # blank line ends header
            if (/^# /) { sub(/^# /, "## "); print; next }                   # convert any remaining H1 → H2
            print
        }
    ' "$ROOT_DIR/CHANGELOG.md" > "$SCRIPT_DIR/_includes/changelog.md"
    echo "   ✓ Copied CHANGELOG.md (stripped top H1 block, H1→H2)"
else
    echo "_No CHANGELOG.md found at the repo root yet — check back after the first release._" > "$SCRIPT_DIR/_includes/changelog.md"
    echo "   ⚠ CHANGELOG.md not found, wrote placeholder"
fi

# Copy MCP Server README (strip top H1 block and badge lines), if present.
if [ -f "$ROOT_DIR/src/PowerPointMcp.McpServer/README.md" ]; then
    awk '
        BEGIN { inheader=0; headerdone=0 }
        {
            if (headerdone==0 && /^# /) { inheader=1; next }                 # drop H1 title
            if (inheader==1 && /^<!-- mcp-name/) { next }                    # drop mcp-name comment
            if (inheader==1 && /^mcp-name:/) { next }                        # drop mcp-name line
            if (inheader==1 && /^\[!\[/) { next }                            # drop badge lines
            if (inheader==1 && /^$/) { inheader=0; headerdone=1; next }      # blank line ends header
            if (/^# /) { sub(/^# /, "## "); print; next }                   # convert any remaining H1 → H2
            print
        }
    ' "$ROOT_DIR/src/PowerPointMcp.McpServer/README.md" > "$SCRIPT_DIR/_includes/mcp-server.md"
    echo "   ✓ Copied MCP Server README (stripped top H1 block, badges, H1→H2)"
else
    echo "_MCP Server README not published yet — see the [GitHub repository](https://github.com/sbroenne/mcp-server-powerpoint) for current status._" > "$SCRIPT_DIR/_includes/mcp-server.md"
    echo "   ⚠ src/PowerPointMcp.McpServer/README.md not found, wrote placeholder"
fi

# Copy CLI README (strip top H1 block and badge lines), if present.
if [ -f "$ROOT_DIR/src/PowerPointMcp.CLI/README.md" ]; then
    awk '
        BEGIN { inheader=0; headerdone=0 }
        {
            if (headerdone==0 && /^# /) { inheader=1; next }                 # drop H1 title
            if (inheader==1 && /^\[!\[/) { next }                            # drop badge lines
            if (inheader==1 && /^$/) { inheader=0; headerdone=1; next }      # blank line ends header
            if (/^# /) { sub(/^# /, "## "); print; next }                   # convert any remaining H1 → H2
            print
        }
    ' "$ROOT_DIR/src/PowerPointMcp.CLI/README.md" > "$SCRIPT_DIR/_includes/cli.md"
    echo "   ✓ Copied CLI README (stripped top H1 block, badges, H1→H2)"
else
    echo "_CLI README not published yet — see the [GitHub repository](https://github.com/sbroenne/mcp-server-powerpoint) for current status._" > "$SCRIPT_DIR/_includes/cli.md"
    echo "   ⚠ src/PowerPointMcp.CLI/README.md not found, wrote placeholder"
fi

# Copy Agent Skills README (strip top H1 block)
awk '
    BEGIN { inheader=0; headerdone=0 }
    {
        if (headerdone==0 && /^# /) { inheader=1; next }                 # drop H1 title
        if (inheader==1 && /^$/) { inheader=0; headerdone=1; next }      # blank line ends header
        if (/^# /) { sub(/^# /, "## "); print; next }                   # convert any remaining H1 → H2
        print
    }
' "$ROOT_DIR/skills/README.md" > "$SCRIPT_DIR/_includes/skills.md"
echo "   ✓ Copied Agent Skills README (stripped top H1 block, H1→H2)"

# Determine build mode
if [ "$1" == "serve" ]; then
    echo ""
    echo "🚀 Starting Jekyll server..."
    cd "$SCRIPT_DIR"
    bundle exec jekyll serve --host 127.0.0.1 --port 4000
elif [ "$1" == "production" ] || [ "$JEKYLL_ENV" == "production" ]; then
    echo ""
    echo "🏗️  Building for production..."
    cd "$SCRIPT_DIR"
    JEKYLL_ENV=production bundle exec jekyll build
else
    echo ""
    echo "🏗️  Building for development..."
    cd "$SCRIPT_DIR"
    bundle exec jekyll build
fi

echo ""
echo "✅ Build complete!"
