# Filesystem MCP server (optional)

The skill itself writes files through Python (`save_and_validate()`), so no MCP server is
required for Claude Code. Add the filesystem MCP server when the host cannot touch the disk
directly (Claude Desktop) or when you want to restrict the skill to an explicit allow-list of
directories for source data and `.xlsx` output.

Every path passed after the package name is an allowed root. The server refuses reads and
writes outside those roots, so list the folder where inputs live and the folder where
generated workbooks should land.

## Pick the config for your OS

| OS | File | Why it differs |
|----|------|----------------|
| macOS / Linux | [`filesystem.macos-linux.json`](./filesystem.macos-linux.json) | `npx` is resolved on `PATH`; no shell wrapper needed |
| Windows | [`filesystem.windows.json`](./filesystem.windows.json) | `npx` is a `.cmd` shim, so it must be launched through `cmd /c`; paths use backslashes (escaped as `\\` in JSON) |

Replace the placeholder paths before use. Mixing the two (a `cmd /c` launcher with
`/Users/...` paths) silently fails on both platforms: Windows has no `/Users` tree and
macOS has no `cmd`.

## Where to put it

| Host | Location |
|------|----------|
| Claude Desktop | `claude_desktop_config.json` (Settings → Developer → Edit Config); paste the `mcpServers` block |
| Claude Code (project) | `.mcp.json` at the repo root, same shape |
| Claude Code (one-liner) | `claude mcp add filesystem -- npx -y @modelcontextprotocol/server-filesystem /abs/path/one /abs/path/two` |

Requires Node.js 18+ (`node --version`). The first launch downloads the package via `npx -y`.

## Verify

Restart the host, then ask: `List the files in <allowed dir>`. If the tool call fails,
check that each path is absolute, exists, and that `npx` runs from a plain terminal.
