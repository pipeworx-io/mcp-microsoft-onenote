# mcp-microsoft-onenote

Microsoft OneNote (Microsoft 365) MCP Pack

Part of [Pipeworx](https://pipeworx.io) — an MCP gateway connecting AI agents to 1683+ live data sources.

## Tools

| Tool | Description |
|------|-------------|
| `list_notebooks` | List the OneNote notebooks in the signed-in user's Microsoft 365 / OneNote account. Returns each notebook's id, display name, created/last-modified times, and whether it is the default notebook. Use to discover a user's OneNote notebooks before browsing sections or pages. |
| `list_sections` | List the sections in a OneNote notebook (Microsoft 365). Pass a notebook_id to list that notebook's sections, or omit it to list all sections across every notebook. Returns each section's id, display name, and created/last-modified times. Use to navigate a user's OneNote notes before listing pages. |
| `list_pages` | List OneNote pages (Microsoft 365 notes). Pass a section_id to list pages in that section, or omit it to list recent pages across all notebooks (newest first). Optionally filter with a free-text search query. Returns each page's id, title, and created/last-modified times. Use to find a user's OneNote notes/pages before reading their content. |
| `get_page_content` | Get the full HTML content of a single OneNote page (Microsoft 365 note) by its ID. Returns the page body as HTML. Use after list_pages to read a user's OneNote note in full. |

## Quick Start

Add to your MCP client (Claude Desktop, Cursor, Windsurf, etc.):

```json
{
  "mcpServers": {
    "microsoft-onenote": {
      "url": "https://gateway.pipeworx.io/microsoft-onenote/mcp"
    }
  }
}
```

### What this endpoint actually serves

`tools/list` at `https://gateway.pipeworx.io/microsoft-onenote/mcp` returns the tools in the table
above **plus the shared Pipeworx meta-tools** — `ask_pipeworx`,
`discover_tools`, `search_within`, `remember`/`recall` and the rest of the
gateway-wide set. So the tool count you see is larger than this table: a
single-pack endpoint currently lists roughly 30 shared tools alongside the
pack's own. The connection's `initialize` response states its exact scope, and
is the authoritative answer for a given day.

This is deliberate, not multiplexing by accident. The meta-tools are what let a
scoped connection answer a question this pack does not cover — via
`ask_pipeworx`, which routes across the whole catalog — without you adding a
second MCP server. There is currently no way to mount a pack endpoint without
them; if the extra schemas cost you more context than the routing is worth,
connect to the full gateway once rather than to several pack endpoints.

Or connect to the full Pipeworx gateway to get every pack's tools listed
directly, instead of just this one's:

```json
{
  "mcpServers": {
    "pipeworx": {
      "url": "https://gateway.pipeworx.io/mcp"
    }
  }
}
```

Both URLs reach the same gateway and the same 1683+ data sources. The
only difference is which pack's tools are listed **directly**; `ask_pipeworx`
reaches all of them from either one.

## No MCP client? Call it over HTTP

This pack runs against a connected microsoft account, so it needs a Pipeworx key: sign in at https://pipeworx.io/account, connect microsoft, then call `POST https://gateway.pipeworx.io/v1/tools/list_notebooks` with `Authorization: Bearer <your Pipeworx key>`. Inspect any tool: `GET https://gateway.pipeworx.io/v1/tools/list_notebooks`. Find one: `POST https://gateway.pipeworx.io/v1/tools/search_packs` with `{"query":"..."}`.

## Standalone (no gateway account)

This package also runs as a local stdio MCP server — no Pipeworx account, no
gateway round-trip:

```json
{
  "mcpServers": {
    "microsoft-onenote": {
      "command": "npx",
      "args": ["-y", "@pipeworx/mcp-microsoft-onenote"]
    }
  }
}
```

Or run it directly to confirm it starts:

```bash
npx -y @pipeworx/mcp-microsoft-onenote
```

It speaks MCP over stdin/stdout and answers `initialize`/`tools/list`/`tools/call`
for **only** this pack's tools — none of the shared meta-tools the gateway
connection above adds. Same source, same tools, no ask_pipeworx routing.

## Using with ask_pipeworx

Instead of calling tools directly, you can ask questions in plain English —
this works on the pack endpoint above as well as on the full gateway:

```
ask_pipeworx({ question: "your question about Microsoft Onenote data" })
```

The gateway picks the right tool and fills the arguments automatically.

## More

- [Docs and guides](https://pipeworx.io/docs)
- [pipeworx.io](https://pipeworx.io)

## License

MIT
