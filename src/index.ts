interface McpToolDefinition {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
}

interface McpToolExport {
  tools: McpToolDefinition[];
  callTool: (name: string, args: Record<string, unknown>) => Promise<unknown>;
  meter?: { credits: number };
  cost?: Record<string, unknown>;
  provider?: string;
}

/**
 * Microsoft OneNote (Microsoft 365) MCP Pack
 *
 * Requires OAuth connection — gateway injects credentials via _context.microsoft.
 * Read-only access to OneNote notebooks, sections, and pages via Microsoft Graph v1.0.
 * Tools: list notebooks, list sections, list pages, get page content.
 */


interface OneNoteContext {
  microsoft?: { accessToken: string };
}

const API = 'https://graph.microsoft.com/v1.0';

/**
 * JSON fetch helper for Microsoft Graph.
 * - Returns { error: 'connection_required' } when no OAuth token is present.
 * - Returns { error: <status>, message: <body text> } on non-2xx responses.
 * - Otherwise returns the parsed JSON body.
 */
async function gFetch(
  ctx: OneNoteContext,
  url: string,
  options: RequestInit = {},
): Promise<unknown> {
  if (!ctx.microsoft) {
    return {
      error: 'connection_required',
      message: 'Connect your Microsoft 365 account at https://pipeworx.io/account',
    };
  }
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${ctx.microsoft.accessToken}`,
      'Content-Type': 'application/json',
      ...(options.headers ?? {}),
    },
  });
  if (!res.ok) {
    const text = await res.text();
    return { error: res.status, message: text };
  }
  return res.json();
}

/**
 * Text fetch helper for the OneNote page content endpoint, which returns
 * text/html rather than JSON. Same auth/error contract as gFetch, but reads
 * the body as text instead of parsing JSON.
 * - Returns { error: 'connection_required' } when no OAuth token is present.
 * - Returns { error: <status>, message: <body text> } on non-2xx responses.
 * - Otherwise returns the raw text body.
 */
async function gFetchText(
  ctx: OneNoteContext,
  url: string,
  options: RequestInit = {},
): Promise<string | { error: unknown; message: string }> {
  if (!ctx.microsoft) {
    return {
      error: 'connection_required',
      message: 'Connect your Microsoft 365 account at https://pipeworx.io/account',
    };
  }
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${ctx.microsoft.accessToken}`,
      ...(options.headers ?? {}),
    },
  });
  if (!res.ok) {
    const text = await res.text();
    return { error: res.status, message: text };
  }
  return res.text();
}

const tools: McpToolExport['tools'] = [
  {
    name: 'list_notebooks',
    description:
      'List the OneNote notebooks in the signed-in user\'s Microsoft 365 / OneNote account. Returns each notebook\'s id, display name, created/last-modified times, and whether it is the default notebook. Use to discover a user\'s OneNote notebooks before browsing sections or pages.',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'list_sections',
    description:
      'List the sections in a OneNote notebook (Microsoft 365). Pass a notebook_id to list that notebook\'s sections, or omit it to list all sections across every notebook. Returns each section\'s id, display name, and created/last-modified times. Use to navigate a user\'s OneNote notes before listing pages.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        notebook_id: {
          type: 'string',
          description: 'Optional OneNote notebook ID (from list_notebooks). If omitted, returns sections across all notebooks.',
        },
      },
      required: [],
    },
  },
  {
    name: 'list_pages',
    description:
      'List OneNote pages (Microsoft 365 notes). Pass a section_id to list pages in that section, or omit it to list recent pages across all notebooks (newest first). Optionally filter with a free-text search query. Returns each page\'s id, title, and created/last-modified times. Use to find a user\'s OneNote notes/pages before reading their content.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        section_id: {
          type: 'string',
          description: 'Optional OneNote section ID (from list_sections). If omitted, returns recent pages across all notebooks.',
        },
        top: {
          type: 'number',
          description: 'Maximum number of pages to return (default 50, max 100).',
        },
        search: {
          type: 'string',
          description: 'Optional free-text search query across OneNote pages (matches title and content).',
        },
      },
      required: [],
    },
  },
  {
    name: 'get_page_content',
    description:
      'Get the full HTML content of a single OneNote page (Microsoft 365 note) by its ID. Returns the page body as HTML. Use after list_pages to read a user\'s OneNote note in full.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        page_id: {
          type: 'string',
          description: 'The ID of the OneNote page to retrieve (from list_pages).',
        },
      },
      required: ['page_id'],
    },
  },
];

async function callTool(name: string, args: Record<string, unknown>): Promise<unknown> {
  const context = (args._context ?? {}) as OneNoteContext;
  delete args._context;

  switch (name) {
    case 'list_notebooks': {
      const params = new URLSearchParams({
        $select: 'id,displayName,createdDateTime,lastModifiedDateTime,isDefault',
      });
      const result = await gFetch(context, `${API}/me/onenote/notebooks?${params}`);
      const value = (result as { value?: unknown[] }).value;
      if (Array.isArray(value)) return value;
      return result;
    }
    case 'list_sections': {
      const notebookId = args.notebook_id as string | undefined;
      const params = new URLSearchParams({
        $select: 'id,displayName,createdDateTime,lastModifiedDateTime',
      });
      const url = notebookId
        ? `${API}/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sections?${params}`
        : `${API}/me/onenote/sections?${params}`;
      const result = await gFetch(context, url);
      const value = (result as { value?: unknown[] }).value;
      if (Array.isArray(value)) return value;
      return result;
    }
    case 'list_pages': {
      const sectionId = args.section_id as string | undefined;
      const top = Math.min(100, Math.max(1, (args.top as number) ?? 50));
      const search = args.search as string | undefined;

      const params = new URLSearchParams({
        $top: String(top),
        $select: 'id,title,createdDateTime,lastModifiedDateTime',
      });
      if (search) {
        params.set('$search', search);
      }

      let url: string;
      if (sectionId) {
        url = `${API}/me/onenote/sections/${encodeURIComponent(sectionId)}/pages?${params}`;
      } else {
        if (!search) {
          params.set('$orderby', 'lastModifiedDateTime desc');
        }
        url = `${API}/me/onenote/pages?${params}`;
      }
      const result = await gFetch(context, url);
      const value = (result as { value?: unknown[] }).value;
      if (Array.isArray(value)) return value;
      return result;
    }
    case 'get_page_content': {
      const pageId = args.page_id as string;
      const result = await gFetchText(
        context,
        `${API}/me/onenote/pages/${encodeURIComponent(pageId)}/content`,
      );
      if (typeof result === 'string') {
        return {
          page_id: pageId,
          html: result.slice(0, 100000),
          truncated: result.length > 100000,
        };
      }
      return result;
    }
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

export default { tools, callTool, meter: { credits: 1 }, provider: 'microsoft' } satisfies McpToolExport;
