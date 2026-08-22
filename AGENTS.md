# Repository Guidance for Coding Agents

This file applies to the entire repository. It is the authoritative engineering
guidance for automated contributors. Repository maintainers may approve an
explicit exception, but an agent must explain the tradeoff before implementing
one.

## Project boundary

`mcp-microsoft` is a Microsoft Graph integration exposed through the Model
Context Protocol (MCP). Keep the server focused on Graph operations and MCP
primitives. A feature being technically possible inside the server does not
mean it belongs in this layer.

Before implementing a feature, verify all three of these points:

1. The behavior fits the intent and current specification of MCP.
2. The Microsoft Graph endpoint, query, permission, and response shape are
   supported by the official Graph documentation.
3. The feature belongs in a generic Graph MCP server rather than in the MCP
   host, client UI, model runtime, or a specialized optional adapter.

Do not implement a feature that conflicts with MCP's intended separation of
responsibilities or established protocol best practices. Explain the conflict
and propose the correct-layer alternative instead. If a useful subset belongs
here, split it into a bounded server change and keep the host/client behavior
separate.

Examples of this boundary:

- Return original file bytes and metadata through appropriate MCP/FastMCP
  content types. Do not add generic PDF parsing, OCR, document rendering,
  summarization, or model-oriented chunking to this server; those depend on the
  host's model capabilities, context policy, and user experience.
- Use negotiated MCP capabilities for optional client interactions. Do not
  assume that every client implements elicitation, roots, sampling, tasks, or
  MCP Apps.
- In remote HTTP mode, the server filesystem is not the caller's filesystem.
  Do not expose server-local paths as if a remote client could use them.

Primary protocol references:

- https://modelcontextprotocol.io/specification/latest/server
- https://modelcontextprotocol.io/specification/latest/server/tools
- https://modelcontextprotocol.io/specification/latest/schema

## Documentation-first implementation

Every implementation must be grounded in current documentation for the APIs,
protocols, SDKs, and libraries it calls. Do not rely only on model memory,
third-party examples, or behavior inferred from names.

- Before coding, consult up-to-date primary documentation through Context7,
  official-source web search, the Firecrawl developer index, or an equivalent
  documentation retrieval tool. Prefer official specifications, API
  references, SDK references, and versioned changelogs over blogs or snippets.
- Match library guidance to the version locked by this repository. Inspect the
  installed API or source when versioned documentation is ambiguous.
- For Graph changes, verify the exact resource, HTTP method, API version,
  parameters, response fields, pagination behavior, permissions, licensing,
  and documented limitations. For MCP changes, verify the current protocol
  capability and schema rather than assuming all clients support it.
- Link the decisive documentation in the PR description or review discussion
  and distinguish what was verified from what remains assumed or untestable.
- If authoritative documentation is unavailable or contradicts the proposed
  behavior, do not guess. Stop, explain the uncertainty, and request evidence
  or propose a documented alternative.

## Tool contracts

Tools are an API for models and hosts. Their contracts must be typed, stable,
and self-explanatory.

- Define public tool inputs as Pydantic models derived from
  `ToolRequestModel`. Use precise field types, constraints, defaults, and
  descriptions. Validate mutually exclusive fields and bounded values before
  making network calls.
- Define public outputs as explicit structured response models. Prefer named
  Pydantic models with stable fields over raw `dict[str, Any]`, ambiguous text,
  or provider-shaped payloads. Keep unavoidable untyped Graph data isolated at
  the Graph boundary and convert it before returning from the tool.
- Provide clear tool and model descriptions. A model should be able to tell
  when to use the tool, what each argument means, what comes back, which
  permissions or capabilities are required, and whether the operation has side
  effects without reading the implementation.
- Preserve machine-readable success and error information. Do not hide a
  partial failure inside a success-shaped response or require clients to parse
  prose to determine the outcome.
- Register tools through the repository's `register_tool` helper with accurate
  annotations. Mark read-only, idempotent, mutating, and destructive behavior
  honestly; annotations are part of the public contract.
- Keep tool names and schemas compatible when possible. If transports require
  different safe inputs, register a transport-specific typed model under the
  same public tool name and test the exposed schemas.
- Add schema-level tests for new or changed tools, not only implementation
  tests. Verify required fields, omitted unsafe fields, output structure, and
  transport-specific behavior.

## Microsoft Graph correctness

- Treat Microsoft Learn's Graph API documentation as authoritative. Verify the
  exact endpoint, API version, supported OData parameters, escaping rules,
  permission type, and least-privileged delegated scope before coding.
- Do not invent filters or query combinations that Graph does not document.
  When a local fallback is necessary, bound its pages, items, memory, retries,
  and elapsed work, and preserve a usable pagination contract.
- Follow `@odata.nextLink` only within an explicit work budget. Treat
  continuation tokens and URLs as opaque and validate any client-provided
  cursor according to the tool contract.
- Request the least-privileged permission that supports the feature. New
  scopes, admin consent, enterprise-only behavior, beta endpoints, or licensing
  requirements must be documented and tested where possible.
- Keep retries bounded and limited to operations that are safe to retry. Never
  automatically replay an irreversible or billable request unless the API's
  idempotency guarantees make that safe.

Microsoft Graph permission guidance:

- https://learn.microsoft.com/graph/best-practices-graph-permission

## Security and reliability

- Fail closed for requested confirmation, authorization, destructive actions,
  unsupported client capabilities, and invalid security-sensitive input.
- Preserve the existing delegated-user identity model. Never fall back from a
  caller's identity to a shared profile in HTTP mode.
- Do not log access tokens, secrets, message bodies, attachment contents, tool
  arguments, or full tool results.
- Bound all externally controlled collections, uploads, downloads, pagination,
  concurrency, caches, and retry loops. Avoid loading unbounded Graph results
  or files into memory.
- Keep error details useful locally while respecting the HTTP transport's
  masking and audit-log boundaries.

## Change discipline

- Read the surrounding module, response models, registration code, tests, and
  relevant documentation before editing.
- Make the smallest coherent change. Keep unrelated refactors and dependency
  updates out of focused fixes.
- Add or update tests for success, validation, failure, capability fallback,
  pagination/bounds, and transport differences as applicable.
- Update README/tool counts, DevOps guides, permissions, manifests, and lock
  files whenever the public surface or dependencies change.
- Run focused tests while iterating, then run the full locked suite before
  publishing:

  ```bash
  uv run --frozen pytest -q
  ```

- Do not weaken tests, disable checks, or bypass branch protections to obtain a
  passing result.
