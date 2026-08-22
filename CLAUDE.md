# Claude Code Instructions

Read and follow [`AGENTS.md`](AGENTS.md) before planning or changing this
repository. It is the authoritative repository-wide engineering policy.

In particular:

- Keep this project a Microsoft Graph MCP server. Do not move host/client/model
  responsibilities such as generic PDF extraction, OCR, rendering,
  summarization, or context chunking into the server.
- Verify proposed behavior against the current MCP specification and official
  Microsoft Graph documentation. If a request conflicts with protocol intent
  or belongs in another layer, explain why and propose that alternative instead
  of implementing it here.
- Ground every implementation in current primary API and library documentation
  using Context7, official-source web search, the Firecrawl developer index, or
  an equivalent retrieval tool. Match library guidance to locked versions,
  cite the decisive sources in the PR, and do not guess when documentation is
  missing or contradictory.
- Define tools with typed `ToolRequestModel` inputs, explicit structured output
  models, accurate annotations, and clear descriptions of purpose, arguments,
  results, permissions, capabilities, and side effects.
- Negotiate optional client capabilities and provide safe fallbacks. Never
  assume elicitation or another optional MCP feature exists.
- Keep remote HTTP contracts free of caller-unusable server-local paths.
- Bound pagination, files, memory, concurrency, and retries; use least-privilege
  Graph permissions; fail closed for sensitive or destructive behavior.
- Make focused changes with schema-level tests and run the full locked suite
  before publishing.
