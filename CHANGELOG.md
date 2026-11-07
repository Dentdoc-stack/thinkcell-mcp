# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Chart templates (`library.py`), with three new tools:
  - `list_chart_templates`: think-cell's own chart library, cut into 34
    templates from the local install, plus the user's saved slides.
  - `fill_chart_template`: fills a template's charts through `ppttc.exe`,
    sets the title and texts, and checks which charts took the data.
  - `template_from_slide`: saves a slide of the user's own deck (or the one
    open in PowerPoint) as a template, naming its charts in code.
- The "100%=" datasheet row is detected and filled automatically.

## [0.2.0] - 2026-11-07

### Added

- **Claude Desktop extension:** `manifest.json` plus `build_mcpb.py` pack a
  one-click `.mcpb` (uv runtime, chosen output folder).
- **Remote connector** for Claude in PowerPoint and claude.ai:
  - `server.py --http` serves streamable HTTP on 127.0.0.1.
  - OAuth 2.1 sign-in approved with a connect code (`oauth.py`).
  - Path arguments are limited to `output/`, `templates/` and the folders
    chosen in setup.
- `connector.py`: `setup`, `start`/`stop`/`status`, `code`, `signout` and
  `autostart`. It runs the server and a fixed ngrok address in the
  background. See `CONNECTOR.md`.
- New tools:
  - `insert_into_powerpoint`: inserts a converted deck's slides into the
    presentation open in PowerPoint.
  - `import_chart_data`: reads chart data from Excel or CSV.
  - `list_color_palettes`.
- `table` and `process` chart types; auto-deck layouts and palettes.
- `THINKCELL_OUTPUT_DIR` moves the generated files.
- **think-cell Assist support:** the remote connector now handles Assist,
  following its manual's "Connect MCP servers" section:
  - It answers CORS preflights for `https://app.prod.ai.think-cell.com` and
    sends the required allow and expose headers.
  - It accepts an API key as `Authorization: Bearer`.
  - `connector.py assist` prints the Windows group-policy and macOS
    `defaults` settings.
- **MCP 2026-07-28 on the remote connector** (`modern_http.py`): stateless
  requests that carry the version in `_meta` are answered directly, with
  `server/discover`, `tools/list` and `tools/call`, the `MCP-Protocol-Version`,
  `Mcp-Method` and `Mcp-Name` header checks (-32020), unsupported versions
  (-32022) and 404 for other methods. Older clients still use `initialize` on
  the same URL. stdio stays on the older protocol.

### Changed

- The think-cell install folder is found from the Windows registry, so
  installs outside `C:\Program Files (x86)\think-cell` work without
  `THINKCELL_DIR`.
- `pyproject.toml` lists every runtime dependency (adds `openpyxl` and
  `pywin32`).

### Fixed

- **Origin check:** remote mode now checks `Origin`, as the MCP Streamable
  HTTP transport requires. Requests from origins it does not know get 403;
  requests without an Origin still pass.
- **GET/DELETE on `/mcp`:** these now get 405. Before, GET opened a stream
  that never ended.
- **Dates in `validate_ppttc`:** date cells must be ISO `YYYY-MM-DD`, as the
  think-cell manual requires.
- Antivirus web shields (Norton) broke HTTP replies that carried
  `Content-Length`. Remote mode now always sends chunked replies.

## [0.1.0] - 2026-05-16

### Added

- MCP server (Python, stdio transport) that wraps think-cell's `.ppttc` JSON
  automation: builds `.ppttc` files from plain data and converts them into
  PowerPoint decks (`.pptx`) via think-cell's `ppttc.exe`.
- 8 MCP tools: `create_chart`, `build_presentation`, `create_auto_deck`,
  `set_deck_branding`, `convert_to_pptx`, `validate_ppttc`, `list_chart_types`,
  and `diagnose_thinkcell`.
- 8 chart types: `waterfall`, `bar`, `stacked_bar`, `line`, `scatter`,
  `mekko`, `area`, and `combo` — each with per-type input validation and data
  shaping.
- `create_auto_deck`: zero-setup multi-slide decks built on think-cell's
  official automation template, requiring no PowerPoint template preparation
  or element naming. Each slide carries a fixed layout (slide title plus a
  left and right chart, both on a date axis); each slide becomes one `.ppttc`
  entry.
- `set_deck_branding`: re-themes the auto-deck template — removes the
  think-cell logo and applies a configurable accent palette (a
  Big-Four-inspired palette by default). Only standard PowerPoint parts are
  touched, so named think-cell elements stay intact.
- `build_presentation`: combines multiple charts and slides into a single
  `.ppttc` file, with per-chart names emitted as one `.ppttc` entry per slide.
- Structural `.ppttc` validator (`validate_ppttc`) that checks JSON structure,
  required keys, cell types, and row consistency.
- Read-only think-cell environment diagnostics (`diagnose_thinkcell`):
  Windows registry and filesystem checks for the install, PowerPoint, the COM
  add-in registration, disabled add-ins, licence, and template content — no
  processes are launched.
- `list_chart_types`: catalog of every chart type with its required/optional
  parameters, expected data shape, and a ready-to-use example.
- Self-contained test suite (`python tests.py`) and GitHub Actions CI.

[0.1.0]: https://github.com/Hectirry/thinkcell-mcp/releases/tag/v0.1.0
