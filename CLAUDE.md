# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Testing
```bash
# Run all tests with coverage (requires >=90% coverage)
uv run tox

# Run Python tests only
uv run pytest -v

# Run specific test file
uv run pytest tests/test_docx_post_process.py -v

# Run linting and type checking
uv run tox -e lint
```

### Code Quality
```bash
# Run pre-commit hooks on all files
uv run pre-commit run --all

# Format code with ruff
uv run ruff format

# Check code with ruff
uv run ruff check

# Type check with mypy
uv run mypy .
```

### Docker Development
```bash
# Build Docker image
docker build --build-arg APP_IMAGE_VERSION=0.0.0 --file Dockerfile --tag pandoc-service:0.0.0 .

# Run development container
docker run --init --detach --publish 9082:9082 --name pandoc-service pandoc-service:0.0.0

# Run with docker-compose
docker-compose up -d

# Stop container
docker container stop pandoc-service
```

### Container Testing
```bash
# Structure test
container-structure-test test --image pandoc-service:local --config ./tests/container/container-structure-test.yaml

# Integration test script
bash tests/shell/test_pandoc_service.sh
```

## Architecture Overview

### Core Components
- **app/PandocController.py**: FastAPI application with REST endpoints and conversion logic
- **app/PandocServiceApplication.py**: Application entry point with logging setup
- **app/DocxPostProcess.py**: DOCX-specific post-processing (tables, images)
- **app/DocxReferencesPostProcess.py**: DOCX table-of-contents and field update post-processing
- **app/PptxPostProcess.py**: PPTX slide size post-processing
- **app/schema.py**: Pydantic models for API responses
- **app/svg_processor.py**: Finds SVGs in incoming HTML and rasterizes them to PNG (port of weasyprint-service's `SvgProcessor`)
- **app/chromium_manager.py**: Persistent headless Chromium (Playwright) used for SVG→PNG, with metrics, retries and a health-monitor loop (port of weasyprint-service's `ChromiumManager`)
- **app/constants.py**: Small env-var helpers (`get_bool_env`)
- **app/prometheus_metrics.py**, **app/pandoc_metrics.py**, **app/metrics_server.py**: Prometheus metrics instrumentation (incl. SVG/Chromium metrics)

### Security Model
- Uses allowlisted pandoc options to prevent command injection
- Direct subprocess calls to pandoc binary (no shell=True)
- Input validation with 200MB request size limit
- Format validation for source/target combinations

### Conversion Pipeline
1. Request validation (format, size, encoding)
2. For HTML sources: SVG-to-PNG rasterization via headless Chromium (`preprocess_html_svgs` → `SvgProcessor` → `ChromiumManager`), so SVGs render in targets with weak SVG support (e.g. Word). Best effort: if Chromium is disabled/unavailable, the HTML passes through unchanged.
3. Pandoc subprocess execution with security constraints
4. Post-processing (especially for DOCX files)
5. Response formatting with appropriate MIME types

### SVG to PNG conversion
- Runs only for `source_format == "html"`, before pandoc.
- Density is a device scale factor: per-request `scale_factor` query param (overrides `DEVICE_SCALE_FACTOR` env, default 1.0). docx-exporter sends its "Image density" setting via this param.
- The Chromium lifecycle is managed in the FastAPI `lifespan` (`_start_chromium`/`_stop_chromium`); a start failure is logged and swallowed.
- Browser comes from Playwright's bundled Chromium (Debian base; `playwright install chromium`), not a system package. Tunable via `ENABLE_SVG_CONVERSION`, `MAX_CONCURRENT_CONVERSIONS`, `CHROMIUM_CONVERSION_TIMEOUT`, `CHROMIUM_MAX_CONVERSION_RETRIES`, `CHROMIUM_RESTART_AFTER_N_CONVERSIONS`, `CHROMIUM_HEALTH_CHECK_ENABLED`, `CHROMIUM_HEALTH_CHECK_INTERVAL` (see README).

### Supported Formats
- **Source**: docx, epub, fb2, html, json, latex, markdown, rtf, textile
- **Target**: docx, epub, fb2, html, json, latex, markdown, odt, pdf, plain, pptx, rtf, textile

### Key Configuration
- **Python** required (see `.tool-versions` for exact version)
- **uv** for dependency management (`--frozen` flag used in CI)
- **Ruff** for linting (line length: 240, TCH rule enforces `if TYPE_CHECKING:` import guards)
- **Mypy** for type checking (strict mode)
- **Pytest** with >=90% coverage requirement; SVG/Chromium tests are real-browser (`pytest-asyncio`), and `tox` provisions Chromium via `playwright install chromium` before running them
- **Debian base image** (`debian:trixie-slim`, same as weasyprint-service) — required because Playwright has no musllinux wheel, so Alpine is not viable
- **Pandoc** and **tectonic** binaries in container, plus Playwright's bundled **Chromium** (see Dockerfile)

## API Endpoints
- `GET /health` - Health check (pandoc, tectonic, filesystem, and informational chromium status)
- `GET /version` - Service version information (python, pandoc, pandocService, timestamp, chromium)
- `GET /docx-template` - Download default DOCX template
- `GET /pptx-template` - Download default PPTX template
- `POST /convert/{source_format}/to/{target_format}` - General conversion (HTML sources accept a `scale_factor` query param)
- `POST /convert/{source_format}/to/docx-with-template` - DOCX with custom template (accepts `scale_factor`)
- `POST /convert/{source_format}/to/pptx-with-template` - PPTX with custom template (accepts `scale_factor`)

Service runs on port 9082 with health checks and comprehensive logging.

## CI/CD Workflows
- **Build & Release** (`.github/workflows/ci.yml`): Tests with tox, SonarCloud analysis, Hadolint, release-please, Docker build & publish to GHCR
- **Claude Code Review** (`.github/workflows/claude-code-review.yml`): Automated PR review (skips bot/fork PRs)
- **Add Issue to Project** (`.github/workflows/add-issue-to-project.yml`): Auto-adds new issues to GitHub project board
