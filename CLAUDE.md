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
- **app/schema.py**: Pydantic models for API responses

### Security Model
- Uses allowlisted pandoc options to prevent command injection
- Direct subprocess calls to pandoc binary (no shell=True)
- Input validation with 200MB request size limit
- Format validation for source/target combinations

### Conversion Pipeline
1. Request validation (format, size, encoding)
2. Pandoc subprocess execution with security constraints
3. Post-processing (especially for DOCX files)
4. Response formatting with appropriate MIME types

### Supported Formats
- **Source**: docx, epub, fb2, html, json, latex, markdown, rtf, textile
- **Target**: docx, epub, fb2, html, json, latex, markdown, odt, pdf, plain, pptx, rtf, textile

### Key Configuration
- **Python 3.14** required (see `.tool-versions` for exact patch version)
- **uv** for dependency management (`--frozen` flag used in CI)
- **Ruff** for linting (line length: 240, TCH rule enforces `if TYPE_CHECKING:` import guards)
- **Mypy** for type checking (strict mode)
- **Pytest** with >=90% coverage requirement
- **Pandoc v3.9** binary in container

## API Endpoints
- `GET /health` - Health check (pandoc, tectonic, filesystem status)
- `GET /version` - Service version information
- `GET /docx-template` - Download default DOCX template
- `GET /pptx-template` - Download default PPTX template
- `POST /convert/{source_format}/to/{target_format}` - General conversion
- `POST /convert/{source_format}/to/docx-with-template` - DOCX with custom template
- `POST /convert/{source_format}/to/pptx-with-template` - PPTX with custom template

Service runs on port 9082 with health checks and comprehensive logging.

## CI/CD Workflows
- **Build & Release** (`.github/workflows/ci.yml`): Tests with tox, SonarCloud analysis, Hadolint, release-please, Docker build & publish to GHCR
- **Claude Code Review** (`.github/workflows/claude-code-review.yml`): Automated PR review (skips bot/fork PRs)
- **Add Issue to Project** (`.github/workflows/add-issue-to-project.yml`): Auto-adds new issues to GitHub project board
