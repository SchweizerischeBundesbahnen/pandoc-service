# Testing Documentation

This directory contains all tests for the pandoc-service project. The testing is structured as follows:

## Directory Structure

- `tests/`: Main test directory
  - `data/`: Test data files (DOCX templates, example files, expected outputs)
  - `shell/`: Shell scripts for testing the Docker container
  - `*.py`: Python tests

## Test Types

### Python Tests

- `test_pandoc_controller.py`: Unit tests for the PandocController functionality
- `test_container.py`: Integration tests for the Docker container

### Shell Tests

- `tests/shell/test_pandoc_service.sh`: Smoke test for building and testing the Docker image
- `tests/shell/test_strictdoc_service.sh`: Tests for the StrictDoc service (for reference)

## Running Tests

### Python Tests

To run Python tests:

```bash
poetry run tox
poetry run pytest
```

For a specific test file:

```bash
poetry run pytest tests/test_pandoc_controller.py
```

### Docker Container Tests

To run the Docker container smoke test:

```bash
bash tests/shell/test_pandoc_service.sh
```

This script:
1. Builds the Docker image
2. Starts a container
3. Tests all endpoints (version, docx-template, convert, convert-with-template)
4. Verifies multiple format conversions
5. Cleans up resources afterwards

## Test Data

Test data files are located in the `tests/data/` directory. These include:
- Template DOCX files
- Test input files
- Expected output files for validation

When adding new tests, place test data files in this directory to keep tests organized.
