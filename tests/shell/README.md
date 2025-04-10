# Shell Tests

This directory contains shell scripts for testing Docker containers.

## Available Tests

### pandoc-service Test

`test_pandoc_service.sh` is a smoke test for the pandoc-service Docker image. It:

- Builds the Docker image
- Starts a container
- Tests all endpoints (version, docx-template, convert, convert-with-template)
- Verifies format conversions (markdown to HTML, markdown to DOCX, HTML to markdown)
- Tests with custom templates
- Validates responses

### Running Tests

To run the pandoc-service smoke test:

```bash
bash test_pandoc_service.sh
```

The script will output progress information and clean up all Docker resources when finished.

### Test Script Structure

The script uses the following structure:

1. **Setup and Cleanup**: Manages Docker resources and performs cleanup on exit
2. **Environment Checks**: Verifies Docker and curl are available
3. **Container Management**: Builds and starts the container
4. **Health Checks**: Verifies the container is ready
5. **Endpoint Tests**: Tests each API endpoint
6. **Format Conversion Tests**: Tests various format conversions
7. **Template Tests**: Tests with custom DOCX templates
8. **Cleanup**: Removes temporary files and Docker resources

### Adding New Tests

When adding new tests:

1. Use consistent logging (log/error/warn functions)
2. Check response status codes
3. Validate response content where appropriate
4. Clean up any temporary files
5. Add meaningful error messages
6. Use test data from tests/data directory
