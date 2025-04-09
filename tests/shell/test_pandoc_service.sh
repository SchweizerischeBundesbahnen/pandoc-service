#!/bin/bash
# Pandoc Service Docker Image Smoke Test
#
# This script performs a complete test of the pandoc-service Docker image by:
# - Building the Docker image
# - Starting a container
# - Testing all API endpoints
# - Verifying format conversions
# - Cleaning up resources
#
# Usage: bash test_pandoc_service.sh
#
# Author: Infrastructure Team
# Last updated: April 2025

set -e

# Define constants
CONTAINER_NAME="pandoc_service_test"
IMAGE_NAME="pandoc-service:test"
PORT=9082                       # Port used for connecting to the service
BASE_URL="http://localhost:${PORT}"

# Colors for output
GREEN="\033[0;32m"
RED="\033[0;31m"
YELLOW="\033[0;33m"
NC="\033[0m" # No Color

# Log function
log() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

# Error function
error() {
    echo -e "${RED}[ERROR]${NC} $1"
    # Show container logs if container exists
    if docker ps -q -f name=${CONTAINER_NAME} &>/dev/null; then
        echo -e "\nContainer logs:"
        docker logs ${CONTAINER_NAME}
    fi
    exit 1
}

# Warning function
warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

# Cleanup function
cleanup() {
    log "Cleaning up resources..."

    # Stop container if running
    if docker ps -q -f name=${CONTAINER_NAME} &>/dev/null; then
        log "Stopping container ${CONTAINER_NAME}..."
        docker stop ${CONTAINER_NAME} >/dev/null 2>&1 || warn "Failed to stop container"
    fi

    # Remove container if it exists
    if docker ps -aq -f name=${CONTAINER_NAME} &>/dev/null; then
        log "Removing container ${CONTAINER_NAME}..."
        docker rm -f ${CONTAINER_NAME} >/dev/null 2>&1 || warn "Failed to remove container"
    fi

    # Remove image if it exists
    if docker images ${IMAGE_NAME} -q &>/dev/null; then
        log "Removing image ${IMAGE_NAME}..."
        docker rmi -f ${IMAGE_NAME} >/dev/null 2>&1 || warn "Failed to remove image"
    fi

    log "Cleanup completed"
}

# Register cleanup on script exit
trap cleanup EXIT

# Initial cleanup to ensure clean state
cleanup

###########################################
# ENVIRONMENT CHECKS
###########################################

# Check if Docker is available
if ! command -v docker &> /dev/null; then
    error "Docker is not installed or not in PATH"
fi

# Check if curl is available
if ! command -v curl &> /dev/null; then
    error "curl is not installed or not in PATH"
fi

###########################################
# DOCKER IMAGE BUILD & RUN
###########################################

# Build Docker image
log "Building Docker image..."
if ! docker build -t ${IMAGE_NAME} . >/dev/null 2>&1; then
    error "Failed to build Docker image"
fi

# Run container
log "Starting container..."
if ! docker run -d --name ${CONTAINER_NAME} -p ${PORT}:9082 ${IMAGE_NAME} >/dev/null 2>&1; then
    error "Failed to start container"
fi

###########################################
# SERVICE HEALTH CHECK
###########################################

# Wait for container to be healthy
log "Waiting for container to be ready..."
attempt=1
max_attempts=10
until [ "$attempt" -gt "$max_attempts" ] || curl -s ${BASE_URL}/version >/dev/null 2>&1; do
    if ! docker ps -q -f name=${CONTAINER_NAME} >/dev/null 2>&1; then
        error "Container stopped unexpectedly. Check docker logs for details"
    fi
    log "Attempt $attempt/$max_attempts - Waiting for container to be ready..."
    sleep 5
    ((attempt++))
done

if [ "$attempt" -gt "$max_attempts" ]; then
    # Show container logs before failing
    echo -e "\nContainer logs:"
    docker logs ${CONTAINER_NAME}
    error "Container failed to become ready after $max_attempts attempts"
fi

log "Container is ready. Running tests..."

###########################################
# API ENDPOINT TESTS
###########################################

# Test 1: Check version endpoint
log "Test 1: Checking version endpoint..."
VERSION_RESPONSE=$(curl -s -o /dev/null -w "%{http_code}" ${BASE_URL}/version)
if [ "$VERSION_RESPONSE" -ne 200 ]; then
    error "Version endpoint returned non-200 status code: $VERSION_RESPONSE"
fi

VERSION_JSON=$(curl -s ${BASE_URL}/version)
log "Version info: $VERSION_JSON"

# Test 2: Check docx-template endpoint
log "Test 2: Checking docx-template endpoint..."
TEMPLATE_RESPONSE=$(curl -s -o /dev/null -w "%{http_code}" ${BASE_URL}/docx-template)
if [ "$TEMPLATE_RESPONSE" -ne 200 ]; then
    error "docx-template endpoint returned non-200 status code: $TEMPLATE_RESPONSE"
fi

###########################################
# TEST DATA PREPARATION
###########################################

# Create temporary directory for test files
TEST_DIR=$(mktemp -d)
MD_FILE="${TEST_DIR}/input.md"
HTML_FILE="${TEST_DIR}/output.html"
DOCX_FILE="${TEST_DIR}/output.docx"
TEMPLATE_FILE="${TEST_DIR}/template.docx"

# Download template for testing
log "Downloading template for testing..."
if ! curl -s -o "${TEMPLATE_FILE}" ${BASE_URL}/docx-template; then
    error "Failed to download template file"
fi

# Create test markdown file
cat > "${MD_FILE}" << 'EOF'
# Test Document

## Section 1

This is a test paragraph.

- List item 1
- List item 2
- List item 3

## Section 2

Here is a table:

| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |

EOF

###########################################
# FORMAT CONVERSION TESTS
###########################################

# Test 3: Convert markdown to HTML
log "Test 3: Converting markdown to HTML..."
CONVERT_RESPONSE=$(curl -s -o "${HTML_FILE}" -w "%{http_code}" \
    -X POST \
    -H "Content-Type: text/plain" \
    --data-binary "@${MD_FILE}" \
    "${BASE_URL}/convert/markdown/to/html")

if [ "$CONVERT_RESPONSE" -ne 200 ]; then
    error "Convert endpoint returned non-200 status code: $CONVERT_RESPONSE"
fi

log "HTML conversion successful. Response code: $CONVERT_RESPONSE"

# Check HTML file size
HTML_SIZE=$(wc -c < "${HTML_FILE}")
if [ "${HTML_SIZE}" -lt 100 ]; then
    error "HTML output is too small (${HTML_SIZE} bytes)"
fi
log "HTML output size: ${HTML_SIZE} bytes"

# Test 4: Convert markdown to DOCX
log "Test 4: Converting markdown to DOCX..."
CONVERT_RESPONSE=$(curl -s -o "${DOCX_FILE}" -w "%{http_code}" \
    -X POST \
    -H "Content-Type: text/plain" \
    --data-binary "@${MD_FILE}" \
    "${BASE_URL}/convert/markdown/to/docx")

if [ "$CONVERT_RESPONSE" -ne 200 ]; then
    error "Convert endpoint returned non-200 status code: $CONVERT_RESPONSE"
fi

log "DOCX conversion successful. Response code: $CONVERT_RESPONSE"

# Check DOCX file size
DOCX_SIZE=$(wc -c < "${DOCX_FILE}")
if [ "${DOCX_SIZE}" -lt 1000 ]; then
    error "DOCX output is too small (${DOCX_SIZE} bytes)"
fi
log "DOCX output size: ${DOCX_SIZE} bytes"

###########################################
# TEMPLATE TESTS
###########################################

# Test 5: Convert with template
log "Test 5: Converting with template..."
TEMPLATE_CONVERT_RESPONSE=$(curl -s -o "${TEST_DIR}/template_output.docx" -w "%{http_code}" \
    -X POST \
    -F "source=@${MD_FILE}" \
    -F "template=@${TEMPLATE_FILE}" \
    "${BASE_URL}/convert/markdown/to/docx-with-template")

if [ "$TEMPLATE_CONVERT_RESPONSE" -ne 200 ]; then
    error "Template convert endpoint returned non-200 status code: $TEMPLATE_CONVERT_RESPONSE"
fi

log "Template conversion successful. Response code: $TEMPLATE_CONVERT_RESPONSE"

# Check template output file size
TEMPLATE_OUTPUT_SIZE=$(wc -c < "${TEST_DIR}/template_output.docx")
if [ "${TEMPLATE_OUTPUT_SIZE}" -lt 1000 ]; then
    error "Template output is too small (${TEMPLATE_OUTPUT_SIZE} bytes)"
fi
log "Template output size: ${TEMPLATE_OUTPUT_SIZE} bytes"

###########################################
# TEST DATA FILE TESTS
###########################################

# Test with various test files from the tests/data directory
log "Test 6: Testing with files from tests/data directory..."

# Test HTML to Markdown conversion
HTML_TO_MD=$(curl -s -o "${TEST_DIR}/output.md" -w "%{http_code}" \
    -X POST \
    -H "Content-Type: text/html" \
    --data-binary @"tests/data/test-input.docx" \
    "${BASE_URL}/convert/html/to/markdown")

if [ "$HTML_TO_MD" -ne 200 ]; then
    error "HTML to Markdown conversion failed with status code: $HTML_TO_MD"
fi
log "HTML to Markdown conversion successful"

# Test with custom template file
TEMPLATE_TEST=$(curl -s -o "${TEST_DIR}/template_test.docx" -w "%{http_code}" \
    -X POST \
    -F "source=@${MD_FILE}" \
    -F "template=@tests/data/template-red.docx" \
    "${BASE_URL}/convert/markdown/to/docx-with-template")

if [ "$TEMPLATE_TEST" -ne 200 ]; then
    error "Template test failed with status code: $TEMPLATE_TEST"
fi
log "Template test successful"

###########################################
# CLEANUP AND COMPLETION
###########################################

# Clean up test directory
rm -rf "${TEST_DIR}"

log "All tests passed successfully!"

# Show docker logs for debugging
log "Container logs:"
docker logs ${CONTAINER_NAME}

log "Test completed successfully!"
