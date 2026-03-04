#!/bin/bash
set -e

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${BLUE}Generating test load for Pandoc Service...${NC}"
echo ""

# Configuration
BASE_URL="http://localhost:9082"
REQUESTS=${1:-100}
CONCURRENCY=${2:-10}

echo -e "${YELLOW}Configuration:${NC}"
echo -e "  Requests:    ${REQUESTS}"
echo -e "  Concurrency: ${CONCURRENCY}"
echo ""

# Check if service is running
if ! curl -sf ${BASE_URL}/version > /dev/null 2>&1; then
    echo -e "${RED}Error: Pandoc service is not running at ${BASE_URL}${NC}"
    echo -e "${YELLOW}Run ./monitoring/start-monitoring.sh first${NC}"
    exit 1
fi

# Function to send markdown to HTML conversion request
send_markdown_request() {
    local id=$1
    curl -s -X POST ${BASE_URL}/convert/markdown/to/html \
         -H "Content-Type: text/plain" \
         -d "# Test Document $id

This is a test document generated at $(date).

## Features
- Item 1
- Item 2
- Item 3

**Bold text** and *italic text*." \
         -o /dev/null
}

# Function to send HTML to DOCX conversion request
send_docx_request() {
    local id=$1
    curl -s -X POST ${BASE_URL}/convert/html/to/docx \
         -H "Content-Type: text/html" \
         -d "<html><body><h1>Test Document $id</h1><p>Generated at $(date)</p></body></html>" \
         -o /dev/null
}

# Generate mixed load
echo -e "${YELLOW}Generating ${REQUESTS} requests...${NC}"
echo -n "Progress: "

count=0
pids=()

for i in $(seq 1 $REQUESTS); do
    # Mix of different conversion types
    if [ $((i % 3)) -eq 0 ]; then
        send_docx_request $i &
    else
        send_markdown_request $i &
    fi

    pids+=($!)
    count=$((count + 1))

    # Limit concurrency
    if [ ${#pids[@]} -ge $CONCURRENCY ]; then
        wait ${pids[0]}
        pids=("${pids[@]:1}")
    fi

    # Progress indicator
    if [ $((count % 10)) -eq 0 ]; then
        echo -n "."
    fi
done

# Wait for remaining requests
for pid in "${pids[@]}"; do
    wait $pid
done

echo -e " ${GREEN}Done!${NC}"
echo ""
echo -e "${GREEN}Generated ${REQUESTS} requests${NC}"
echo ""
echo -e "${BLUE}Check metrics at:${NC}"
echo -e "  • Prometheus:        ${YELLOW}http://localhost:9090/graph${NC}"
echo -e "  • Grafana:           ${YELLOW}http://localhost:3000/d/pandoc-service${NC}"
echo ""
