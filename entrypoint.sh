#!/bin/bash

BUILD_TIMESTAMP="$(cat /opt/pandoc/.build_timestamp)"
export PANDOC_SERVICE_BUILD_TIMESTAMP=${BUILD_TIMESTAMP}

# Also possible with:
# source .venv/bin/activate
# python -m app.PandocServiceApplication &
uv run python -m app.PandocServiceApplication &

wait -n

exit $?
