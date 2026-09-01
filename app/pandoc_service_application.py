import argparse
import logging
import os
from datetime import UTC, datetime
from pathlib import Path

from app import pandoc_controller
from app.constants import get_graceful_shutdown_timeout

logger = logging.getLogger(__name__)


def configure_uvicorn_logging(level: int) -> None:
    """
    Route the messages uvicorn writes itself through the handlers of the root logger.

    uvicorn ships a logging configuration which gives its loggers a handler of their
    own and stops them from propagating. Two problems follow: the messages never reach
    the log file, and the level is global, so the metrics server sets the level of the
    main server as well. Both servers are started with ``log_config=None``, which leaves
    the loggers to this function.

    Access logging stays off below DEBUG. The Docker healthcheck calls /version every 30
    seconds and Prometheus scrapes /metrics, so a line per request buries the rest.

    Args:
        level: The level configured through LOG_LEVEL.
    """
    for logger_name in ("uvicorn", "uvicorn.error", "uvicorn.asgi"):
        uvicorn_logger = logging.getLogger(logger_name)
        uvicorn_logger.handlers.clear()
        uvicorn_logger.propagate = True
        uvicorn_logger.setLevel(level)

    access_logger = logging.getLogger("uvicorn.access")
    access_logger.handlers.clear()
    access_logger.setLevel(level)
    access_logger.propagate = level <= logging.DEBUG


def setup_logging() -> Path:
    """
    Configure logging for the Pandoc service with both file and console output.

    The function:
    - Sets log level from LOG_LEVEL environment variable (defaults to INFO)
    - Creates timestamped log files in /opt/pandoc/logs directory
    - Configures both file and console logging handlers
    - Uses format: timestamp - logger name - log level - message

    The log files are not rotated and a new file is created on each service start.

    Returns:
        Path: The path to the created log file
    """
    # Clean up any existing handlers
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        handler.close()
        root_logger.removeHandler(handler)

    log_level = os.getenv("LOG_LEVEL", "INFO").upper()
    # Use LOG_DIR environment variable if set, otherwise use default
    log_dir = Path(os.getenv("LOG_DIR", "/opt/pandoc/logs"))
    log_dir.mkdir(parents=True, exist_ok=True)

    # Create log filename with timestamp
    current_time = datetime.now(tz=UTC).strftime("%Y-%m-%d_%H-%M-%S")
    log_file = log_dir / f"pandoc-service_{current_time}.log"

    # Configure logging format
    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")

    # Configure file handler (no rotation)
    file_handler = logging.FileHandler(
        log_file,
        encoding="utf-8",
        delay=False,  # Create file immediately
    )
    file_handler.setFormatter(formatter)

    # Configure console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    # Setup root logger
    configured_level = getattr(logging, log_level, logging.INFO)  # Default to INFO if invalid
    root_logger.setLevel(configured_level)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    configure_uvicorn_logging(configured_level)

    # Force immediate file creation
    root_logger.info(f"Logging initialized with level: {log_level}")
    root_logger.info(f"Log file: {log_file}")

    # Ensure everything is written
    for handler in root_logger.handlers:
        handler.flush()

    return log_file  # Return log file path for testing


def main() -> None:
    """
    Main entry point for the Pandoc service.

    Parses command line arguments, initializes logging, and starts the server.
    The service port can be specified via command line argument (defaults to 9082).

    The metrics server lifecycle is managed by FastAPI's lifespan context manager
    in pandoc_controller, ensuring proper startup and cleanup.
    """
    parser = argparse.ArgumentParser(description="Pandoc service")
    parser.add_argument("--port", default=9082, type=int, required=False, help="Service port")
    args = parser.parse_args()

    setup_logging()
    logger.info("Pandoc service listening port: %d", args.port)
    logger.info("Graceful shutdown timeout: %d seconds", get_graceful_shutdown_timeout())

    # Start the server - metrics server lifecycle is managed by FastAPI lifespan
    pandoc_controller.start_server(args.port)


if __name__ == "__main__":
    main()
