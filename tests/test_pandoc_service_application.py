import logging
import sys

from app import pandoc_service_application


def test_main_runs(monkeypatch, tmp_path):
    """Test that main runs correctly with mocked dependencies."""

    # Set up temporary log directory
    log_dir = tmp_path / "logs"
    monkeypatch.setenv("LOG_DIR", str(log_dir))

    # Mock command line arguments
    monkeypatch.setattr(sys, "argv", ["pandoc_service_application.py", "--port", "9999"])

    # Set up fake server
    logger = logging.getLogger("test")

    def fake_start_server(port):
        logger.info(f"Fake server started on port {port}")

    monkeypatch.setattr(pandoc_service_application.pandoc_controller, "start_server", fake_start_server)

    # Run main and verify
    pandoc_service_application.main()

    # Verify log directory was created
    assert log_dir.exists()
    assert any(log_dir.glob("pandoc-service_*.log"))


def test_uvicorn_loggers_reach_the_root_handlers(monkeypatch, tmp_path):
    """uvicorn writes through the root logger, so its messages land in the log file."""
    monkeypatch.setenv("LOG_DIR", str(tmp_path / "logs"))

    log_file = pandoc_service_application.setup_logging()

    uvicorn_logger = logging.getLogger("uvicorn.error")
    assert uvicorn_logger.handlers == []
    assert uvicorn_logger.propagate is True
    assert uvicorn_logger.level == logging.INFO

    uvicorn_logger.info("Started server process [1]")
    for handler in logging.getLogger().handlers:
        handler.flush()

    assert "Started server process [1]" in log_file.read_text(encoding="utf-8")


def test_access_logging_is_off_above_debug(monkeypatch, tmp_path):
    """Per-request lines stay out of the log: the healthcheck and Prometheus poll constantly."""
    monkeypatch.setenv("LOG_DIR", str(tmp_path / "logs"))

    pandoc_service_application.setup_logging()

    access_logger = logging.getLogger("uvicorn.access")
    assert access_logger.handlers == []
    assert access_logger.propagate is False


def test_access_logging_follows_debug_level(monkeypatch, tmp_path):
    """DEBUG asks for everything, so the per-request lines are let through."""
    monkeypatch.setenv("LOG_DIR", str(tmp_path / "logs"))
    monkeypatch.setenv("LOG_LEVEL", "DEBUG")

    pandoc_service_application.setup_logging()

    assert logging.getLogger("uvicorn.access").propagate is True


def test_start_server_passes_the_graceful_shutdown_timeout(monkeypatch):
    """start_server hands the timeout to uvicorn, which applies it on SIGTERM."""
    monkeypatch.setenv("GRACEFUL_SHUTDOWN_TIMEOUT", "12")
    captured = {}

    def fake_run(**kwargs):
        captured.update(kwargs)

    monkeypatch.setattr(pandoc_service_application.pandoc_controller.uvicorn, "run", fake_run)

    pandoc_service_application.pandoc_controller.start_server(9999)

    assert captured["timeout_graceful_shutdown"] == 12
    assert captured["log_config"] is None
    assert captured["port"] == 9999
