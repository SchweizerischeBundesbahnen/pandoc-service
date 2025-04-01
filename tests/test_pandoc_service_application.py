import logging
import sys

from app import PandocServiceApplication


def test_main_runs(monkeypatch, tmp_path):
    """Test that main runs correctly with mocked dependencies."""

    # Mock command line arguments
    monkeypatch.setattr(sys, "argv", ["PandocServiceApplication.py", "--port", "9999"])

    # Set up fake server
    logger = logging.getLogger("test")

    def fake_start_server(port):
        logger.info(f"Fake server started on port {port}")

    monkeypatch.setattr(PandocServiceApplication.PandocController, "start_server", fake_start_server)

    # Run main and verify
    PandocServiceApplication.main()
