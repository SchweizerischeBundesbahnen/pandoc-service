import argparse
import logging

from app import PandocController


def main() -> None:
    """
    Main entry point for the Pandoc service.

    Parses command line arguments, initializes logging, and starts the server.
    The service port can be specified via command line argument (defaults to 9082).
    """
    parser = argparse.ArgumentParser(description="Pandoc service")
    parser.add_argument("--port", default=9082, type=int, required=False, help="Service port")
    args = parser.parse_args()

    logging.getLogger().setLevel(logging.INFO)
    logging.info("Pandoc service listening port: " + str(args.port))
    logging.getLogger().setLevel(logging.WARN)

    PandocController.start_server(args.port)


if __name__ == "__main__":
    main()
