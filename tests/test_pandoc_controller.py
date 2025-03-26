import os
import platform

from app.PandocController import app


def test_version():
    os.environ["PANDOC_SERVICE_VERSION"] = "test1"
    os.environ["PANDOC_SERVICE_BUILD_TIMESTAMP"] = "test2"
    with app.test_client() as test_client:
        version = test_client.get("/version").json

        assert version["python"] == platform.python_version()
        assert version["pandoc"] is not None
        assert version["pandocService"] == "test1"
        assert version["timestamp"] == "test2"
