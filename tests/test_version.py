import unittest

import requests

from tests.arguments_parser import get_argument


class PandocServiceTestCase(unittest.TestCase):
    def setUp(self):
        self.base_url = get_argument(
            "APP_URL",
            True,
            "'app_url' is not provided: use --app_url or APP_URL env variable",
        )
        self.request_session = requests.Session()

    def test_convert_simple_html(self):
        url = f"{self.base_url}/version"
        response = self.request_session.request(method="GET", url=url)

        self.assertEqual(response.status_code, 200)
        json_settings = response.json()
        self.assertIsInstance(json_settings["python"], str)
        self.assertIsInstance(json_settings["pandoc"], str)
        self.assertIsInstance(json_settings["pandocService"], str)
        self.assertIsInstance(json_settings["timestamp"], str)
