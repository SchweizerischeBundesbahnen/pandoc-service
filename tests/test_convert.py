import io
import logging
import unittest
from pathlib import Path

import requests
from docx import Document

from tests.arguments_parser import get_argument


class PandocServiceTestCase(unittest.TestCase):
    def setUp(self):
        self.base_url = get_argument(
            "APP_URL",
            True,
            "'app_url' is not provided: use --app_url or APP_URL env variable",
        )
        self.request_session = requests.Session()
        self.test_html = """
            <html>
                <body>
                    <h1>Simple html with an ordered list</h1>
                    <ol>
                        <li>First</li>
                        <li>Second</li>
                        <li>Third</li>
                    </ol>
                    <p>Some <b>bold German vowels ä, ö, and ü</b> at the bottom.</p>
                </body>
            </html>
            """

    def test_convert(self):
        response = self.__send_request(data=self.test_html, source_format="html", target_format="markdown", print_error=True)
        self.assertEqual(response.status_code, 200)
        with Path.open("test-data/expected-html-to-md.md", encoding="utf-8") as f:
            self.assertEqual(response.content.decode("utf-8"), f.read())

        response = self.__send_request(data=self.test_html, source_format="html", target_format="textile", print_error=True)
        self.assertEqual(response.status_code, 200)
        with Path.open("test-data/expected-html-to-textile.textile", encoding="utf-8") as f:
            self.assertEqual(response.content.decode("utf-8"), f.read())

        response = self.__send_request(data=self.test_html, source_format="html", target_format="plain", print_error=True)
        self.assertEqual(response.status_code, 200)
        with Path.open("test-data/expected-html-to-txt.txt", encoding="utf-8") as f:
            self.assertEqual(response.content.decode("utf-8"), f.read())

        with Path.open("test-data/test-input.docx", "rb") as i:
            response = self.__send_request(data=i.read(), source_format="docx", target_format="plain", print_error=True)
            self.assertEqual(response.status_code, 200)
            with Path.open("test-data/expected-docx-to-txt.txt", encoding="utf-8") as f:
                self.assertEqual(response.content.decode("utf-8"), f.read())

        response = self.__send_request(data=self.test_html, source_format="html", target_format="docx", print_error=True)
        self.assertEqual(response.status_code, 200)

        document = Document(io.BytesIO(response.content))

        paragraphs = []
        for paragraph in document.paragraphs:
            paragraphs.append(paragraph.text)

        self.assertListEqual(
            [
                "Simple html with an ordered list",
                "First",
                "Second",
                "Third",
                "Some bold German vowels ä, ö, and ü at the bottom.",
            ],
            paragraphs,
        )

    def __send_request(self, source_format, target_format, data, print_error):
        url = f"{self.base_url}/convert/{source_format}/to/{target_format}"
        try:
            response = self.request_session.request(method="POST", url=url, data=data, verify=True)
            if response.status_code // 100 != 2 and print_error:
                logging.error(f"Error: Unexpected response: '{response}'")
                logging.error(f"Error: Response content: '{response.content}'")
            return response
        except requests.exceptions.RequestException as e:
            logging.error(f"Error: {e}")
            raise
