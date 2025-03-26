import io
import logging
import unittest
from pathlib import Path

import requests
from docx import Document
from docx.shared import RGBColor

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
                <head>
                    <title>Test doc title</title>
                </head>
                <body>
                    <h1>Simple html with several headings</h1>
                    <p>Some content 1</p>
                    <h2>Second heading with German vowels ä, ö, and ü</h2>
                    <p>Some content 2</p>
                    <h3>Third</h3>
                    <p>Some content 3</p>
                </body>
            </html>

        """

    def test_convert_with_docx_template(self):
        # First test without template - it has some default headings color
        response = self.__send_docx_with_template_request(data=self.test_html, source_format="html", print_error=True)
        self.assertDocContainsSpecificHeadersColor(RGBColor(15, 71, 97), response.content)

        # Now test with 'RED' template - it forces red color for headings
        with Path.open("test-data/template-red.docx", "rb") as t:
            template = t.read()
        response = self.__send_docx_with_template_request(data=self.test_html, template=template, source_format="html", print_error=True)
        self.assertDocContainsSpecificHeadersColor(RGBColor(255, 0, 0), response.content)

    def assertDocContainsSpecificHeadersColor(self, color, doc_content):
        document = Document(io.BytesIO(doc_content))

        # Check for specific headings colors and extract their text
        headings = []
        for paragraph in document.paragraphs:
            if paragraph.style.style_id.startswith("Heading"):
                self.assertTrue(color in {paragraph.style.base_style.font.color.rgb, paragraph.style.font.color.rgb})
                headings.append(paragraph.text.replace("\xa0", " "))

        self.assertListEqual(
            [
                "Simple html with several headings",
                "Second heading with German vowels ä, ö, and ü",
                "Third",
            ],
            headings,
        )

    def __send_docx_with_template_request(self, source_format, data, template=None, print_error=False):
        url = f"{self.base_url}/convert/{source_format}/to/docx-with-template"
        files = {"source": ("file.html", data)}
        if template:
            files["template"] = ("template.docx", template)
        try:
            response = self.request_session.request(method="POST", url=url, files=files, verify=True)
            if response.status_code // 100 != 2 and print_error:
                logging.error(f"Error: Unexpected response: '{response}'")
                logging.error(f"Error: Response content: '{response.content}'")
            return response
        except requests.exceptions.RequestException as e:
            logging.error(f"Error: {e}")
            raise
