# Test Data

This directory contains data files used for testing the pandoc-service.

## Files

- `template-red.docx` - DOCX template with red headings for testing custom templates
- `test-input.docx` - Sample DOCX file for testing document conversions
- `expected-docx-to-txt.txt` - Expected output when converting DOCX to text
- `expected-html-to-md.md` - Expected output when converting HTML to Markdown
- `expected-html-to-textile.textile` - Expected output when converting HTML to Textile
- `expected-html-to-txt.txt` - Expected output when converting HTML to text
- `big_image_in_base64.txt` - Base64-encoded image for testing large file handling

## Usage

These files are used by both Python tests and shell scripts:

- Python tests use them to verify conversion results (`test_container.py`)
- Shell scripts use them for testing Docker container functionality (`test_pandoc_service.sh`)

## Adding New Test Data

When adding new test data:

1. Use descriptive filenames that indicate the file's purpose
2. For expected outputs, use the naming pattern `expected-[source]-to-[target].[ext]`
3. For templates, use a descriptive name that indicates its special properties
4. Update this README with information about the new files
5. Consider adding a small comment at the top of text files to explain their purpose
