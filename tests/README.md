# Automated regression testing

GitHub Actions runs `tests/regression.py` on every push to `main` and on pull requests.

## LA-350

For the complete LA-350 regression test, place the original interactive reference PDF at:

`tests/fixtures/LA-350.pdf`

The test then:

- verifies the PDF is detected as an interactive LA-350 form;
- generates the native editable DOCX using the production converter;
- checks required form content;
- renders the DOCX back to PDF with LibreOffice;
- requires the result to remain one US-letter page; and
- uploads the generated DOCX, PDF, and PNG render as a GitHub Actions artifact.

Until the reference PDF is committed, the workflow still verifies that the application and test harness import successfully, but skips the full LA-350 conversion.
