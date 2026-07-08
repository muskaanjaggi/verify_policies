# Policy PDF Verifier

A Streamlit app that verifies insurance policy details against their PDF documents.
Upload an Excel file with policy numbers, dates, and PDF links — the app downloads
each PDF, checks the policy number and start/end dates against the document text,
and returns a color-coded flagged Excel file you can download. 
Deployed Link:
https://verifypolicies.streamlit.app/

## What it checks

For each row in your Excel file, the app:

1. Downloads the PDF linked in the `link` column
2. Extracts the text (with OCR fallback for scanned PDFs, if enabled)
3. Verifies:
   - **Policy number** (`number` column) appears in the PDF — handles Excel
     distortions like scientific notation, trailing `.0`, and precision loss
   - **Start date** (`start_date` column) appears in the PDF, in any common
     date format
   - **End date** (`end_date` column) appears in the PDF, in any common
     date format
4. Marks each row `PASS`, `FLAG` (with the specific mismatch listed), or
   `ERROR` (with the exact download/parse failure reason)

## Required Excel columns

| Column       | Description                              |
|--------------|-------------------------------------------|
| `number`     | Policy number to verify                   |
| `link`       | Direct URL to the policy PDF              |
| `start_date` | Expected policy start date                |
| `end_date`   | Expected policy end date                  |

Column names are case-sensitive. Any other columns in your sheet are kept
as-is in the output.

## Output

The downloaded file (`flagged.xlsx`) contains all your original columns plus:

- **Verification_Status** — `PASS`, `FLAG: <reason(s)>`, or `ERROR: <reason>`
- **Download_Error_Detail** — the exact cause of a download/parse failure
  (timeout, expired link, encrypted PDF, scanned PDF with no text layer, etc.)

Rows are color-coded: green = pass, red = flagged, yellow = download error.

## Deploying on Streamlit Community Cloud

1. Push this repo (`streamlit_app.py`, `requirements.txt`, `packages.txt`) to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect the repo.
3. Set the main file path to `streamlit_app.py`, then deploy.

`packages.txt` installs `tesseract-ocr` and `poppler-utils` at the system
level, which power the OCR fallback for scanned/image-only PDFs. If you don't
need OCR, you can delete `packages.txt` and remove `pytesseract` and
`pdf2image` from `requirements.txt` for a lighter, faster deploy.

## Notes and limitations

- **Large batches take time.** Each PDF download + parse takes roughly
  0.5–2 seconds depending on file size and network speed, plus retry backoff
  on failures. A few hundred rows can take several minutes.
- **Streamlit Cloud has execution limits.** Free-tier apps sleep after
  inactivity and may time out on very large batches run in a single click.
  For large files, consider splitting the upload into smaller chunks.
- **Signed/expiring URLs.** If your `link` column contains time-limited
  signed URLs (e.g. S3 presigned links), make sure they haven't expired
  before uploading — an expired link shows up as an `HTTP 403` or similar
  in `Download_Error_Detail`.
- **OCR is a fallback, not primary extraction.** It only kicks in on pages
  where fewer than ~20 characters of text are found, to catch scanned PDFs
  without slowing down normal text-based PDFs.
