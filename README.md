# Bangladesh Educational System Data Automation

Automates Bangladesh Education Board result lookups and writes student result data back into an Excel workbook.

The main script is:

```text
Result Data V2.py
```

Current script version:

```text
2.2.0
```

It supports regular education-board results from:

```text
https://www.educationboardresults.gov.bd/v2/home
```

For Technical board rows, it uses the Bangladesh Technical Education Board result service:

```text
https://result.bteb.gov.bd/result-search
```

## What It Extracts

For each valid row, the script fills these output columns:

- `Name`
- `Father's Name`
- `Mother's Name`
- `GPA`
- `Date of Birth`

Date of birth values are normalized to Bangladesh format:

```text
dd-mm-yyyy
```

The DOB cell is written as text so Excel does not convert it into a US-style date or another automatic date format.

## Workbook Format

By default, V2 looks for:

```text
Result data.xlsx
```

The first row must contain these headers:

| Column | Purpose |
| --- | --- |
| `Exam` | Exam type, for example `SSC/Dakhil/Equivalent` or `SSC/Dakhil (VOC)` |
| `Passing Year` | Exam year, for example `2019` |
| `Board` | Board name, for example `Dhaka` or `Technical` |
| `Roll` | Student roll number |
| `Reg` | Student registration number |
| `Name` | Output field |
| `Father's Name` | Output field |
| `Mother's Name` | Output field |
| `GPA` | Output field |
| `Date of Birth` | Output field |

Rows with any existing `Name` value are skipped by default, including rows marked with an error such as `Result Error`. Clear the `Name` cell for a row if you want the script to check it again. Use `--force` to reprocess rows even when `Name` is already filled.

## Technical Board Rows

Rows where `Board` is `Technical` or `TEC` are sent to the BTEB public result service instead of the regular education-board endpoint.

For common SSC/Dakhil vocational rows, the script can infer:

- `27` for `SSC (Vocational)`
- `77` for `Dakhil (Vocational)`
- semester/class `2` for Class 10

If inference is not enough, add optional columns to the workbook:

| Optional Column | Example |
| --- | --- |
| `Curriculum Code` | `27` |
| `Semester` or `Class` | `2` or `Class 10` |

## Requirements

Install the Python packages used by the script:

```powershell
pip install openpyxl requests beautifulsoup4
```

The regular education-board site may show an automatic browser verification page before the result form. To let the script open a real browser for that verification and then continue the normal CAPTCHA flow, install Playwright:

```powershell
pip install playwright
python -m playwright install chromium
```

Optional Gemini CAPTCHA support requires:

```powershell
pip install -U google-genai pillow
```

## Usage

Run from the project folder:

```powershell
python "Result Data V2.py"
```

To reprocess rows that already have output values:

```powershell
python "Result Data V2.py" --force
```

To print the script version:

```powershell
python "Result Data V2.py" --version
```

To use a workbook in another location:

```powershell
python "Result Data V2.py" --file "C:\path\to\Result data.xlsx"
```

To provide a Gemini API key:

```powershell
python "Result Data V2.py" --gemini-api-key "YOUR_KEY"
```

Or set an environment variable:

```powershell
$env:GEMINI_API_KEY="YOUR_KEY"
python "Result Data V2.py"
```

## CAPTCHA Workflow

For regular education-board rows, the site may first show an automatic browser verification page. When that happens, the script opens a browser window, waits for the real result form, copies the verified session cookies back into the script, and then continues.

The regular site also uses an image CAPTCHA. If no Gemini API key is provided, the script uses a manual flow:

1. Opens the result homepage to create a valid session.
2. Completes browser verification if the site requires it.
3. Downloads the CAPTCHA image.
4. Saves it to `.captchas/`.
5. Opens the image for you.
6. Prompts you to type the visible digits.
7. Submits the result lookup.

If the CAPTCHA is incorrect, the script retries up to three times by default.

Technical-board rows use the BTEB public result API and do not use this CAPTCHA flow.

## Useful Options

```text
--file PATH                 Use a specific workbook path.
--gemini-api-key KEY        Use Gemini API for CAPTCHA OCR.
--force                     Reprocess rows that already contain a successful Name value.
--captcha-dir PATH          Save CAPTCHA images in a custom folder.
--no-open-captcha           Do not open CAPTCHA images automatically.
--no-browser-verify         Disable the browser verification bootstrap.
--browser-verify-timeout N  Seconds to wait for browser verification. Default: 45.
--headless-browser-verify   Run browser verification without showing a browser window.
--no-pause                  Exit immediately instead of waiting for Enter at the end.
--no-enter-pause            Disable the Enter key pause/resume checkpoint during processing.
--max-captcha-attempts N    Set CAPTCHA retry limit per row. Default: 3.
--delay SECONDS             Delay between successful lookups. Default: 2.
```

## Terminal Behavior

Results are saved to the workbook immediately after each row writes either extracted data or an error status. Progress is not held until the end of the full run.

While the script is processing rows, press Enter to save current progress and pause at the next safe checkpoint. Press Enter again to resume. Use `--no-enter-pause` to disable this behavior.

The script waits at the end with:

```text
Finished. Press Enter to close this window...
```

This keeps the terminal open after a double-click run so you can read the result summary. Use `--no-pause` to disable this behavior.

## Notes

- Keep the workbook closed while the script is saving results.
- Generated CAPTCHA images and Python cache files are ignored by Git.
- Website-structure folders are local reference material and are not required to run the script.
- The script supports the current education-board `/v2` flow and the current BTEB public result flow. It may need updates if either website changes again.
