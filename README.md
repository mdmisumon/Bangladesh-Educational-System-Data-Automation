# Bangladesh Educational System Data Automation

Automates Bangladesh Education Board result lookups and writes the extracted student information back into an Excel workbook.

The current script is built for the newer Education Board result site:

```text
https://www.educationboardresults.gov.bd/v2/home
```

It reads student lookup data from `Result data.xlsx`, prompts you to solve the website CAPTCHA manually, submits the lookup to the current `/v2/getres` endpoint, and saves the returned result data into the same workbook.

## What It Extracts

For each valid row, the script fills these output columns:

- `Name`
- `Father's Name`
- `Mother's Name`
- `GPA`
- `Date of Birth`

## Workbook Format

The workbook must be named:

```text
Result data.xlsx
```

By default, the script looks for this file in the folder where the script is run. If it is not found there, it looks beside `Result Data.py`.

The first row must contain these headers:

| Column | Purpose |
| --- | --- |
| `Exam` | Exam type, for example `SSC/Dakhil/Equivalent` |
| `Passing Year` | Exam year, for example `2015` |
| `Board` | Board name, for example `Dhaka` |
| `Roll` | Student roll number |
| `Reg` | Student registration number |
| `Name` | Output field |
| `Father's Name` | Output field |
| `Mother's Name` | Output field |
| `GPA` | Output field |
| `Date of Birth` | Output field |

Rows with existing successful `Name` values are skipped by default. Use `--force` to reprocess them.

## Requirements

Install the Python packages used by the script:

```powershell
pip install openpyxl requests beautifulsoup4
```

## Usage

Run from the project folder:

```powershell
python "Result Data.py"
```

To reprocess rows that already have output values:

```powershell
python "Result Data.py" --force
```

To use a workbook in another location:

```powershell
python "Result Data.py" --file "C:\path\to\Result data.xlsx"
```

To save CAPTCHA images without opening them automatically:

```powershell
python "Result Data.py" --no-open-captcha
```

## CAPTCHA Workflow

The Education Board website now uses an image CAPTCHA. The script does not bypass this protection.

For each row that needs processing, the script:

1. Opens the result homepage to create a valid session.
2. Downloads the CAPTCHA image.
3. Saves it to `.captchas/`.
4. Opens the image for you.
5. Prompts you to type the visible digits.
6. Submits the result lookup.

If the CAPTCHA is incorrect, the script retries up to three times by default.

You can change the attempt count:

```powershell
python "Result Data.py" --max-captcha-attempts 5
```

## Useful Options

```text
--file PATH                 Use a specific workbook path.
--force                     Reprocess rows that already contain a successful Name value.
--captcha-dir PATH          Save CAPTCHA images in a custom folder.
--no-open-captcha           Do not open CAPTCHA images automatically.
--max-captcha-attempts N    Set CAPTCHA retry limit per row. Default: 3.
--delay SECONDS             Delay between successful lookups. Default: 2.
```

## Notes

- Keep `Result data.xlsx` closed while the script is saving results.
- Generated CAPTCHA images and Python cache files are ignored by Git.
- The `website structure/` folder is local reference material and is not required to run the script.
- The script supports the current `/v2` result flow and may need updates if the Education Board website changes again.
