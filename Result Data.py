import openpyxl
import requests
from bs4 import BeautifulSoup
import time
import re # For regular expressions to parse captcha

# --- Configuration ---
# IMPORTANT: Use a raw string (r'') for Windows paths to avoid issues with backslashes
excel_file_path = r'C:\Result data.xlsx' #Your File path Here
base_url = "http://www.educationboardresults.gov.bd/"
post_url = "http://www.educationboardresults.gov.bd/result.php" # The form's action URL

# Mapping for Exam types from Excel (left) to Website's dropdown values (right)
EXAM_TYPE_MAPPING = {
    'SSC/Dakhil/Equivalent': 'ssc',
    'JSC/JDC': 'jsc',
    'SSC/Dakhil': 'ssc',
    'SSC(Vocational)': 'ssc_voc',
    'HSC/Alim': 'hsc',
    'HSC(Vocational)': 'hsc_voc',
    'HSC(BM)': 'hsc_hbm',
    'Diploma in Commerce': 'hsc_dic',
    'Diploma in Business Studies': 'hsc',
}

# Mapping for Board names from Excel (left) to Website's dropdown values (right)
BOARD_MAPPING = {
    'Barisal': 'barisal',
    'Chittagong': 'chittagong',
    'Comilla': 'comilla',
    'Dhaka': 'dhaka',
    'Dinajpur': 'dinajpur',
    'Jessore': 'jessore',
    'Mymensingh': 'mymensingh',
    'Rajshahi': 'rajshahi',
    'Sylhet': 'sylhet',
    'Madrasah': 'madrasah',
    'Technical': 'tec',
    'DIBS(Dhaka)': 'dibs',
}

# --- Helper function to extract data from result page ---
def get_result_value(soup, label_text):
    """
    Finds a table data cell (<td>) containing the label_text (case-insensitive, partial match)
    and returns the text from its immediate sibling <td>.
    Assumes a structure like <td>Label:</td><td>Value</td>
    """
    label_td = soup.find('td', string=lambda text: text and label_text.lower() in text.lower())

    if label_td:
        value_td = label_td.find_next_sibling('td')
        if value_td:
            extracted_value = value_td.get_text(strip=True)
            return extracted_value
    return "N/A" # Return "N/A" if value not found

# --- Main Script ---
def extract_results_to_excel():
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
        sheet = workbook.active
    except FileNotFoundError:
        print(f"Error: Excel file not found at '{excel_file_path}'")
        return
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    headers = [cell.value for cell in sheet[1]]
    column_map = {header: i for i, header in enumerate(headers)}

    required_cols_input = ['Exam', 'Passing Year', 'Board', 'Roll', 'Reg']
    required_cols_output = ['Name', 'Father\'s Name', 'Mother\'s Name', 'GPA', 'Date of Birth']
    all_required_cols = required_cols_input + required_cols_output

    for col in all_required_cols:
        if col not in column_map:
            print(f"Error: Required column '{col}' not found in Excel sheet. Please check column headers.")
            return

    exam_col_idx = column_map['Exam'] + 1
    year_col_idx = column_map['Passing Year'] + 1
    board_col_idx = column_map['Board'] + 1
    roll_col_idx = column_map['Roll'] + 1
    reg_col_idx = column_map['Reg'] + 1

    name_col_idx = column_map['Name'] + 1
    father_name_col_idx = column_map['Father\'s Name'] + 1
    mother_name_col_idx = column_map['Mother\'s Name'] + 1
    gpa_col_idx = column_map['GPA'] + 1
    dob_col_idx = column_map['Date of Birth'] + 1

    session = requests.Session()

    # Define standard headers to mimic a browser
    common_headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Referer': base_url # The page from which the form was submitted
    }

    print("Starting data extraction...")

    for row_idx in range(2, sheet.max_row + 1):
        # Read data from Excel for current row
        exam_type_excel = sheet.cell(row=row_idx, column=exam_col_idx).value
        passing_year_excel = sheet.cell(row=row_idx, column=year_col_idx).value
        board_excel = sheet.cell(row=row_idx, column=board_col_idx).value
        roll_excel = sheet.cell(row=row_idx, column=roll_col_idx).value
        reg_no_excel = sheet.cell(row=row_idx, column=reg_col_idx).value

        # --- ADJUSTMENT HERE ---
        # Collect all essential input values for this row
        input_values_for_row = [exam_type_excel, passing_year_excel, board_excel, roll_excel, reg_no_excel]
        
        # Check if ALL essential input values are None or effectively empty strings
        # This prevents processing truly empty rows at the bottom of the sheet
        if all(val is None or (isinstance(val, str) and not val.strip()) for val in input_values_for_row):
            # We found a completely empty row regarding input data
            # Check if there's any content *at all* in the "Name" column for this row.
            # If so, it might be a previous error message, which we still want to skip or overwrite if retry is needed.
            current_name_value_for_empty_check = sheet.cell(row=row_idx, column=name_col_idx).value
            if current_name_value_for_empty_check is None or (isinstance(current_name_value_for_empty_check, str) and not current_name_value_for_empty_check.strip()):
                # Only print and skip if the row is actually empty and doesn't have an error message already
                print(f"Skipping row {row_idx}: All essential input fields are empty. No action taken.")
            else:
                # If it had an error message, we might still want to explicitly skip it based on the main skip logic below
                pass # Let the main skip logic handle this
            continue # Skip to the next row without further processing for this completely empty row

        print(f"\n--- Processing Row {row_idx} ---")

        # --- Main skip logic (for rows with some input, or existing error messages) ---
        current_name_value = sheet.cell(row=row_idx, column=name_col_idx).value
        if current_name_value and current_name_value not in [
            "Result Not Found", "Invalid Criteria", "HTTP Error", "Extraction Error",
            "Missing Input Data", "Captcha Not Found", "Captcha Parse Error",
            "Unknown Exam Type", "Unknown Board", "HTML Parse Error (Captcha)", "Unexpected Error during Captcha",
            "Extraction Failed"
        ]:
            print(f"Skipping row {row_idx}: Name already present ('{current_name_value}'). Assuming already processed.")
            continue

        # Check if essential input data is missing for a row that IS NOT completely empty (i.e., partially filled)
        if not all(input_values_for_row): # This checks if ANY of the values are None or False (like empty string)
            print(f"Skipping row {row_idx}: Missing one or more essential input data (Exam, Year, Board, Roll, Reg. No.).")
            sheet.cell(row=row_idx, column=name_col_idx, value="Missing Input Data")
            continue

        mapped_exam_type = EXAM_TYPE_MAPPING.get(str(exam_type_excel).strip(), None)
        mapped_board = BOARD_MAPPING.get(str(board_excel).strip(), None)

        if not mapped_exam_type:
            print(f"Skipping row {row_idx}: Unknown Exam Type '{exam_type_excel}'. Please add to EXAM_TYPE_MAPPING.")
            sheet.cell(row=row_idx, column=name_col_idx, value="Unknown Exam Type")
            continue
        if not mapped_board:
            print(f"Skipping row {row_idx}: Unknown Board '{board_excel}'. Please add to BOARD_MAPPING.")
            sheet.cell(row=row_idx, column=name_col_idx, value="Unknown Board")
            continue

        # --- Step 1: Get the initial page to extract the CAPTCHA ---
        try:
            print("Attempting to fetch homepage for captcha...")
            response_get = session.get(base_url, headers=common_headers)
            response_get.raise_for_status()
            soup_get = BeautifulSoup(response_get.text, 'html.parser')
            print(f"Homepage fetch successful. Status code: {response_get.status_code}")
            
            captcha_input_tag = soup_get.find('input', {'name': 'value_s'})
            if not captcha_input_tag:
                print(f"DEBUG: Could not find input tag with name 'value_s'.")
                raise ValueError("Captcha input tag not found.")

            captcha_td_with_input = captcha_input_tag.parent
            if not captcha_td_with_input or captcha_td_with_input.name != 'td':
                print(f"DEBUG: Parent of captcha input is not a <td> tag or not found.")
                raise ValueError("Captcha input's parent TD not found.")

            equals_td = captcha_td_with_input.find_previous_sibling('td')
            if not equals_td:
                print(f"DEBUG: Could not find the 'equals' TD (first previous sibling to input's parent TD).")
                raise ValueError("Equals TD not found.")

            captcha_text_td = equals_td.find_previous_sibling('td')
            if not captcha_text_td:
                print(f"DEBUG: Could not find the captcha text TD (second previous sibling to input's parent TD).")
                raise ValueError("Captcha text TD not found.")

            captcha_text = captcha_text_td.get_text(strip=True)
            if not captcha_text:
                print(f"DEBUG: Extracted captcha text is empty from TD: {captcha_text_td}")
                raise ValueError("Extracted captcha text is empty.")

            print(f"Extracted raw captcha text: '{captcha_text}'")

            match = re.search(r'(\d+)\s*([\+\-])\s*(\d+)', captcha_text)
            captcha_result = None
            if match:
                num1 = int(match.group(1))
                operator = match.group(2)
                num2 = int(match.group(3))
                if operator == '+':
                    captcha_result = num1 + num2
                elif operator == '-':
                    captcha_result = num1 - num2
            
            if captcha_result is None:
                print(f"Error: Could not parse captcha expression '{captcha_text}' for row {row_idx}.")
                sheet.cell(row=row_idx, column=name_col_idx, value="Captcha Parse Error")
                continue

            print(f"Calculated Captcha Result: {captcha_result}")

        except requests.exceptions.RequestException as e:
            print(f"HTTP Request error while getting captcha for row {row_idx}: {e}")
            sheet.cell(row=row_idx, column=name_col_idx, value=f"HTTP Error: {e}")
            continue
        except ValueError as e:
            print(f"Captcha extraction error for row {row_idx}: {e}")
            sheet.cell(row=row_idx, column=name_col_idx, value="HTML Parse Error (Captcha)")
            continue
        except Exception as e:
            print(f"An unexpected error occurred during captcha retrieval for row {row_idx}: {e}")
            sheet.cell(row=row_idx, column=name_col_idx, value=f"Unexpected Error during Captcha: {type(e).__name__}")
            continue

        # --- Step 2: Prepare and submit the POST request ---
        payload = {
            'exam': mapped_exam_type,
            'year': str(passing_year_excel),
            'board': mapped_board,
            'roll': str(roll_excel),
            'reg': str(reg_no_excel),
            'value_s': str(captcha_result),
            'sr': '3', # HIDDEN FIELD: Static value from homepage HTML
            'et': '2'  # HIDDEN FIELD: Static value from homepage HTML
        }

        print(f"Submitting form for Roll: {roll_excel}, Reg: {reg_no_excel}")

        try:
            post_response = session.post(post_url, data=payload, headers=common_headers)
            post_response.raise_for_status()

            result_soup = BeautifulSoup(post_response.text, 'html.parser')

            # --- ROBUST RESULT CHECK ---
            if not result_soup.find('td', string=lambda text: text and "Roll No" in text):
                print(f"Result Not Found for Roll: {roll_excel}, Reg: {reg_no_excel}.")
                print("The page returned does not contain the expected result table (no 'Roll No' label).")
                sheet.cell(row=row_idx, column=name_col_idx, value="Result Not Found")
                continue

            # If the "Roll No" label IS found, proceed with extraction
            print("Expected result table found. Attempting to extract data...")
            name = get_result_value(result_soup, "Name")
            father_name = get_result_value(result_soup, "Father's Name")
            mother_name = get_result_value(result_soup, "Mother's Name")
            gpa = get_result_value(result_soup, "GPA")
            dob = get_result_value(result_soup, "Date of Birth")

            # Validate if core data was actually extracted (not just "N/A")
            if name == "N/A" and father_name == "N/A" and mother_name == "N/A" and gpa == "N/A" and dob == "N/A":
                print(f"Warning: Even though 'Roll No' found, could not extract any other main result data for Roll: {roll_excel}, Reg: {reg_no_excel}.")
                print("This suggests a very subtle change in the HTML structure of the actual data fields.")
                # print("--- Full Result Page HTML for Debugging (Start) ---") # Uncomment for debugging
                # print(post_response.text)
                # print("--- Full Result Page HTML for Debugging (End) ---")
                sheet.cell(row=row_idx, column=name_col_idx, value="Extraction Failed")
                continue

            # Update Excel sheet with extracted data
            sheet.cell(row=row_idx, column=name_col_idx, value=name)
            sheet.cell(row=row_idx, column=father_name_col_idx, value=father_name)
            sheet.cell(row=row_idx, column=mother_name_col_idx, value=mother_name)
            sheet.cell(row=row_idx, column=gpa_col_idx, value=gpa)
            sheet.cell(row=row_idx, column=dob_col_idx, value=dob)

            print(f"Successfully extracted: Name='{name}', Father='{father_name}', GPA='{gpa}', DoB='{dob}'")

        except requests.exceptions.RequestException as e:
            print(f"HTTP Request error for Roll: {roll_excel}, Reg: {reg_no_excel}: {e}")
            sheet.cell(row=row_idx, column=name_col_idx, value=f"HTTP Error: {e}")
        except Exception as e:
            print(f"An unexpected error occurred while processing result for Roll: {roll_excel}, Reg: {reg_no_excel}: {e}")
            print("This usually means the HTML structure for result extraction changed or was not found.")
            sheet.cell(row=row_idx, column=name_col_idx, value=f"Extraction Error: {e}")

        # Be polite, wait a bit before the next request
        time.sleep(2)

    # Save the modified workbook
    try:
        workbook.save(excel_file_path)
        print(f"\nData extraction complete. Results saved to '{excel_file_path}'")
    except Exception as e:
        print(f"\nError saving Excel file: {e}. Please ensure the file is not open or write-protected.")

if __name__ == "__main__":
    extract_results_to_excel()
