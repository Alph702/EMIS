import os
import re
import time
import Config
import pandas as pd
from playwright.sync_api import sync_playwright
import logging

# --- Logging Setup ---
logging.basicConfig(
    filename='process_log.txt',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

ERROR_CODES = {
    'LOGIN_FAILED': 'E001',
    'NAVIGATION_FAILED': 'E002',
    'EXCEL_READ_ERROR': 'E003',
    'DROPDOWN_ERROR': 'E004',
    'INPUT_ERROR': 'E005',
    'DATE_FORMAT_ERROR': 'E006',
    'ELEMENT_NOT_FOUND': 'E007',
    'NETWORK_ERROR': 'E008',
    'UNKNOWN_ERROR': 'E999'
}

def log_error(logger, error_code, message, gr_no=None):
    """Log errors with consistent format."""
    error_msg = f"[{error_code}] {message}"
    if gr_no:
        error_msg += f" (GR NO: {gr_no})"
    logger.error(error_msg)
    return error_msg

# --- Helper Functions ---
def select_dropdown(page, element_id, value, error_code, field_name, gr_no):
    """Select an option from a dropdown by visible text."""
    try:
        page.click(f"id={element_id}")
        page.click(f"xpath=//mat-option/span[normalize-space(text())='{value}']")
    except Exception as e:
        log_error(logger, error_code, f'Error in {field_name}: {e}', gr_no)

def fill_input(page, xpath, value, error_code, field_name, gr_no, is_int=False):
    """Fill an input field if value is not NaN."""
    try:
        if pd.notna(value):
            fill_val = str(int(value)) if is_int else str(value)
            page.locator(xpath).fill(fill_val)
    except Exception as e:
        log_error(logger, error_code, f'Error in {field_name}: {e}', gr_no)

def fill_date(page, xpath, value, error_code, field_name, gr_no):
    """Fill a date input field with formatted date."""
    try:
        if pd.notna(value):
            formatted_date = pd.Timestamp(str(value)).strftime('%m/%d/%Y')
            page.locator(xpath).fill(formatted_date)
    except Exception as e:
        log_error(logger, error_code, f'Error in {field_name}: {e}', gr_no)

def upload_image(page, gr_no, error_code):
    """Upload image if exists for the given GR NO."""
    try:
        image_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Photos")
        image_path = os.path.join(image_dir, f"{gr_no}.jpg")
        if os.path.exists(image_path):
            page.set_input_files(
                "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[2]/div[1]/app-file-upload/div/input",
                image_path
            )
        else:
            log_error(logger, error_code, f"Image not found for GR {gr_no}", gr_no)
    except Exception as e:
        log_error(logger, error_code, f'Error in Image: {e}', gr_no)

def navigate_to(page, option: str, error_code, gr_no):
    """Navigate to a specific section in the EMIS portal."""
    try:
        page.click("xpath=//html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/a")
        if option == "Add Student":
            page.click("xpath=/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[2]/a")
            page.wait_for_timeout(3000)
        elif option == "Student List":
            page.click("xpath=/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[1]/a")
            page.wait_for_timeout(3000)
    except Exception as e:
        log_error(logger, error_code, f"Error in navigation: {e}", gr_no)

def select_student_by_gr(page, gr_no, error_code):
    """Select a student from the list by GR NO."""
    try:
        # navigate to student list
        navigate_to(page, "Student List", ERROR_CODES['NAVIGATION_FAILED'], gr_no)
        fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-all-students/section/div/div[2]/div/div/div/div/div/div/div[1]/div/div[1]/ul/li[2]/input", gr_no, ERROR_CODES['INPUT_ERROR'], "Search", gr_no)
        page.wait_for_timeout(100)
        page.click("xpath=/html/body/app-root/app-main-layout/div/app-all-students/section/div/div[2]/div/div/div/div/div/div/div[1]/div/div[1]/ul/li[3]/div/button")
        page.wait_for_timeout(2000)
        # Build exact regex pattern
        pattern = re.compile(rf"^\s*{str(gr_no)}\s*$")

        # Find the GR cell with exact match
        gr_cell = page.locator("mat-cell.cdk-column-grNo").filter(has_text=pattern)

        student = gr_cell.locator("xpath=ancestor::mat-row")
        student.locator("button[mat-icon-button]").nth(0).click()
    except Exception as e:
        log_error(logger, error_code, f"Error selecting student with GR NO {gr_no}: {e}", gr_no)

# --- Main Form Filling Logic ---
def fill_form_from_excel():
    ver = None
    excel_file = Config.excel_file
    data = pd.read_excel(excel_file)
    data.columns = data.columns.str.strip()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # --- Login ---
        try:
            page.goto("https://emis.sef.edu.pk/")
            page.wait_for_timeout(2000)
            page.fill("xpath=/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[1]/div/mat-form-field/div/div[1]/div[3]/input", Config.Username)
            page.fill("xpath=/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[2]/div/mat-form-field/div/div[1]/div[3]/input", Config.Password)
            page.click("xpath=/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[3]/div/button")
            page.wait_for_timeout(10000)
        except Exception as e:
            log_error(logger, ERROR_CODES['LOGIN_FAILED'], f"Error in login: {e}")
            browser.close()
            return

        # --- Iterate Through Excel Rows ---
        for index, row in data.iterrows():
            ver = row["GR NO"]
            if row["Admission Type"] == "New Admission":
                # --- Navigate to Enrollment Section ---
                navigate_to(page, "Add Student", ERROR_CODES['NAVIGATION_FAILED'], ver)
                # --- Admission Details ---
                select_dropdown(page, "mat-select-0", "New Admission", ERROR_CODES['DROPDOWN_ERROR'], "Admission Type", ver)
                fill_date(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[1]/div/div/div[2]/div/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/input", row["Admission Date"], ERROR_CODES['INPUT_ERROR'], "Admission Date", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[1]/div/div/div[2]/div/div[2]/div[2]/mat-form-field/div/div[1]/div[3]/input", row["GR NO"], ERROR_CODES['INPUT_ERROR'], "GR NO", ver)
                select_dropdown(page, "mat-select-2", row['Class Admitted'], ERROR_CODES['DROPDOWN_ERROR'], "Class Admitted", ver)
                select_dropdown(page, "mat-select-4", row['Current Class'], ERROR_CODES['DROPDOWN_ERROR'], "Current Class", ver)
                select_dropdown(page, "mat-select-6", row['Select Section'], ERROR_CODES['DROPDOWN_ERROR'], "Select Section", ver)
                select_dropdown(page, "mat-select-8", row['Medium'], ERROR_CODES['DROPDOWN_ERROR'], "Medium", ver)
                select_dropdown(page, "mat-select-10", row['Shift'], ERROR_CODES['DROPDOWN_ERROR'], "Shift", ver)

                # --- Student Details ---
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[1]/mat-form-field/div/div[1]/div[3]/input", row["Students Name"], ERROR_CODES['INPUT_ERROR'], "Students Name", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input", row["Student Surname"], ERROR_CODES['INPUT_ERROR'], "Student Surname", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[3]/mat-form-field/div/div[1]/div[3]/input", row["B-FORM"], ERROR_CODES['INPUT_ERROR'], "B-FORM", ver, is_int=True)
                fill_date(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[4]/mat-form-field/div/div[1]/div[3]/input", row["Date Of Birth"], ERROR_CODES['INPUT_ERROR'], "Date Of Birth", ver)
                select_dropdown(page, "mat-select-12", row['Gender'], ERROR_CODES['DROPDOWN_ERROR'], "Gender", ver)
                religion_value = "Muslim" if row["Religion"] == 'Islam' else "Non Muslim"
                select_dropdown(page, "mat-select-14", religion_value, ERROR_CODES['DROPDOWN_ERROR'], "Religion", ver)
                disability_value = "NO" if str(row["Disability"]).lower() == "no" else str(row["Disability"])
                select_dropdown(page, "mat-select-16", disability_value, ERROR_CODES['DROPDOWN_ERROR'], "Disability", ver)
                blood_group_value = "N/A" if pd.isna(row['Blood Group']) else str(row['Blood Group'])
                select_dropdown(page, "mat-select-18", blood_group_value, ERROR_CODES['DROPDOWN_ERROR'], "Blood Group", ver)
                select_dropdown(page, "mat-select-20", row['Mother Tongue'], ERROR_CODES['DROPDOWN_ERROR'], "Mother Tongue", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[10]/mat-form-field/div/div[1]/div[3]/input", row["Emergency Contact Name"], ERROR_CODES['INPUT_ERROR'], "Emergency Contact Name", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[11]/mat-form-field/div/div[1]/div[3]/input", row["Emergency Contact Number"], ERROR_CODES['INPUT_ERROR'], "Emergency Contact Number", ver, is_int=True)
                upload_image(page, ver, ERROR_CODES['INPUT_ERROR'])

                # --- Location Details ---
                select_dropdown(page, "mat-select-22", row['Region'], ERROR_CODES['DROPDOWN_ERROR'], "Region", ver)
                select_dropdown(page, "mat-select-24", row['District'], ERROR_CODES['DROPDOWN_ERROR'], "District", ver)
                select_dropdown(page, "mat-select-26", row['Taluka'], ERROR_CODES['DROPDOWN_ERROR'], "Taluka", ver)
                select_dropdown(page, "mat-select-28", row['Union Coucil'], ERROR_CODES['DROPDOWN_ERROR'], "Union Coucil", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[3]/div/div/div[2]/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/textarea", row["Cily/Village/Area"], ERROR_CODES['INPUT_ERROR'], "Cily/Village/Area", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[3]/div/div/div[2]/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/textarea", row["Address"], ERROR_CODES['INPUT_ERROR'], "Address", ver)

                # --- Father's Details ---
                select_dropdown(page, "mat-select-30", row['Salutaion'], ERROR_CODES['DROPDOWN_ERROR'], "Salutaion", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input", row["Name"], ERROR_CODES['INPUT_ERROR'], "Father's Name", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/input", row["Surname"], ERROR_CODES['INPUT_ERROR'], "Father's Surname", ver)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[2]/div[2]/mat-form-field/div/div[1]/div[3]/input", row["CNIC"], ERROR_CODES['INPUT_ERROR'], "CNIC", ver, is_int=True)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[3]/div[2]/mat-form-field/div/div[1]/div[3]/input", row["Mobile No"], ERROR_CODES['INPUT_ERROR'], "Mobile No", ver, is_int=True)
                fill_input(page, "xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[4]/div[2]/mat-form-field/div/div[1]/div[3]/input", row["Occupation"], ERROR_CODES['INPUT_ERROR'], "Occupation", ver)

                # --- Qualification ---
                qualification_value = "Matriculation" if row['Qualification'] in ["Primary", "Matric"] else row['Qualification']
                select_dropdown(page, "mat-select-32", qualification_value, ERROR_CODES['DROPDOWN_ERROR'], "Qualification", ver)

                # --- Submit Form ---
                # try:
                #     page.click("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/footer/div/div/button[1]")
                #     time.sleep(1)
                # except Exception as e:
                #     log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Submit: {e}', ver)

                # --- Prepare for Next Record ---
                time.sleep(2)
                page.click("xpath=/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[3]/a")
                page.wait_for_timeout(3000)
                page.reload()
                page.click("xpath=//html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/a")
                page.click("xpath=/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[2]/a")
                page.wait_for_timeout(2000)
            
            elif row["Admission Type"] == "Promoted":
                # --- Promotion Logic Here ---
                # Select student by GR NO
                select_student_by_gr(page, ver, ERROR_CODES['INPUT_ERROR'])

                time.sleep(5)
                pass
            elif row["Admission Type"] == "Retained":
                pass
            elif row["Admission Type"] == "Passout":
                pass
            elif row["Admission Type"] == "Dropout":
                pass
            elif row["Admission Type"] == "TC":
                pass

        browser.close()

# --- Entry Point ---
try:
    fill_form_from_excel()
except Exception as e:
    log_error(logger, ERROR_CODES['UNKNOWN_ERROR'], f"Error: {e}", None)