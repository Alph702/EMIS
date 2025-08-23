import os
import time
import datetime
import Config # Assuming Config.py exists and contains necessary configurations
import pandas as pd
from playwright.sync_api import sync_playwright
import logging
import math # Import math to check for NaN

# Set up logging configuration
logging.basicConfig(
    filename='process_log.txt',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Error codes for different types of failures
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
    """Utility function to log errors with consistent format"""
    error_msg = f"[{error_code}] {message}"
    if gr_no:
        error_msg += f" (GR NO: {gr_no})"
    logger.error(error_msg)
    return error_msg

ver = None  # To track the current GR NO

# Read Excel data
excel_file = Config.excel_file
data = pd.read_excel(excel_file)  # Read the Excel file, assuming the second row is the header
data.columns = data.columns.str.strip()  # Remove any whitespace around column names

def fill_form_from_excel():
    global ver
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # You can set headless=True for a headless browser
        page = browser.new_page()

        # Open the target URL
        form_url = "https://emis.sef.edu.pk/"
        try:
            page.goto(form_url)  # Navigate to the login page
            page.wait_for_timeout(2000)  # Allow page to load
        except Exception as e:
            log_error(logger, ERROR_CODES['NAVIGATION_FAILED'], f"Error navigating to login page: {e}")
            browser.close()
            return

        # Log in to the system
        try:
            page.fill("xpath=/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[1]/div/mat-form-field/div/div[1]/div[3]/input", Config.Username)
            page.fill("xpath=/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[2]/div/mat-form-field/div/div[1]/div[3]/input", Config.Password)
            page.click("xpath=/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[3]/div/button") # Corrected XPATH for login button
            page.wait_for_timeout(10000)  # Wait for login to complete
        except Exception as e:
            log_error(logger, ERROR_CODES['LOGIN_FAILED'], f"Error in login: {e}")
            browser.close()
            return

        # Navigate to the enrollment section
        try:
            page.click("xpath=//html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/a")  # Open dropdown
            page.click("xpath=/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[2]/a")  # Select specific option
            page.wait_for_timeout(5000)  # Wait for page to load
        except Exception as e:
            log_error(logger, ERROR_CODES['NAVIGATION_FAILED'], f"Error in navigation: {e}")
            browser.close()
            return

        # Iterate through rows in the Excel file
        for index, row in data.iterrows():
            ver = row["GR NO"]  # Update ver with the current GR NO
            if row["Admission Type"] == "New Admission":
                # Open the dropdown
                try:
                    page.click("id=mat-select-0")
                    page.wait_for_selector(".mat-select-panel")

                    # Select the desired option
                    page.click(f"xpath=//mat-option//span[normalize-space(text())='New Admission']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in dropdown: {e}', ver)

                # Fill Admission Date
                try:
                    da_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[1]/div/div/div[2]/div/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/input")
                    timestamp = pd.Timestamp(f'{row["Admission Date"]}')
                    formatted_date = timestamp.strftime('%m/%d/%Y')
                    da_element.fill(formatted_date)
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Date of admission: {e}', ver)

                # Fill GR NO
                try:
                    gr_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[1]/div/div/div[2]/div/div[2]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                    gr_element.fill(str(row["GR NO"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in GR: {e}', ver)

                # Fill Class Admitted
                try:
                    page.click("id=mat-select-2")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Class Admitted']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Current admission class: {e}', ver)

                # Fill Current Class
                try:
                    page.click("id=mat-select-4")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Current Class']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Current Class: {e}', ver)
                
                # Fill Select Section
                try:
                    page.click("id=mat-select-6")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Select Section']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Select Section: {e}', ver)

                # Fill Medium
                try:
                    page.click("id=mat-select-8")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Medium']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Medium: {e}', ver)

                # Fill Shift
                try:
                    page.click("id=mat-select-10")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Shift']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Shift: {e}', ver)

                # Fill Students Name
                try:
                    sname_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[1]/mat-form-field/div/div[1]/div[3]/input")
                    sname_element.fill(str(row["Students Name"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Students Name: {e}', ver)

                # Fill Student Surname
                try:
                    ssname_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                    ssname_element.fill(str(row["Student Surname"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Student Surname: {e}', ver)

                # Fill B-FORM
                try:
                    bform_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[3]/mat-form-field/div/div[1]/div[3]/input")
                    # Only fill if not NaN
                    if pd.notna(row["B-FORM"]):
                        bform_element.fill(str(int(row["B-FORM"]))) # Convert to int first to remove .0, then to str
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in B-FORM: {e}', ver)
                
                # Fill Date Of Birth
                try:
                    dob_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[4]/mat-form-field/div/div[1]/div[3]/input")
                    timestamp = pd.Timestamp(f'{row["Date Of Birth"]}')
                    formatted_date = timestamp.strftime('%m/%d/%Y')
                    dob_element.fill(formatted_date)
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in DOB: {e}', ver)

                # Fill Gender
                try:
                    page.click("id=mat-select-12")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Gender']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Gender: {e}', ver)

                # Fill Religion
                try:
                    religion_value = row["Religion"]
                    if religion_value == 'Islam':
                        religion_value = "Muslim"
                    else:
                        religion_value = "Non Muslim"
                    page.click("id=mat-select-14")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{religion_value}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Religion: {e}', ver)

                # Fill Disability
                try:
                    page.click("id=mat-select-16")
                    disability_value = "NO" if str(row["Disability"]).lower() == "no" else str(row["Disability"])
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{disability_value}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Disability: {e}', ver)

                # Fill BloodGroup
                try:
                    page.click("id=mat-select-18")
                    blood_group_value = str(row['Blood Group'])
                    if pd.isna(row['Blood Group']): # Check if NaN explicitly for dropdown
                        blood_group_value = "N/A" # Default if Blood Group is missing in Excel
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{blood_group_value}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in BloodGroup: {e}', ver)

                # Fill MotherTongue
                try:
                    page.click("id=mat-select-20")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Mother Tongue']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in MotherTongue: {e}', ver)

                # Fill Emergency Contact Name
                try:
                    ecname_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[10]/mat-form-field/div/div[1]/div[3]/input")
                    if pd.notna(row["Emergency Contact Name"]):
                        ecname_element.fill(str(row["Emergency Contact Name"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Emergency Contact Name: {e}', ver)

                # Fill Emergency Contact Number
                try:
                    ecno_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[11]/mat-form-field/div/div[1]/div[3]/input")
                    # Check if the value is not NaN before attempting to fill
                    if pd.notna(row["Emergency Contact Number"]):
                        # Convert to string to handle potential float conversion from pandas for numbers
                        ecno_element.fill(str(int(row["Emergency Contact Number"])))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Emergency Contact Number: {e}', ver)

                # Upload Image
                try:
                    image_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Photos")
                    image_path = os.path.join(image_dir, f"{row['GR NO']}.jpg")
                    if os.path.exists(image_path):
                        # Playwright's set_input_files works directly on input type="file" elements
                        page.set_input_files("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[2]/div[1]/app-file-upload/div/input", image_path)
                    else:
                        log_error(logger, ERROR_CODES['INPUT_ERROR'], f"Image not found for GR {row['GR NO']}", ver)
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Image: {e}', ver)
                

                # Fill Region
                try:
                    page.click("id=mat-select-22")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Region']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Region: {e}', ver)

                # Fill District
                try:
                    page.click("id=mat-select-24")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['District']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in District: {e}', ver)

                # Fill Taluka
                try:
                    page.click("id=mat-select-26")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Taluka']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Taluka: {e}', ver)

                # Fill Union Council
                try:
                    page.click("id=mat-select-28")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Union Coucil']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in UnionCoucil: {e}', ver)

                # Fill City/Village/Area
                try:
                    area_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[3]/div/div/div[2]/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/textarea")
                    # Only fill if not NaN
                    if pd.notna(row["Cily/Village/Area"]):
                        area_element.fill(str(row["Cily/Village/Area"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Cily/Village/Area: {e}', ver)

                # Fill Address
                try:
                    add_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[3]/div/div/div[2]/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/textarea")
                    # Only fill if not NaN
                    if pd.notna(row["Address"]):
                        add_element.fill(str(row["Address"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Address: {e}', ver)

                # Fill Salutation
                try:
                    page.click("id=mat-select-30")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{row['Salutaion']}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Salutaion: {e}', ver)

                # Fill Father's Name
                try:
                    fname_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                    fname_element.fill(str(row["Name"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Fname: {e}', ver)

                # Fill Father's Surname
                try:
                    surname_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/input")
                    surname_element.fill(str(row["Surname"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in FSurname: {e}', ver)

                # Fill CNIC
                try:
                    cnic_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[2]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                    # Only fill if not NaN
                    if pd.notna(row["CNIC"]):
                        cnic_element.fill(str(int(row["CNIC"]))) # Convert to int first to remove .0, then to str
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in CNIC: {e}', ver)

                # Fill Mobile No
                try:
                    mobile_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[3]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                    if pd.notna(row["Mobile No"]):
                        mobile_element.fill(str(int(row["Mobile No"])))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Mobile: {e}', ver)

                # Fill Qualification
                try:
                    qualification_value = row['Qualification']
                    if qualification_value == "Primary" or qualification_value == "Matric":
                        qualification_value = "Matriculation"
                    page.click("id=mat-select-32")
                    page.click(f"xpath=//mat-option/span[normalize-space(text())='{qualification_value}']")
                except Exception as e:
                    log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Qualification: {e}', ver)

                # Fill Occupation
                try:
                    occ_element = page.locator("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[4]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                    # Only fill if not NaN
                    if pd.notna(row["Occupation"]):
                        occ_element.fill(str(row["Occupation"]))
                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Occupation: {e}', ver)

                # Submit button (commented out as in original)
                try:
                    # pass
                    page.click("xpath=/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/footer/div/div/button[1]")
                    time.sleep(1)  # Wait for the form to submit

                except Exception as e:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Summit: {e}', ver)

                # page.evaluate("window.scrollTo(0, 0);")
                # page.goto(form_url)  # Navigate to the login page
                # page.wait_for_timeout(2000)  # Allow page to load
                time.sleep(2)  # Wait for the form to submit
                page.click("xpath=/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[3]/a")  # Open dropdown
                page.wait_for_timeout(5000)
                page.reload()
                page.click("xpath=//html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/a")  # Open dropdown
                page.click("xpath=/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[2]/a")  # Select specific option
                page.wait_for_timeout(2000)  # Wait for page to load
                # Wait before moving to the next record

        browser.close()

try:
    fill_form_from_excel()
except Exception as e:
    log_error(logger, ERROR_CODES['UNKNOWN_ERROR'], f"Error: {e} at GR No . {ver}")