import os
import time
import datetime
import Config
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service  # Import Service class
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import logging

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

# Set up the WebDriver Service
driver_path = Config.driver_path
service = Service(driver_path) 
driver = webdriver.Chrome(service=service)  # Initialize the driver with the service

# Read Excel data
excel_file = Config.excel_file
data = pd.read_excel(excel_file)  # Read the Excel file, assuming the second row is the header
data.columns = data.columns.str.strip()  # Remove any whitespace around column names
# print(data.columns)  # Debug: Verify column names

def fill_form_from_excel():
    global ver
    # Open the target URL
    form_url = "https://emis.sef.edu.pk/"
    driver.get(form_url)  # Navigate to the login page
    time.sleep(2)  # Allow page to load
    
    # Log in to the system
    try:
        driver.find_element(By.XPATH, "/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[1]/div/mat-form-field/div/div[1]/div[3]/input").send_keys(Config.Username)  
        driver.find_element(By.XPATH, "/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[2]/div/mat-form-field/div/div[1]/div[3]/input").send_keys(Config.Password)  # Enter password
        driver.find_element(By.XPATH, "/html/body/app-root/app-auth-layout/app-signin/div/div/div[2]/div/div/form/div[3]/div/button").click()  # Click login button
        driver.maximize_window()
        time.sleep(10)  # Wait for login to complete
    except Exception as e:
        log_error(logger, ERROR_CODES['LOGIN_FAILED'], f"Error in login: {e}")
    
    # Navigate to the enrollment section
    try:
        driver.find_element(By.XPATH, "//html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/a").click()  # Open dropdown
        driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[2]/a").click()  # Select specific option
        time.sleep(5)  # Wait for page to load
    except Exception as e:
        log_error(logger, ERROR_CODES['NAVIGATION_FAILED'], f"Error in navigation : {e}")

    # Iterate through rows in the Excel file
    for index, row in data.iterrows():
        ver = row["GR NO"]  # Update ver with the current GR NO
        if row["Admission Type"] == "New Admission":
            wait = WebDriverWait(driver, 20)  # Increase timeout to 20 seconds

            # Open the dropdown
            try:
                dropdown = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-0")))
                dropdown.click()
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "mat-select-panel")))

                # Select the desired option
                desired_option = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, f"//mat-option//span[normalize-space(text())='New Admission']")
                    )
                )
                desired_option.click()
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in dropdown: {e}', ver)

            # Wait for the options to appear
            try:
                da =  driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[1]/div/div/div[2]/div/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/input")
                actions = ActionChains(driver)
                actions.move_to_element(da).perform()
                da.clear()
                timestamp = pd.Timestamp(f'{row['Admission Date']}')
                formatted_date = timestamp.strftime('%m/%d/%Y')
                driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));", da, formatted_date)
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in  Date of admission: {e}', ver)
            
            try:
                gr = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[1]/div/div/div[2]/div/div[2]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                gr.clear()
                gr.send_keys(row["GR NO"])
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in  GR: {e}', ver)

            try:
                dropdownca = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-2")))
                actions = ActionChains(driver)
                actions.move_to_element(dropdownca).perform()
                dropdownca.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row['Class Admitted']}']")))
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Current admission class: {e}', ver)

            try:
                dropdowncc = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-4")))
                actions = ActionChains(driver)
                actions.move_to_element(dropdowncc).perform()
                dropdowncc.click()

                # Wait for the dropdown options to appear and select one by visible text
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row['Current Class']}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in  : {e}', ver)
            
            try:
                dropdownss = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-6")))
                actions = ActionChains(driver)
                actions.move_to_element(dropdownss).perform()
                dropdownss.click()

                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row['Select Section']}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Select Section: {e}', ver)

            try:
                dropdownm = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-8")))
                actions = ActionChains(driver)
                actions.move_to_element(dropdownm).perform()
                dropdownm.click()

                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row['Medium']}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Medium: {e}', ver)

            try:
                dropdowns = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-10")))
                actions = ActionChains(driver)
                actions.move_to_element(dropdowns).perform()
                dropdowns.click()

                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row['Shift']}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Shift: {e}', ver)

            try:
                sname = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[1]/mat-form-field/div/div[1]/div[3]/input")
                sname.clear()
                sname.send_keys(row["Students Name"])
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Students Name: {e}', ver)

            try:
                ssname = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                ssname.clear()
                ssname.send_keys(row["Student Surname"])
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Student Surname: {e}', ver)

            try:
                bform = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[3]/mat-form-field/div/div[1]/div[3]/input")
                bform.clear()
                if str(row["B-FORM"]) != "nan":
                    bform.send_keys(row["B-FORM"])
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in B-FORM: {e}', ver)
            
            try:
                dob =  driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[4]/mat-form-field/div/div[1]/div[3]/input")
                actions = ActionChains(driver)
                actions.move_to_element(dob).perform()
                dob.clear()
                timestamp = pd.Timestamp(f'{row['Date Of Birth']}')
                formatted_date = timestamp.strftime('%m/%d/%Y')
                driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));", dob, formatted_date)
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in DOB: {e}', ver)

            try:
                gender = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-12")))
                actions = ActionChains(driver)
                actions.move_to_element(gender).perform()
                gender.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row['Gender']}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Gender: {e}', ver)

            try:
                if row["Religion"] == 'Islam':
                    row["Religion"] = "Muslim"
                else:
                    row["Religion"] = "Non Muslim"
                religion = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-14")))
                actions = ActionChains(driver)
                actions.move_to_element(religion).perform()
                religion.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row['Religion']}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Religion: {e}', ver)

            try:
                Disability = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-16")))
                actions = ActionChains(driver)
                actions.move_to_element(Disability).perform()
                Disability.click()
                if row["Disability"] == "No":
                    row["Disability"] == "NO"
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='NO']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Disability: {e}', ver)

            try:
                BloodGroup = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-18")))
                actions = ActionChains(driver)
                actions.move_to_element(BloodGroup).perform()
                BloodGroup.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='N/A']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in BloodGroup: {e}', ver)

            try:
                MotherTongue = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-20")))
                actions = ActionChains(driver)
                actions.move_to_element(MotherTongue).perform()
                MotherTongue.click()

                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row["Mother Tongue"]}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                 log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in MotherTongue: {e}', ver)

            try:
                ecname = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[10]/mat-form-field/div/div[1]/div[3]/input")
                ecname.clear()
                if str(row["Emergency Contact Name"]) != "nan":
                    ecname.send_keys(row["Emergency Contact Name"])
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in  Emergency Contact Name: {e}', ver)

            try:
                ecno = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[1]/div[11]/mat-form-field/div/div[1]/div[3]/input")
                ecno.clear()
                ecno.send_keys(row["Emergency Contact Number"])
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Emergency Contact Number: {e}', ver)

            try:
                image_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Photos") 
                image_path = os.path.join(image_dir, f"{row['GR NO']}.jpg")
                if os.path.exists(image_path):
                    try:
                        upload_element = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[2]/div[1]/app-file-upload/div/input")
                        actions = ActionChains(driver)
                        actions.move_to_element(upload_element).perform()
                        upload_element.send_keys(image_path)  # Upload image
                    except Exception as e: 
                        log_error(logger, ERROR_CODES['INPUT_ERROR'], f"Error uploading image for GR {row['GR NO']}: {e}", ver)
                else:
                    log_error(logger, ERROR_CODES['INPUT_ERROR'], f"Image not found for GR {row['GR NO']}", ver)
            except Exception as e:
                 log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Image: {e}', ver)

            try:
                Region = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-22")))
                actions = ActionChains(driver)
                actions.move_to_element(Region).perform()
                Region.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row["Region"]}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Region : {e}', ver)

            try:
                District = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-24")))
                actions = ActionChains(driver)
                actions.move_to_element(District).perform()
                District.click()

                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row["District"]}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in District: {e}', ver)

            try:
                Taluka = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-26")))
                actions = ActionChains(driver)
                actions.move_to_element(Taluka).perform()
                Taluka.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row["Taluka"]}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Taluka: {e}', ver)

            try:
                UnionCoucil = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-28")))
                actions = ActionChains(driver)
                actions.move_to_element(UnionCoucil).perform()
                UnionCoucil.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row["Union Coucil"]}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in UnionCoucil: {e}', ver)

            try:
                Area = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[3]/div/div/div[2]/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/textarea")
                Area.clear()
                Area.send_keys(row["Cily/Village/Area"])
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in  Cily/Village/Area: {e}', ver)

            try:
                add = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[3]/div/div/div[2]/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/textarea")
                add.clear()
                add.send_keys(row["Address"])
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Address: {e}', ver)

            try:
                Salutaion = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-30")))
                actions = ActionChains(driver)
                actions.move_to_element(Salutaion).perform()
                Salutaion.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row["Salutaion"]}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Salutaion: {e}', ver)

            try:
                fname = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                fname.clear()
                fname.send_keys(row["Name"])
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Fname : {e}', ver)
            try:
                Surname = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[2]/div[1]/mat-form-field/div/div[1]/div[3]/input")
                Surname.clear()
                Surname.send_keys(row["Surname"])
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in FSurname : {e}', ver)

            try:
                Cnic = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[2]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                Cnic.clear()
                Cnic.send_keys(row["CNIC"])
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in CNIC: {e}', ver)

            try:
                Mobile = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[3]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                Mobile.clear()
                if str(row["Mobile No"]) != 'nan':
                    Mobile.send_keys(row["Mobile No"])
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Mobile: {e}', ver)

            if row['Qualification'] == "Primary":
                row['Qualification'] = "Matriculation"
            elif row['Qualification'] == "Matric":
                row['Qualification'] = "Matriculation"

            try:
                Qualification = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-32")))
                actions = ActionChains(driver)
                actions.move_to_element(Qualification).perform()
                Qualification.click()
                desired_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//mat-option/span[normalize-space(text())='{row["Qualification"]}']")))
                # desired_option.click()
                driver.execute_script("arguments[0].click();", desired_option)
            except Exception as e:
                log_error(logger, ERROR_CODES['DROPDOWN_ERROR'], f'Error in Qualification: {e}', ver)

            try:
                Occ = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[5]/div/div/div[2]/div/div[4]/div[2]/mat-form-field/div/div[1]/div[3]/input")
                Occ.clear()
                Occ.send_keys(row["Occupation"])
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Occupation: {e}', ver)

            try:
                Summit = driver.find_element(By.XPATH, "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/footer/div/div/button[1]")
                actions = ActionChains(driver)
                actions.move_to_element(Qualification).perform()
                # driver.execute_script("arguments[0].click();", Summit)
            except Exception as e:
                log_error(logger, ERROR_CODES['INPUT_ERROR'], f'Error in Summit: {e}', ver)

            driver.execute_script("window.scrollTo(0, 0);")
            # Wait before moving to the next record
            time.sleep(0.5)

    # Close the driver after completion
    driver.quit()

try:
    fill_form_from_excel()
except Exception as e:
    log_error(logger, ERROR_CODES['UNKNOWN_ERROR'], f"Error: {e} at GR No . {ver}")