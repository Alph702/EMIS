# README

## Automated Student Enrollment Script

### Description
This script automates the process of filling out student enrollment forms on the EMIS (Education Management Information System) platform using data from an Excel file. It leverages Python, Selenium WebDriver, and Pandas to automate web form interactions and populate fields with student information.

---

### Features
- Reads data from an Excel file (`DOC-20240330-WA0009.xlsx`).
- Automates login to the EMIS platform.
- Navigates to the student enrollment section.
- Iterates through each student record and fills in:
  - Personal information (e.g., name, gender, religion, DOB).
  - Admission details (e.g., admission type, class, section, shift).
  - Parent/guardian details.
  - Address and contact information.
- Handles image uploads for each student if the corresponding image exists in a `Photos` directory.
- Uses Selenium's advanced features such as `ActionChains` and explicit waits for reliability.

---

### Prerequisites
1. **Python**: Ensure Python is installed on your system.
2. **Dependencies**: Install the required Python libraries:
   ```bash
   pip install -r requirements.txt
   ```
3. **ChromeDriver**: Download the appropriate version of ChromeDriver for your Chrome browser and ensure it matches your browser version. Update the path in the script (`chromedriver.exe`).
4. **Excel File**: Place the `DOC-20240330-WA0009.xlsx` file in the same directory as the script. Ensure the file format and column headers match the expected format.
5. **Photos Directory**: Create a directory named `Photos` in the script's folder. Add student images named as `<GR NO>.jpg`.

---

### Setup
1. Update the script with your EMIS login credentials:
   - Replace `280103006` and `280103006-Mpk` with the appropriate username and password.
2. Verify the paths for:
   - ChromeDriver (`driver_path`).
   - Excel file (`excel_file`).

---

### Important Warning ⚠️
- **DO NOT minimize or unfocus the Chrome window while the bot is running!** The bot requires the Chrome window to be visible and in focus at all times.
- Minimizing or switching to another window may cause the bot to fail or input data incorrectly.
- Keep the Chrome window visible and active throughout the entire process.

---

### How to Run
1. Open a terminal or command prompt.
2. Navigate to the directory containing the script.
3. Run the script:
   ```bash
   python EMIS_Bot.py
   ```

---

### Expected Excel Format
The Excel file should have the following columns:
- `Admission Type`
- `Admission Date`
- `GR NO`
- `Class Admitted`
- `Current Class`
- `Select Section`
- `Medium`
- `Shift`
- `Students Name`
- `Student Surname`
- `B-FORM`
- `Date Of Birth`
- `Gender`
- `Religion`
- `Disability`
- `Mother Tongue`
- `Emergency Contact Name`
- `Emergency Contact Number`
- `Region`
- `District`
- `Taluka`
- `Union Coucil`
- `Cily/Village/Area`
- `Address`
- `Salutaion`
- `Name`
- `Surname`
- `CNIC`
- `Mobile No`
- `Qualification`
- `Occupation`

---

### Error Codes Reference
The script uses the following error codes for logging and troubleshooting:

| Error Code | Description | Possible Causes/Solutions |
|------------|-------------|-------------------------|
| E001 | LOGIN_FAILED | - Invalid credentials<br>- Network connectivity issues<br>- EMIS portal unavailable |
| E002 | NAVIGATION_FAILED | - Page elements changed<br>- Slow network connection<br>- Session timeout |
| E003 | EXCEL_READ_ERROR | - Invalid Excel format<br>- Missing required columns<br>- Corrupted Excel file |
| E004 | DROPDOWN_ERROR | - Invalid dropdown option<br>- Element not clickable<br>- Dynamic content not loaded |
| E005 | INPUT_ERROR | - Invalid data format<br>- Required field missing<br>- Field validation failure |
| E006 | DATE_FORMAT_ERROR | - Invalid date format in Excel<br>- Date parsing error<br>- Invalid date range |
| E007 | ELEMENT_NOT_FOUND | - Page structure changed<br>- Element not loaded<br>- Invalid XPath |
| E008 | NETWORK_ERROR | - Internet connection issues<br>- Server timeout<br>- API failure |
| E999 | UNKNOWN_ERROR | - Unexpected exceptions<br>- System-level errors<br>- Unhandled edge cases |

All errors are logged to `process_log.txt` with timestamps and detailed error messages.

---

### Notes
- **Error Handling**: The script logs missing images or errors during execution for debugging.
- **Delay Management**: Adjust `time.sleep()` delays as needed for your internet speed and server response times.
- **Security**: Avoid hardcoding sensitive information in the script.

---

### Disclaimer
Use this script responsibly and ensure it complies with your organization's policies and terms of service for using the EMIS platform. This script is intended for educational purposes and should be tested thoroughly before use in production.