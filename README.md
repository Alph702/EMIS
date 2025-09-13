# EMIS Form Filler

A Streamlit web application that automates filling out the EMIS web form using data from an Excel file.

## Installation and Setup

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/Alph702/EMIS.git
    ```

2.  **Install dependencies:**
    Make sure you have `uv` installed. Then, run the following command to install the project dependencies:
    ```bash
    uv sync
    ```

3.  **Install the browser for Playwright:**
    Playwright requires a browser to automate. This command will download and install the Chromium browser for Playwright to use.
    ```bash
    uv run playwright install chromium
    ```

## How to Run the Application

To run the Streamlit application, use the following command:

```bash
uv run streamlit run app.py
```

The application will be available at `http://localhost:8501`.

## Template File

The application requires an Excel file with a specific format. You can download a template file named `template.xlsx` from the application's user interface. Make sure your Excel file has the same columns as the template file.
