# AUTO_FAST

## Overview

This script automates the process of scraping phone numbers and email addresses from the FastPeopleSearch website based on the data provided in an Excel file.

## Requirements

- Python 3.11.8
- Selenium
- OpenPyXL
- TheFuzz
- Uc ChromeDriver

## Installation

1. Install Python 3.11.8 from [python.org](https://www.python.org/downloads/release/python-3118/).
2. Install the required Python packages:
   - Update pip to version 24 for better results:
     ```
     python -m pip install --upgrade pip
     pip install selenium openpyxl thefuzz undetectable chromedriver
     ```
3. Download Uc ChromeDriver from [here](https://pypi.org/project/undetected-chromedriver/) and import it into your code.

## Usage

1. Organize your data in an Excel file (default name: `100.xlsx`) with columns named: Name, Address, City, State, Phone Number, and Email Address.
2. Run the script using Python:
   ```
   python fast_try.py
   ```
3. The script will automatically navigate to the FastPeopleSearch website, search for each person's information, extract phone numbers and email addresses, and save them to the Excel file.

## Notes

- Please ensure to adjust the file path, column numbers, and XPATHs in the script according to your Excel file and the structure of the FastPeopleSearch website.
- This script assumes that each row in the Excel file corresponds to one person's information.
- Ensure you have a stable internet connection while running the script as it relies on web scraping.
- It's recommended to review and understand the code before running it to ensure it meets your requirements and doesn't violate any website's terms of service.
