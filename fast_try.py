from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from thefuzz import fuzz
import json
import time
import random
from name_check import name_validations as nValidate
from address_check import address_validations as aValidate
import openpyxl


class Iterations(nValidate, aValidate):
    def __init__(self, file_path, last_row_file, variables):
        self.var = variables
        self.file_path = file_path
        self.last_row_file = last_row_file
        self.last_row_number = self.load_last_row_number()
        self.driver = self.setup_driver()
        self.noRecordsFound = (
            "We could not find any records associated with that address."
        )

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.page_load_strategy = "eager"
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_options,
        )
        return driver

    def sleeper(self):
        time.sleep(
            float(
                "0."
                + random.choice(timers[0:4])
                + random.choice(timers[0:7])
                + random.choice(timers[0:9])
            )
        )

    def load_last_row_number(self):
        try:
            with open(self.last_row_file, "r") as file:
                last_row_number = int(file.read())
            return last_row_number
        except FileNotFoundError:
            return 0

    def get_next_row_data(self):
        wb = openpyxl.load_workbook(self.file_path)
        sheet = wb.active

        # Assuming the columns are named Name, Address, City, State, and Phone Number
        user_name = sheet.cell(row=self.last_row_number + 2, column=1).value
        formal_address = sheet.cell(row=self.last_row_number + 2, column=5).value
        city = sheet.cell(row=self.last_row_number + 2, column=6).value
        state = sheet.cell(row=self.last_row_number + 2, column=7).value
        phone_number = sheet.cell(row=self.last_row_number + 2, column=8).value
        email_address = sheet.cell(row=self.last_row_number + 2, column=9).value

        return user_name, formal_address, city, state, phone_number

    def save_last_row_number(self, last_row_number):
        with open(self.last_row_file, "w") as file:
            file.write(str(last_row_number))

    def read_data(self):  # READ DATA FROM THE PHONENUMBERS TEXT FILE
        phone_numbers = []
        with open("phone_numbers.txt", "r") as file:
            for line in file:
                phone_numbers.append(line)

        return phone_numbers  # RETURN PHONENUMBERS LIST

    def save_phone_number(self, phone_numbers):
        wb = openpyxl.load_workbook(self.file_path)
        sheet = wb.active
        try:
            # Assuming the phone number column is column H (column index 8)
            sheet.cell(row=self.last_row_number + 2, column=8).value = phone_numbers

        except ValueError:
            # Remove commas from phone numbers and concatenate them into one string
            phone_numbers_str = "".join(
                [phone.replace(",", "\n") for phone in phone_numbers]
            )
            sheet.cell(row=self.last_row_number + 2, column=8).value = phone_numbers_str

        wb.save(self.file_path)

        # Print the phone numbers in the terminal
        print("Phone Numbers Saved")

    def no_index(self, NoRecord):
        wb = openpyxl.load_workbook(self.file_path)
        sheet = wb.active
        # Assuming the phone number column is column H (column index 8)
        sheet.cell(row=self.last_row_number + 2, column=8).value = "no_Index"

        wb.save(self.file_path)
        print(f"NOTE: {NoRecord} ")

    def check_name(self, user_name):
        return nValidate.__checkValid__(self, user_name)

    def check_address(self, formal_address, city, state):
        return aValidate.__init__(self, formal_address, city, state)

    def search_bests(self):
        self.enteringDetails = True
        while self.enteringDetails:
            user_name, formal_address, city, state, _ = self.get_next_row_data()
            if self.check_name(user_name) == True:
                print("Company Found.")
                self.last_row_number += 1
                self.save_last_row_number(self.last_row_number)
                continue
            else:
                if self.check_address(formal_address, city, state) == True:
                    print("Address Found.")
                    self.save_phone_number(self.read_data())
                    print("Phone Numbers Received")
                    self.last_row_number += 1
                    self.save_last_row_number(self.last_row_number)
                    continue
                else:

                    try:
                        # Open the website
                        # time.sleep(5)
                        input_element = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.ID, self.var[0]))
                        )
                        input_element.click()
                        address_bar = self.driver.find_element(By.ID, self.var[1])
                        address_bar.send_keys(str(formal_address))
                        # self.sleeper()

                        city_bar = self.driver.find_element(By.ID, self.var[2])
                        city_bar.send_keys(str(city + "," + state))
                        # self.sleeper()

                        self.driver.find_element(By.XPATH, self.var[3]).click()

                        # Wait for search results to load
                        WebDriverWait(self.driver, 50).until(
                            EC.presence_of_all_elements_located(
                                (By.CLASS_NAME, "larger")
                            )
                        )
                        # Wait for the search results to load
                        if self.noRecordsFound in self.driver.page_source:
                            print("No INDEX")
                            self.no_index(self.noRecordsFound)
                            self.back()
                            self.last_row_number += 1
                            self.save_last_row_number(self.last_row_number)
                            continue

                        # Extracting names from the search results
                        titles = self.driver.find_elements(By.CLASS_NAME, "larger")
                        output_names = [names.text for names in titles]

                        # Check if output_names is available
                        print(f"Output Names: {output_names}")

                        if not output_names:
                            print("No output names were found.")
                            return

                        # Save the output names to a file
                        with open("output_names.txt", "w") as output:
                            for name in output_names:
                                output.write(f"{name}\n")

                        best_matches = []

                        # Iterate over the names to find best matches
                        for index in output_names:
                            # Print the current comparison
                            print(
                                f"Comparing '{user_name.lower()}' with '{index.lower()}'"
                            )
                            ratio = fuzz.ratio(user_name.lower(), index.lower())
                            print(f"Match Ratio: {ratio}")  # Print the calculated ratio
                            if ratio >= 40:  # adjust the ratio to 50
                                best_matches.append((index, ratio))

                        # Check if best_matches are available
                        print(f"Best Matches: {best_matches}")

                        # Sort matches by ratio in descending order from highest to lowest
                        best_matches.sort(key=lambda x: x[1], reverse=True)

                        # Save the best matches to a JSON file
                        with open("best_matches.json", "w") as outfile:
                            json.dump(best_matches, outfile)

                        # Print the best matches
                        print(f"Best matches found and saved to JSON: {best_matches}")

                        # Click on each best match and perform an action

                        for i, (name, ratio) in enumerate(best_matches):
                            print(f"Top {i+1} match: {name}, Score: {ratio}")
                            element = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable(
                                    (By.XPATH, f"//*[contains(text(), '{name}')]")
                                )
                            )
                            element.click()
                            self.driver.implicitly_wait(10)  # Wait for the page to load
                            try:
                                # Look for the element containing the phone numbers
                                phone_numbers_element = WebDriverWait(
                                    self.driver, 10
                                ).until(
                                    EC.presence_of_element_located(
                                        (By.CLASS_NAME, self.var[4])
                                    )
                                )
                                print(f"Phone numbers: {phone_numbers_element.text}")
                                self.save_phone_number(phone_numbers_element.text)
                                with open("phone_numbers.txt", "w") as f:
                                    f.write(phone_numbers_element.text)
                                break  # end the loop for the next data element
                            except NoSuchElementException as e:
                                print(f"Phone numbers not found for: {name}")
                                self.driver.back()  # Go back to search results
                                continue
                        if i == len(best_matches) - 1:
                            print("No phone numbers found for all best matches.")
                            break

                    except TimeoutException as e:
                        print(f"Timeout occurred: {e}")
                    except NoSuchElementException as e:
                        print(f"Element not found: {e}")
                    except Exception as e:
                        print(f"An error occurred: {e}")
                    finally:
                        if not self.get_next_row_data()[0]:
                            print("NO MORE ROWS TO CHECK EXCEL FILE IS CONCLUDED")
                            break
                        self.last_row_number += 1
                        self.save_last_row_number(self.last_row_number)
                        self.driver.get(
                            "https://www.fastpeoplesearch.com/"
                        )  # Go back to search results
                        continue

    def run(self):
        self.driver.get("https://www.fastpeoplesearch.com/")
        self.search_bests()
        self.driver.quit()


timers = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
variables = [
    "search-nav-link-address",
    "search-address-1",
    "search-address-2",
    '//*[@id="form-search-address"]/div[3]/button[2]',
    "detail-box-phone",
]
file_path = "100.xlsx"
last_row_file = "last_row_number.txt"
if __name__ == "__main__":
    iterations = Iterations(file_path, last_row_file, variables)
    iterations.run()
