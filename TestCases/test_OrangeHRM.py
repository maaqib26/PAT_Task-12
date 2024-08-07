from Locators.OrangeHRM_Locators import Locators
from Utilities.excel_functions import Excel_Operations
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pytest

# Load test data from Excel file
test_data_excel_file = Locators().excel_file
test_data_sheet_number = Locators().sheet_number

# Create an object for the excel class
excel_handler = Excel_Operations(test_data_excel_file,test_data_sheet_number)

class Test_OrangeHRM_Login():
    # Fixture to set up and tear down the browser
    @pytest.fixture
    def boot(self):
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 60)
        self.driver.get(Locators.url)
        yield
        self.driver.close()

    def test_login(self,boot):
        try:
            # Get the number of rows in the Excel sheet
            row = excel_handler.row_count()

            # Iterate over each row in the Excel sheet
            for row in range(2,row+1):
                test_username = excel_handler.read_data(row,7)
                test_password = excel_handler.read_data(row,8)

                # Enter the username and password
                user = self.wait.until(EC.presence_of_element_located((By.NAME,Locators().username)))
                user.send_keys(test_username)

                passw = self.wait.until(EC.presence_of_element_located((By.NAME, Locators().password)))
                passw.send_keys(test_password)

                # Click the login button
                login_button = self.wait.until(EC.element_to_be_clickable((By.XPATH,Locators().submit_button)))
                login_button.click()

                # Check if the login was successful
                if Locators().dashboard_url == self.driver.current_url:
                    print("SUCCESS : Login with Username {a} & Password {b}".format(a=test_username, b=test_password))
                    excel_handler.write_data(row,9,Locators.pass_data)
                    assert True, "Login successful"

                    # Logout
                    logout_dropdown_button = self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, Locators().logout_dropdown)))
                    logout_dropdown_button.click()
                    logout_button = self.wait.until(EC.element_to_be_clickable((By.LINK_TEXT,Locators().logout)))
                    logout_button.click()

                elif Locators().url == self.driver.current_url:
                    print("ERROR : Login unsuccessful with username {a} & Password {b}".format(a=test_username, b=test_password))
                    excel_handler.write_data(row,9,Locators().fail_data)
                    assert True, "Login failed"
                    self.driver.refresh()

        except (NoSuchElementException, ElementNotVisibleException) as e:
            print("ERROR : ", e)

