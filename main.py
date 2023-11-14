# Inspecting the web page will let you know the classname or ID attributes of your website elements.
import time

import pandas as pd
from selenium import webdriver
from selenium.common import NoAlertPresentException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Read Excel file
excel_data = pd.read_excel('path_to_your_xl_sheet')

# Initialize WebDriver (assuming Chrome for this example) | For other browsers you may need to download their respective webdriver.
# Link to Edge WebDriver: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/#downloads
# For other web drivers you may require to search that.
driver = webdriver.Chrome()

# Open the website
driver.get('URL_where_your_xl_sheet_data_should_be_transfared.')

# Prompt user to perform manual login
input("Perform your work before the data transfer operation such as login or setting field value. When done press Enter...")

# Iterate through rows and columns, populating text fields
for index, row in excel_data.iterrows():
    # Assuming 'column1' and 'column2' are your Excel column names
    field1 = driver.find_element(By.ID, 'ID_of_textfield1')  # Replace ID_of_textfield1 with the ID attribute used
    field2 = driver.find_element(By.ID, 'ID_of_textfield2')  # Replace ID_of_textfield2 with the ID attribute used
    field2 = driver.find_element(By.ID, 'ID_of_textfield3')  # Replace ID_of_textfield3 with the ID attribute used
    field2 = driver.find_element(By.ID, 'ID_of_textfield4')  # Replace ID_of_textfield4 with the ID attribute used
#   .                                                   5
#   .                                                   6
#   .                                                   7

    field1.send_keys(Keys.CONTROL + "a")  # Select existing text if any
    field1.send_keys(Keys.DELETE)  # Clear existing text
    field1.send_keys(str(row['coulmn_name_that_would_be_transfared']))

    field2.send_keys(Keys.CONTROL + "a")  # Select existing text if any
    field2.send_keys(Keys.DELETE)  # Clear existing text
    field2.send_keys(str(row['coulmn_name_that_would_be_transfared']))
#   .
#   .
#   .
#   .
#   write the above same if any more text field you need.

    submit_button = driver.find_element(By.CLASS_NAME, "class_name_of_your_button")
    submit_button.click()

    # Hold the loop for 1.5 second before the next operation
    time.sleep(1.5)
    try:
        # Any alert present there will be accepted
        alert = driver.switch_to.alert
        alert.accept()
    except NoAlertPresentException:
        # No alert present, continue with the next iteration or perform any other required actions
        pass

    # Add a delay or handle synchronization if needed

# Close the browser when done
driver.quit()
