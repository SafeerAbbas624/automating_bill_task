import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager



# Developer's Stamp
print("This bot is coded and developed by Mr.SAFEER ABBAS. \n https://safeerabbas624.github.io/safeerabbas/"
      "\n https://github.com/SafeerAbbas624")


# Function to load the Excel spreadsheet and get the list of reference numbers
def load_reference_numbers():
    # Load the Excel spreadsheet
    wb = pd.ExcelFile('reference_id.xlsx')
    sheet = wb.parse(sheet_name='Sheet1')

    # Remove empty rows and spaces
    sheet = sheet.dropna()
    sheet = sheet.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # Remove all spaces
    sheet = sheet.replace(r'\s+', '', regex=True)

    # Get the list of reference numbers
    reference_numbers = sheet['Reference Numbers'].tolist()

    return reference_numbers


driver = webdriver.Chrome(ChromeDriverManager().install()) # Replace with your own Chrome driver path


# Function to open the website and enter the reference numbers
def enter_reference_numbers(reference_numbers):
    # Open the website
    count = 0
    driver.get('https://bill.pitc.com.pk/hescobill')
    time.sleep(5)

    # Loop through the list of reference numbers
    for reference_number in reference_numbers:
        # Enter the reference number in the text box
        input_element = driver.find_element(By.XPATH, '//*[@id="searchTextBox"]')
        input_element.clear()
        input_element.send_keys(reference_number)

        # Click the search button
        search_button = driver.find_element(By.XPATH, '//*[@id="btnSearch"]')
        search_button.click()

        # Wait for the bill to load
        time.sleep(3)

        try:
            # Copy the consumer ID and amount to be paid
            consumer_id = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/table/tbody/tr[2]/td[1]').text
            amount = driver.find_element(By.XPATH, '/html/body/div[2]/div[4]/table[2]/tbody/tr[3]/td[2]').text
        except Exception:
            print("Bill Not found")

        # Create a new Excel spreadsheet
        df = pd.DataFrame({'Reference ID': [reference_number], 'Consumer ID': [consumer_id], 'Amount': [amount]})
        # df.to_excel('bills.xlsx', index=False, header=True, mode='a')


        # Try Except condition not to get FileNotFoundError.
        try:
            # Opening Excel file if present, and appending dataframe to it.
             with pd.ExcelWriter('Hesco_data.xlsx', mode='a', if_sheet_exists="overlay") as writer:
                df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
        except FileNotFoundError:

            # Writing new file if not present in above code. Appending headers list also.
            with pd.ExcelWriter('Hesco_data.xlsx', mode='w') as writer:
                df.to_excel(writer, header=True)
        count  += 1
        print(f"appended the data into file count number = {count}")

        # Close the bill
        driver.back()
        time.sleep(2)


    # Close the Chrome window
    driver.quit()

# Call the functions
reference_numbers = load_reference_numbers()
print(reference_numbers)
enter_reference_numbers(reference_numbers)