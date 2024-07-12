import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Function to load the Excel spreadsheet and get the list of consumer IDs and TV sets
def load_consumer_data():
    # Load the Excel spreadsheet
    wb = pd.ExcelFile('Consumer_id.xlsx')
    sheet = wb.parse(sheet_name='Sheet1')

    # Remove empty rows and spaces
    sheet = sheet.dropna()
    sheet = sheet.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # Get the list of consumer IDs and TV sets
    consumer_ids = sheet['Consumer ID'].tolist()
    tv_sets = sheet['TV Sets'].tolist()

    return consumer_ids, tv_sets

# Load consumer data
consumer_ids, tv_sets = load_consumer_data()

# Print to verify the loaded data
print("Consumer IDs:", consumer_ids)
print("TV Sets:", tv_sets)

# Specify the path to the chromedriver executable
chrome_options = Options()
chrome_options.add_argument("--start-maximized")  # Open browser in maximized mode
chrome_options.add_argument("--disable-infobars") # Disable infobars
chrome_options.add_argument("--disable-extensions") # Disable extensions
chrome_options.add_argument("--disable-gpu") # Disable GPU rendering for headless

# Create a new instance of the Chrome driver
driver = webdriver.Chrome(options=chrome_options)

# Open the website
driver.get("http://ptv.org.pk:72")

# Find the username and password input fields and the login button
username_field = driver.find_element(By.ID, "Inspector_Name")
password_field = driver.find_element(By.ID, "Password")
login_button = driver.find_element(By.ID, "Enter")

# Enter the username and password
username = "Soomro"
password = "Soomro509"
username_field.send_keys(username)
password_field.send_keys(password)

# Click the login button
login_button.click()

# Function to enter consumer IDs and TV sets
def enter_consumer_data(consumer_ids, tv_sets):
    for consumer_id, tv_set in zip(consumer_ids, tv_sets):
        print(f"Processing Consumer ID: {consumer_id}")  # Print the current consumer ID

        # Enter the consumer ID
        consumer_id_input = driver.find_element(By.ID, "WAPDAID")
        consumer_id_input.clear()
        consumer_id_input.send_keys(consumer_id)

        # Enter the TV sets
        tv_set_input = driver.find_element(By.ID, "TVSet")
        tv_set_input.clear()
        tv_set_input.send_keys(tv_set)

        # Optionally, wait for some time before moving to the next consumer
        # time.sleep(1)

# Call the function to enter consumer data
enter_consumer_data(consumer_ids, tv_sets)

# Optionally, wait for a while to see the logged-in state
import time
time.sleep(10)

# Close the browser window
driver.quit()
