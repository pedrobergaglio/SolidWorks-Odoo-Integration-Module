import time
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import random
import json
import requests
import mysql.connector
import os
from dotenv import load_dotenv
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
import datetime

# Get the current date
current_date = datetime.datetime.now()

# Extract the year and month
year = current_date.year
month = current_date.month -1


# Initialize the driver
#driver = webdriver.Chrome(ChromeDriverManager().install())
option = webdriver.ChromeOptions()
option.add_argument("start-maximized")
option.add_argument("--headless=new")
#--headless --disable-gpu --disable-software-rasterizer --disable-extensions --no-sandbox
option.add_argument("--disable-gpu")
option.add_argument("--disable-software-rasterizer")
option.add_argument("--disable-extensions")
option.add_argument("--no-sandbox")

appsheet_id = "acf512aa-6952-4aaf-8d17-c200fefa116b"
appsheet_key = "V2-RIUo6-uKEV7-puGvy-TeVYT-K2ag9-85j8j-6IaP2-ZX7Rr"

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=option)

carga = 5
error_flag = True

def scrape_clientes():
    # Call the login function to authenticate
    driver.get("https://docs.google.com/spreadsheets/d/1U-dKMZvsSuibee1C4uQ-ywBeY-RszsNrVHxAA4CIX88/edit#gid=0" )
    
    # Wait for the page to load after login 
    driver.implicitly_wait(carga)

    time.sleep(carga)

    # Now, you can locate the table and scrape its data using a loop
    # Locate the table by its ID
    table = driver.find_element(By.CLASS_NAME, "goog-inline-block grid4-inner-container")

    # Locate the table body (inside the table)
    table_body = table.find_element(By.TAG_NAME, "tbody")

    # Find all rows in the table body
    rows = table_body.find_elements(By.TAG_NAME, "tr")

    counter = 0
    payload_data = []

    # Iterate through rows
    for row in rows:

        counter += 1
        # Find all columns (td elements) in the row
        columns = row.find_elements(By.TAG_NAME, "td")

        # Check if the row has at least 6 columns (to ensure you can access the 2nd, 3rd, 5th, and 6th columns)
        if len(columns) >= 6:
            # Extract content from the desired columns (0-based index)
            cliente = columns[1].text
            cuit = columns[2].text
            deuda = columns[4].text
            ultimopago = columns[5].text

            # Create a dictionary for the row
            row_data = {
                "number": counter,
                "customer": cliente,
                "cuit": cuit,
                "debt": parse_amount_abs(deuda),
                "last_payment": ultimopago,
                "cuenta": cuenta
            }

            payload_data.append(row_data)
    
    # Build the payload with all row data
    payload = {
        "Action": "Add",
        "Properties": {
            "Locale": "es-AR",
            "Timezone": "Argentina Standard Time"
        },
        "Rows": payload_data
    }

    #print(payload)

    # Define the request URL
    requestURL = f"https://api.appsheet.com/api/v2/apps/{appsheet_id}/tables/customer_accounts/Action"

    # Set request headers
    headers = {
        "Content-Type": "application/json",
        "ApplicationAccessKey": appsheet_key
    }

    # Send the request
    response = requests.post(requestURL, data=json.dumps(payload), headers=headers)

    # Check the response status code
    if response.status_code != 200:
        print(f"Request failed with status code: {response.status_code}")
        print(response.text)
        errorflag = False
