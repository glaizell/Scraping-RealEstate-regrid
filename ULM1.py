import time
import pandas as pd
from fake_useragent import UserAgent
from selenium import webdriver
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import config

email = config.email
password = config.password

ua = UserAgent(os='windows', browsers=['edge', 'chrome'], min_percentage=1.3)
random_user_agent = ua.random

# Keep the browser open after the program finishes
options = Options()
options.add_argument("--disable-notifications")
options.add_experimental_option("detach", True)
options.add_argument(f"user-agent={random_user_agent}")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.get("https://regrid.com/")
driver.maximize_window()

mywait = WebDriverWait(driver, 20, poll_frequency=2)

driver.find_element(By.XPATH, "//a[@class='hd-cta-btn']").click()

menu_toggle_button = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH,
                                                                                       "//div[@class='flex-row-between width-100 nomobile']//div[@class='logo-menu-toggle']")))
menu_toggle_button.click()

# Click on the signup link
signup_link = WebDriverWait(driver, 20).until(
    EC.visibility_of_element_located((By.XPATH, "//a[@class='show-signup bold']")))
signup_link.click()

# Enter email and password
driver.find_element(By.XPATH,
                    "//div[@class='backdrop sign-in-card']//input[@id='map_signin_email']").send_keys(email)
driver.find_element(By.XPATH,
                    "//div[@class='backdrop sign-in-card']//input[@id='map_signin_password']").send_keys(password)
driver.find_element(By.XPATH, "//div[@class='backdrop sign-in-card']//input[@name='commit']").click()

file_path = "NC - February 2024 Land County List - Download Date 02.06.24 (For Scraper).xlsx"
ulm_1 = "ULM Mail #1 - Feb24 (Caldwell)"
df = pd.read_excel(file_path, sheet_name=ulm_1)
column_j_values = df.iloc[:, 0]


scraped_data = []

for apn in column_j_values:
    print(apn)
    # Create a dictionary to store all information for the current APN
    apn_data = {"APN": apn}

    search_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@id='glmap-search-query']")))

    try:
        # Execute JavaScript to clear the search input field
        driver.execute_script("arguments[0].value = '';", search_input)

        # Send keys to the cleared search input field
        search_input.send_keys(apn)

        try:
            address_element = mywait.until(EC.visibility_of_element_located(
                (By.XPATH, '//*[@id="glmap-search"]/span/div/div/div[1]/div[1]/div[1]/a')))
            address_element.click()

            # Wait for the parcel details to load
            mywait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="property"]')))

        except TimeoutException:
            print("Address element not found for APN:", apn, "\n")
            apn_data["Parcel Id"] = "Not found"
            apn_data["Deeded Owner"] = "Not found"
            apn_data["Total Parcel Value"] = "Not found"
            apn_data["Centroid Coordinates"] = "Not found"
            apn_data["Legal Description"] = "Not found"
            apn_data["Calculated Acres"] = "Not found"
            apn_data["FEMA Flood Zone Subtype"] = "Not found"
            scraped_data.append(apn_data)
            continue
        finally:
            search_input.clear()
            time.sleep(2)


        # Try to extract data if available
        try:
            # EXTRACT PARCEL ID
            try:
                parcel_id = driver.find_element(By.XPATH,
                                                "//td[text()='Parcel ID']/following-sibling::td[@class='value']").text
                apn_data["Parcel Id"] = parcel_id
                print("Parcel Id: ", parcel_id)
            except NoSuchElementException:
                apn_data["Parcel Id"] = "Not found"
                print("Parcel Id: Not found")
            time.sleep(1)  # Add a wait time after successful extraction

            # EXTRACT DEEDED OWNER
            try:
                deeded_owner = driver.find_element(By.XPATH,
                                               "//td[text()='Deeded Owner']/following-sibling::td[@class='value']").text
                apn_data["Deeded Owner"] = deeded_owner
                print("Deeded Owner: ", deeded_owner)
            except NoSuchElementException:
                apn_data["Deeded Owner"] = "Not found"
                print("Deeded Owner: Not found")
            time.sleep(1)  # Add a wait time after successful extraction

            # EXTRACT TOTAL PARCEL VALUE
            try:
                total_parcel_value = driver.find_element(By.XPATH,
                                                         "//td[text()='Total Parcel Value']/following-sibling::td[@class='value']").text
                apn_data["Total Parcel Value"] = total_parcel_value
                print("Total Parcel Value: ", total_parcel_value)
            except NoSuchElementException:
                apn_data["Total Parcel Value"] = "Not found"
                print("Total Parcel Value: Not found")
            time.sleep(1)  # Add a wait time after successful extraction

            # EXTRACT CENTROID COORDINATES
            try:
                centroid_coordinates = driver.find_element(By.XPATH,
                                                           "//td[text()='Centroid Coordinates']/following-sibling::td[@class='value']").text
                apn_data["Centroid Coordinates"] = centroid_coordinates
                print("Centroid Coordinates: ", centroid_coordinates)
            except NoSuchElementException:
                apn_data["Centroid Coordinates"] = "Not found"
                print("Centroid Coordinates: Not found")
            time.sleep(1)  # Add a wait time after successful extraction

            # EXTRACT CALCULATED ACRES
            try:
                calculated_acres = driver.find_element(By.XPATH,
                                                       "//td[text()='Calculated Acres']/following-sibling::td[@class='value']").text
                apn_data["Calculated Acres"] = calculated_acres
                print("Calculated Acres: ", calculated_acres)
            except NoSuchElementException:
                apn_data["Calculated Acres"] = "Not found"
                print("Calculated Acres: Not found")
            time.sleep(1)  # Add a wait time after successful extraction

            # EXTRACT LEGAL DESCRIPTION
            try:
                legal_description = driver.find_element(By.XPATH,
                                                        "//td[text()='Legal Description']/following-sibling::td[@class='value']").text
                apn_data["Legal Description"] = legal_description
                print("Legal Description: ", legal_description)
            except NoSuchElementException:
                apn_data["Legal Description"] = "Not found"
                print("Legal Description: Not found")
            time.sleep(1)  # Add a wait time after successful extraction

            # EXTRACT FEMA FLOOD ZONE
            try:
                fema_flood_zone_subtype = driver.find_element(By.XPATH,
                                                              "//tr[@class='field premium']/td[@class='key' and contains(text(), 'FEMA Flood Zone Subtype')]/following-sibling::td[@class='value']").text
                apn_data["FEMA Flood Zone Subtype"] = fema_flood_zone_subtype
                print("FEMA Flood Zone Subtype: ", fema_flood_zone_subtype)
            except NoSuchElementException:
                apn_data["FEMA Flood Zone Subtype"] = "Not found"
                print("FEMA Flood Zone Subtype: Not found")
            time.sleep(1)  # Add a wait time after successful extraction

            # EXTRACT CURRENT LINK
            try:
                link = driver.current_url
                apn_data["LINK"] = link
                print("Current Link: ", link)
            except NoSuchElementException:
                apn_data["LINK"] = "Not found"
                print("Current Link: Not found")
            print("")

        except NoSuchElementException:
            print("Some parcel details not found for APN:", apn, "\n")
            continue

        # Append the dictionary containing all information for the current APN to scraped_data
        scraped_data.append(apn_data)

        # Define column headers
        column_headers = ["APN", "Parcel Id", "Deeded Owner", "Total Parcel Value",
                          "Centroid Coordinates", "Calculated Acres", "Legal Description",
                          "FEMA Flood Zone Subtype", "LINK"]

        # Create a DataFrame from the scraped data with defined column headers
        scraped_df = pd.DataFrame(scraped_data, columns=column_headers)

        # # Print the DataFrame
        # print(scraped_df)

        print("Before saving CSV")
        # Save the DataFrame to a CSV file
        scraped_df.to_csv("notFound_check.csv", index=False)
        print("After saving CSV")

    finally:
        search_input.clear()
        time.sleep(5)



driver.quit()
