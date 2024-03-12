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

mywait = WebDriverWait(driver, 30, poll_frequency=2)

driver.find_element(By.XPATH, "//a[@class='hd-cta-btn']").click()

menu_toggle_button = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH,
                                                                                       "//div[@class='flex-row-between width-100 nomobile']//div[@class='logo-menu-toggle']")))
menu_toggle_button.click()
time.sleep(2)
# Click on the signup link
signup_link = WebDriverWait(driver, 30).until(
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
ulm_3 = "ULM Mailer #3 - Apr24 (Chatham)"
df = pd.read_excel(file_path, sheet_name=ulm_3, dtype={'APN': str})
# Selecting rows starting from index 743
column_j_values = df.iloc[629:, 9]
column_a_values = df.iloc[629:, 0]

# Initialize scraped data outside the loop
scraped_data = []
column_headers = ["APN", "Parcel Id", "Deeded Owner", "Total Parcel Value",
                  "Centroid Coordinates", "Calculated Acres", "Legal Description",
                  "FEMA Flood Zone Subtype", "LINK"]

for apn, property_text in zip(column_j_values, column_a_values):
    print(apn)

    apn_data = {"APN": apn}


    # Wait for the search input field to be clickable
    search_input = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@id='glmap-search-query']"))
    )
    # Clear the search input field
    search_input.clear()
    try:
        # Execute JavaScript to clear the search input field
        driver.execute_script("arguments[0].value = '';", search_input)

        # Send keys to the cleared search input field
        search_input.send_keys(apn)

        try:

            # Find and click the "See all results" link using JavaScript
            all_results_link = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, '//p[@class="centered"]/a[@class="all-results"]')))
            driver.execute_script("arguments[0].click();", all_results_link)

            # Wait for the search results to load
            search_results = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="search-results"]')))

            # Find all headline links
            headline_links = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, './/div[@class="col-md-6"]//div[@class="headline parcel-result "]/a'))
            )

            # Flag to check if the property text is found in any link
            property_found = False

            # Check if property_text is not empty before using it
            if property_text and not isinstance(property_text, float):
                for link in headline_links:
                    if property_text in link.text:
                        # Use JavaScript click to avoid ElementClickInterceptedException
                        driver.execute_script("arguments[0].click();", link)
                        property_found = True
                        break  # Exit loop if the link is clicked
            else:
                print("Property text is empty for APN:", apn, "\n")
                apn_data = {
                    "APN": apn,
                    "Parcel Id": "Empty",
                    "Deeded Owner": "Empty",
                    "Total Parcel Value": "Empty",
                    "Centroid Coordinates": "Empty",
                    "Legal Description": "Empty",
                    "Calculated Acres": "Empty",
                    "FEMA Flood Zone Subtype": "Empty"
                }
                scraped_data.append(apn_data)
                time.sleep(2)
                driver.back()  # Navigate back to the previous page

                # Wait for the page to fully load after navigating back
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@id='glmap-search-query']"))
                )
                continue



            if not property_found:
                # If property text is not found in any link, handle it here
                print("APN element not found for APN:", apn, "\n")
                apn_data = {
                    "APN": apn,
                    "Parcel Id": "Not found",
                    "Deeded Owner": "Not found",
                    "Total Parcel Value": "Not found",
                    "Centroid Coordinates": "Not found",
                    "Legal Description": "Not found",
                    "Calculated Acres": "Not found",
                    "FEMA Flood Zone Subtype": "Not found"
                }
                scraped_data.append(apn_data)
                time.sleep(2)  # Introduce a wait time before navigating back
                driver.back()  # Navigate back to the previous page

                # Wait for the page to fully load after navigating back
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@id='glmap-search-query']"))
                )
                continue

            # Wait for the parcel details to load
            mywait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="property"]')))

        except TimeoutException:
            print("APN element not found for APN:", apn, "\n")
            apn_data = {
                "APN": apn,
                "Parcel Id": "Not found",
                "Deeded Owner": "Not found",
                "Total Parcel Value": "Not found",
                "Centroid Coordinates": "Not found",
                "Legal Description": "Not found",
                "Calculated Acres": "Not found",
                "FEMA Flood Zone Subtype": "Not found"
            }
            scraped_data.append(apn_data)
            continue

        finally:
            time.sleep(2)

        # Try to extract data if available
        try:
            # EXTRACT PARCEL ID
            try:
                parcel_id = driver.find_element(By.XPATH,
                                                "//td[text()='Parcel ID']/following-sibling::td[@class='value']").text

                # parcel_id = str(parcel_id)

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

        # Create a DataFrame from the scraped data with defined column headers
        scraped_df = pd.DataFrame(scraped_data, columns=column_headers)

        # # Print the DataFrame
        # print(scraped_df)

        print("Before saving CSV")
        # Save the DataFrame to a CSV file
        scraped_df.to_excel("scraped_data4.xlsx",startrow=629, index=False)
        print("After saving CSV")

    finally:
        time.sleep(2)

driver.quit()




