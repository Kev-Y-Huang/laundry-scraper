from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import xlsxwriter


# CONSTANTS
DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
STARTING_LISTING = 0
NUM_LISTINGS = 50

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(
    "C:\\Users\\kevyh\\Downloads\\chromedriver_win32\\chromedriver.exe", options=options)
driver.maximize_window()

# Go to laundryview url
driver.get("https://www.laundryview.com/selectProperty")

workbook = xlsxwriter.Workbook('laundry.xlsx')

for i in range(STARTING_LISTING, NUM_LISTINGS):
    # Wait until we are on the laundryview page
    element = WebDriverWait(driver, 40).until(
        EC.presence_of_element_located((By.CLASS_NAME, "content-wrapper")))
    time.sleep(1)

    # Look up the HARVARD UNDERGRADUATE HOUSING listing and click on listing
    input_element = driver.find_element(
        By.XPATH, "//input[starts-with(@class, 'property-type-ahead')]")
    input_element.send_keys(Keys.CONTROL + "a")
    input_element.send_keys(Keys.DELETE)
    input_element.send_keys('HARVARD UNDERGRADUATE HOUSING')
    list_item = driver.find_element(
        By.XPATH, "//div[starts-with(@class, 'property-type-ahead-item')]")
    driver.execute_script("return arguments[0].scrollIntoView();", list_item)
    list_item.click()
    time.sleep(0.5)

    # Search for current house based on index
    list = driver.find_element(
        By.CSS_SELECTOR, "div[class='property-type-ahead-items']")
    houses = list.find_elements(
        By.CSS_SELECTOR, "div[class^='property-type-ahead-item']")

    # Get truncated name for listing
    house_name = houses[i + 1].text.split("> ")[1]
    house_name = house_name.replace("HOUSE ", "")
    house_name = house_name.split(" STUDENT")[0]
    worksheet = workbook.add_worksheet(house_name)
    print(i + 1, house_name)

    # Click on and route to house url
    houses[i + 1].click()
    time.sleep(0.5)

    # Route to weekly statistics url
    driver.get(driver.current_url.replace("home", "room"))
    time.sleep(0.5)

    # Write out days (as column headers) to be inserted into the spreadsheet
    for j, day in enumerate(DAYS):
        worksheet.write(0, j + 1, day)

    # Copy the weekly stats table over to the spreadsheet
    body = driver.find_element(By.TAG_NAME, "tbody")
    rows = body.find_elements(By.TAG_NAME, "tr")
    for j in range(len(rows)):
        row = rows[j].find_elements(By.TAG_NAME, "td")
        worksheet.write(j + 1, 0, row[0].text)

        for k in range(1, len(row)):
            # Usage percentage is stored in the opacity (if it exists) of the div
            percent = row[k].get_attribute("style")
            if percent:
                percent = percent[9:-1]
            else:
                percent = 0

            worksheet.write(j + 1, k, float(percent))

    # Reroute back to the main laundry view page
    driver.back()
    driver.back()

workbook.close()
