import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import traceback
import Web_login
import time
import my_mods  # type: ignore

# File paths
file_path = r'C:/Users/dsamu/dsamllc.net/dsamllc.net - Documents/FIS Project Documents/1Power Bi/Link Deviations.xlsx'
HOT_LINK_sheet_name = 'HOT LINK'
city_id_file_path = r'C:/Users/dsamu/dsamllc.net/dsamllc.net - Documents/FIS Project Documents/POP (Play or Pay)/CITY ID.xlsx'
pop_review_file_path = r'C:/Users/dsamu/dsamllc.net/dsamllc.net - Documents/FIS Project Documents/POP (Play or Pay)/POP review spreadsheet.xlsx'
ci_hub = 'City Id Hub'
ci_column = 'City Id'
sub_contractor_column_city_id = 'Sub-Contractor'
PayrollInfo = r'C:\\Users\\dsamu\\dsamllc.net\\dsamllc.net - Documents\\FIS Project Documents\\1Power Bi\\PayrollInfo.xlsx'
B2G_Names = 'B2G Names'
Column_Tier = 'Tier'
Column_Subcontractor = 'Subcontractor'

audit_dates = [
    "9/29/2024", "3/31/2024", "6/30/2024", "12/29/2024",
    "12/31/2023", "9/24/2023", "6/25/2023", "3/26/2023",
    "12/25/2022", "9/25/2022", "6/26/2022", "3/27/2022",
    "12/26/2021", "9/26/2021", "6/27/2021", "3/28/2021",
    "12/27/2020", "9/27/2020", "6/28/2020", "3/29/2020"
]

# Initialize driver and login
driver = Web_login.login_b2g()
print("Driver initialized and logged in successfully.")
Web_login.access_contract(driver)

time.sleep(120)

subcontractor_names = '[Tier 1] Network Cabling Services, Inc.'

# Select the subcontractor
dropdown_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[5]/div/button[@name='customSelectContractVendorID']")))
dropdown_button.click()
Web_login.select_company(driver, '[Tier 1] Network Cabling Services,')
Go_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='ButtonGo']")))
Go_button.click()
print(f"Successfully selected {'[Tier 1] Network Cabling Services, Inc.'}.")
time.sleep(120)
#testsssss
# Load PayrollInfo Excel for specified contractor's audit dates
sheet_name = 'Payroll Details (2)'
column_name_date = 'Payroll Date'
column_name_contractor = 'Contractor Name'
contractor_name = 'Network Cabling Services, Inc.'

# Load Excel and filter by contractor name
df = pd.read_excel(PayrollInfo, sheet_name=sheet_name)
df = df[df[column_name_contractor] == contractor_name]
filtered_audit_dates = set(df[column_name_date].dt.strftime('%m/%d/%Y'))  # Dates in format similar to website

# Locate table rows and process
rows = driver.find_elements(By.XPATH, "/html/body/form/div[7]/div[3]/div/table/tbody/tr")

for row in rows:
    try:
        # Extract the second date from each row
        date_cell = row.find_elements(By.TAG_NAME, "td")[1]
        date_text = date_cell.text.strip()

        # Check if the date is in audit_dates
        if date_text in audit_dates:
            continue  # Skip this row if the date is in audit_dates

        # Perform action if date not in audit_dates
        view_link = row.find_element(By.LINK_TEXT, "View")
        if view_link:
            view_link.click()
            # Add any further processing or data scraping here as needed
            print(f"Processed row with date {date_text}.")

    except Exception as e:
        print(f"Error processing row: {e}")
        traceback.print_exc()






# Close the browser
input("Press Enter to close the browser...")
driver.quit()















