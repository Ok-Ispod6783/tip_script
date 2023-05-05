from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import argparse, time

timeout = 10

def autofillform(driver, row, index, first_name, last_name, email, phone):
    prior_submission = row['Prior Submissions'] if not pd.isna(row['Prior Submissions']) else 'N'
    additional_info = row['Prior Submissions'] if not pd.isna(row['Providing additional Info']) else 'N'
    violator_info = row['Violator Info']
    reporting_from = row['Reporting From'] if not pd.isna(row['Providing additional Info']) else 'INSIDE_US'
    business_name = row['Business Name'] if not pd.isna(row['Business Name']) else ''
    first_name_of_violator = row['First Name'] if not pd.isna(row['First Name']) else ''
    last_name_of_violator = row['Last Name'] if not pd.isna(row['Last Name']) else ''
    address = row['Address'] if not pd.isna(row['Address']) else ''
    address_type = row['Address Type']
    city = row['City'] if not pd.isna(row['City']) else ''
    state = row['State'] if not pd.isna(row['State']) else ''
    zip = row['Zip'] if not pd.isna(row['Zip']) else ''
    country = row['Country'] if not pd.isna(row['Country']) else ''
    dob = row['DOB']
    a_number = row['A-Number'] if not pd.isna(row['A-Number']) else ''
    receipt_number = row['Receipt Number'] if not pd.isna(row['Receipt Number']) else ''
    summary = row['Summary'] if not pd.isna(row['Summary']) else ''
    if city == '' or state == '':
        print(f"Skipping {violator_info} since state or city are mandatory fields")
    else:
        driver.execute_script("window.open('https://www.uscis.gov/report-fraud/uscis-tip-form');")
        driver.switch_to.window(driver.window_handles[index+1])
        # points to ->  Have you previously submitted this information to USCIS?
        prior_submission_id = f"edit-have-you-previously-submitted-this-information-to-uscis-{prior_submission.lower()}"
        element_present = EC.presence_of_element_located((By.ID, prior_submission_id))
        WebDriverWait(driver, timeout).until(element_present)
        # points to ->  Have you previously submitted this information to USCIS? 
        elem = driver.find_element(By.ID, prior_submission_id) 
        driver.execute_script("arguments[0].click();", elem)

        if prior_submission == 'Y':
            elem = driver.find_element(By.ID, f"edit-are-you-providing-additional-information-{additional_info.lower()}")
            driver.execute_script("arguments[0].click();", elem)
        
        # points to -> Your Information (Optional)
        driver.find_element(By.ID, "edit-first-name").send_keys(first_name)
        driver.find_element(By.ID, "edit-last-name").send_keys(last_name)
        driver.find_element(By.ID, "edit-email").send_keys(email)
        driver.find_element(By.ID, "edit-telephone").send_keys(phone)

        # points to ->  Where are you reporting from? 
        id_prefix = 'y' if reporting_from == 'INSIDE_US' else 'n'
        elem = driver.find_element(By.ID, f"edit-where-are-you-reporting-from-{id_prefix}")
        driver.execute_script("arguments[0].click();", elem)

        # points to ->  Suspected Fraud or Abuse
        driver.execute_script("document.getElementById('edit-violation').value = '3'")
        driver.execute_script("document.getElementById('select2-edit-violation-container').innerHTML = 'Employment Fraud - H-1B'")

        # points to -> Suspected Violator Information
        elem = driver.find_element(By.ID, f"edit-the-report-involves-{violator_info.lower()}")
        driver.execute_script("arguments[0].click();", elem)

        # points to -> Name and Location of Suspected Individual or Business
        if violator_info != 'INDIVIDUAL':
            # business or both
            driver.find_element(By.ID, "edit-business-name").send_keys(business_name)
            driver.find_element(By.ID, "edit-city").send_keys(city)
            if zip != '':
                driver.find_element(By.ID, "edit-zip-code").send_keys(str(int(zip)))
        elif violator_info != 'BUSINESS':
            # individual or both
            driver.find_element(By.ID, "edit-reported-first-name").send_keys(first_name_of_violator)
            driver.find_element(By.ID, "edit-reported-last-name").send_keys(last_name_of_violator)
            dob_split = dob.split('/')
            driver.find_element(By.ID, "edit-month").send_keys(dob_split[0])
            driver.find_element(By.ID, "edit-dd").send_keys(dob_split[1])
            driver.find_element(By.ID, "edit-year").send_keys(dob_split[2])
            driver.find_element(By.ID, "edit-alien-registration-number").send_keys(a_number)
            driver.find_element(By.ID, "edit-receipt-number").send_keys(receipt_number)
        elif violator_info == 'BOTH':
            driver.execute_script(f"document.getElementById('edit-address-type').value = '{address_type}'")
            driver.execute_script(f"document.getElementById('select2-edit-address-type-container').innerHTML = '{address_type}'")
        driver.find_element(By.ID, "edit-address").send_keys(address)
        driver.execute_script(f"document.getElementById('edit-state').value = '{state}'")
        driver.execute_script(f"document.getElementById('select2-edit-state-container').innerHTML = '{state}'")
        driver.execute_script(f"document.getElementById('edit-country').value = '{country}'")
        driver.execute_script(f"document.getElementById('select2-edit-country-container').innerHTML = '{country}'")
        # points to -> Summary of Suspected Fraud or Abuse
        driver.find_element(By.ID, "edit-abuse-description").send_keys(summary)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-b', '--browser', help="Choose your browser", required=True, choices=['firefox', 'chrome'])
    parser.add_argument('-c', '--chunk', help="Choose chunk size, default is 5", default=5, required=False, choices=range(5, 11), type=int)
    parser.add_argument('-f', '--file-name', help="Absolute path to the excel sheet", default='uscis_info.xlsx', required=False)
    parser.add_argument('-s', '--sheet-name', help="Sheet name within the file", required=True)
    parser.add_argument('-fname', '--first-name', required=False, help="First name", default='')
    parser.add_argument('-lname', '--last-name', required=False, help="Last name", default='')
    parser.add_argument('-e', '--email', required=False, help="email address", default='')
    parser.add_argument('-p', '--phone', required=False, help="Phone number", default='')
    args = vars(parser.parse_args())
    workbook_df = pd.read_excel(args['file_name'], sheet_name=args['sheet_name'])
    first_name = args['first_name']
    last_name = args['last_name']
    email = args['email']
    phone = args['phone']
    if args['browser'] == 'firefox':
        driver = webdriver.Firefox()
        driver.install_addon('buster.xpi')
        driver.install_addon('noptcha.xpi')
    else:
        options = webdriver.ChromeOptions()
        options.add_extension('./buster.crx')
        options.add_extension('./noptcha.crx')
        driver = webdriver.Chrome(options=options)

    CHUNK_SIZE = args['chunk']
    parent_handle = driver.window_handles[0]
    print("***BELOW INFO IS OPTIONAL, PRESS ENTER TO SKIP***")

    chunked_df = [workbook_df[i:i+CHUNK_SIZE] for i in range(0,len(workbook_df),CHUNK_SIZE)]
    for chunk in chunked_df:
        index = 0
        driver.switch_to.window(parent_handle)
        for _, row in chunk.iterrows():
            autofillform(driver, row, index, first_name, last_name, email, phone)
            index = index + 1
        input(f"***I have verified the contents of the form in the browser tabs. Press enter when done***".upper())
        for handle in driver.window_handles:
            if handle != parent_handle:
                driver.switch_to.window(handle)
                driver.close()

    driver.quit()

if __name__ == "__main__":
    main()