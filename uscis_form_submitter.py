from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd
import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-b', '--browser', help="Choose your browser", required=True, choices=['firefox', 'chrome'])
    parser.add_argument('-c', '--chunk', help="Choose chunk size, default is 5", default=5, required=False, choices=range(5, 10))
    parser.add_argument('-f', '--file-name', help="Absolute path to the excel sheet", default='uscis_info.xlsx', required=False)
    args = vars(parser.parse_args())
    workbook_df = pd.read_excel(args['file_name'], sheet_name='Sheet1')
    if args['browser'] == 'firefox':
        driver = webdriver.Firefox()
    else:
        driver = webdriver.Chrome()

    CHUNK_SIZE = args['chunk']
    parent_handle = driver.window_handles[0]
    print("***BELOW INFO IS OPTIONAL, PRESS ENTER TO SKIP***")
    first_name = input('Your first name \n')
    last_name = input('Your last name \n')
    email = input('Your email \n')
    phone = input('Your phone \n')
    chunked_df = [workbook_df[i:i+CHUNK_SIZE] for i in range(0,len(workbook_df),CHUNK_SIZE)]
    for chunk in chunked_df:
        for _, row in chunk.iterrows():
            driver.switch_to.window(parent_handle)
            driver.execute_script(f"window.open('about:blank');")

        index = 0
        for _, row in chunk.iterrows():
            driver.switch_to.window(driver.window_handles[index+1])
            driver.get('https://www.uscis.gov/report-fraud/uscis-tip-form')
            prior_submission = row['Prior Submissions']
            additional_info = row['Providing additional Info']
            violator_info = row['Violator Info']
            reporting_from = row['Reporting From']
            business_name = row['Business Name']
            first_name_of_violator = row['First Name']
            last_name_of_violator = row['Last Name']
            address = row['Address']
            address_type = row['Address Type']
            city = row['City']
            state = row['State']
            zip = row['Zip']
            country = row['Country']
            dob = row['DOB']
            a_number = row['A-Number']
            receipt_number = row['Receipt Number']
            summary = row['Summary']
            # points to ->  Have you previously submitted this information to USCIS? 
            elem = driver.find_element(By.ID, f"edit-have-you-previously-submitted-this-information-to-uscis-{prior_submission.lower()}")
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
                driver.find_element(By.ID, "edit-zip-code").send_keys(zip)
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
            index = index + 1

        input(f"***I have verified the contents of the form in the browser tabs. Press enter when done***".upper())
        for handle in driver.window_handles:
            if handle != parent_handle:
                driver.switch_to.window(handle)
                driver.close()

    driver.quit()

if __name__ == "__main__":
    main()