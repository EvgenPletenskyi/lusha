import requests
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os
import time

start_time = time.time()

chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

chrome_driver_path = '/Users/a1/Downloads/chromedriver-mac-arm64/chromedriver'
service = Service(chrome_driver_path)

driver = webdriver.Chrome(service=service, options=chrome_options)

login_url = 'https://dashboard.lusha.com/prospecting/explore_companies'
username = 'purchase@johnsiskcontractors.com'
password = 'Laimas11,./'
driver.set_page_load_timeout(10)

try:
    driver.get(login_url)
    username_field = driver.find_element(By.NAME, "email")
    password_field = driver.find_element(By.NAME, "password")

    username_field.send_keys(username)
    password_field.send_keys(password)
    
    login_button = driver.find_element(By.XPATH, "//*[@id=\"__next\"]/div/div[1]/div[3]/div[2]/span/form/button")
    login_button.click()
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(10)
    driver.get('https://dashboard-services.lusha.com/v4/user-events')
    cookies = driver.get_cookies()
    time.sleep(10)
finally:
    driver.quit()

    start_page = 0
    end_page = 0

    cookies_jar = requests.cookies.RequestsCookieJar()
    for cookie in cookies:
        cookies_jar.set(cookie['name'], cookie['value'])

    print(cookies_jar
          )
    headers = {
        'content-type': 'application/json',
    }

    json_data = {
        'filters': {
            'companyIndustryLabels': [
                {
                    'value': 'Wholesale Building Materials',
                    'id': 138,
                    'mainIndustry': 'Wholesale',
                    'mainIndustryId': 20,
                    'subIndustriesCount': 2,
                },
            ],
            'contactLocation': [
                {
                    'country': 'italy',
                    'key': 'country',
                },
            ],
        },
        'display': 'companies',
        'pages': {
            'page': 0,
            'pageSize': 25,
        },
        'sessionId': '4a45a8dc-5b5c-40a1-95ab-2cf0c82b8f45',
        'searchTrigger': 'NewTab',
        'savedSearchId': 0,
        'bulkSearchCompanies': {},
        'isRecent': False,
        'isSaved': False,
        'pageAbove400': None,
        'totalPagesAbove400': 377,
        'excludeRevealedContacts': False,
        'intentPlgRollout': False,
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Companies"

    ws.append(["Company Name", "Website"])

    for page in range(start_page, end_page + 1):
        json_data['pages']['page'] = page

        response = requests.post(
            'https://dashboard-services.lusha.com/v2/prospecting-full',
            cookies=cookies_jar,
            headers=headers,
            json=json_data,
        )

        if response.status_code in [200, 201]:
            data = response.json()

            companies = data.get('companies', {})
            if companies:
                results = companies.get('results', [])
                for result in results:
                    industry_clustering = result.get('industry_clustering', {})
                    if industry_clustering:
                        website = industry_clustering.get('website')
                        name = industry_clustering.get('name')
                        ws.append([name, website])
            else:
                print(f"No companies data found on page {page}.")
        else:
            print(f"Failed to retrieve the data on page {page}. Status code: {response.status_code}")

    file_path = os.path.join(os.getcwd(), "companies.xlsx")
    wb.save(file_path)

    end_time = time.time()
    execution_time = end_time - start_time

    print(f"Data has been written to {file_path}")
    print(f"Execution time: {execution_time} seconds")