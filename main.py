import requests
import openpyxl
import os
import time

start_time = time.time()

start_page = 0
end_page = 48

cookies = {
    'll': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCIsInR5cGUiOiJKV1QifQ.eyJpYXQiOjE3MTc3ODEyNTUsImV4cCI6MTcyMDIwMDQ1NSwiYXVkIjoiaHR0cHM6Ly93d3cubHVzaGEuY28iLCJpc3MiOiI1NzQ2MzMzOS0xMWRmLTQ0ZDUtYjIwZS0zM2EzNjZmMjNkZjIiLCJzdWIiOiJhbm9ueW1vdXMifQ.mkADou9u0qxFZ6l6uaYWar1sx4zVRf7_dFh0KcyI_bM',
}

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
        cookies=cookies,
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