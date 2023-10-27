import pandas as pd
import json

# Load the Excel file into a DataFrame
df = pd.read_excel('kundentag.xlsx', nrows=67)

# Group customers by 'Company' and convert to a JSON structure
grouped = df.groupby('Unternehmen').apply(lambda group: group[['Name', 'Vorname']].to_dict('records')).reset_index(name='Customers')

# Create a list of JSON objects for each company
company_json_list = []
for index, row in grouped.iterrows():
    company_name = row['Unternehmen']
    customers = row['Customers']
    company_data = {
        "Company": company_name,
        "Customers": customers
    }
    company_json_list.append(company_data)

# Convert the list of JSON objects to a JSON string
json_string = json.dumps(company_json_list, ensure_ascii=False, indent=4)

with open('CustomerJson.json', 'w', encoding='utf-8') as file:
    file.write(json_string)
# Alternatively, to print the JSON string to the console
print(json_string)
