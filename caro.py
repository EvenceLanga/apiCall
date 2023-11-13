import json
import pandas as pd
import http.client

# Read Excel data
excel_file_path = "C:\\Users\\Evence Langa\\Downloads\\November MTD.xlsx"

df = pd.read_excel(excel_file_path)

# AEX API endpoint
aex_api_base_url = 'api.fno.za.aex.systems'
aex_api_token = 'eyJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJhZXgtY29yZSIsImlhdCI6MTY5MTA1MDY1NSwidWlkIjoiNTdBNzkxQzEtMkRFQi00NTVELUFCNUQtM0JBOEEzNzI2NjgwIiwiY2xhaW1zIjpbInVzZXJfaWQ6NTdBNzkxQzEtMkRFQi00NTVELUFCNUQtM0JBOEEzNzI2NjgwIiwib3BlcmF0b3JfaWQ6QUVFRjM4QzAtRjFFOS00RTgzLUFENEYtM0UzNUFBQjhERkNGIl19.KfBHYL0Tx2m9YbO5twY6EnscZm7PgziJUstRjR7DgGbWzMxZechCk2w9F43rrKkHECLrpiReYE4ot44-rHHpdw'

# Function to fetch status from API using http.client
def fetch_status(guid):
    try:
        connection = http.client.HTTPSConnection(aex_api_base_url)
        url = f'/work-orders/{guid}'
        headers = {'Authorization': f'Bearer {aex_api_token}'}
        connection.request('GET', url, headers=headers)
        response = connection.getresponse()
        data = response.read().decode('utf-8')
        data_json = json.loads(data)
        
        # Extract specific fields from the JSON response
        status = data_json.get('status', 'status not found')
        
        return {'status': status}
    except Exception as e:
        return {'status': f"Error: {str(e)}"}

# Extract last 36 characters from "WO URL" column
df['Last_36_Chars'] = df['WO URL'].apply(lambda url: url[-36:])

# Iterate through Excel rows and fetch status for each Last_36_Char
status_data = df['Last_36_Chars'].apply(lambda last_36: fetch_status(last_36))

# Create a new DataFrame with the extracted status data
status_df = pd.DataFrame(list(status_data))  # Convert list of dictionaries to DataFrame

# Concatenate the new DataFrame with the original DataFrame
result_df = pd.concat([df, status_df], axis=1)

# Save updated DataFrame to a new Excel file
output_excel_path = 'output_file_with_status.xlsx'
result_df.to_excel(output_excel_path, index=False)

print(f"Results saved to {output_excel_path}")
