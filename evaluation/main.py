import httpx
from openpyxl import load_workbook

# Path to the xlsx file
file_path = './eval_queries.xlsx'  # Ensure this path is correct

# Load the workbook and select the sheet
workbook = load_workbook(filename=file_path)
sheet = workbook['Sheet1']  # Replace 'Sheet1' with your actual sheet name if different

# Specify the column to read (e.g., 'A' for the first column)
column_letter = 'A'

# Define the URL
url = "https://rag-retrieve-pr-299.dev.knowledge.healthcare.elsevier.systems/retrieve_vector"

# Define the headers
headers = {
    "Content-Type": "application/json; charset=utf-8"
}
# Define the data
static_data = {
    "dsl_filter": {
        "terms": {
            "source_id": [
                "375706",
                "797588",
                "797586",
                "797505",
                "797506",
                "799812",
                "797589",
                "799811",
                "298363"
            ]
        }
    }
}
# Iterate over the rows in the specified column and make server requests
for cell in sheet[column_letter]:
    query_value = cell.value
    if query_value:
        # Construct the data with the dynamic query_text
        data = static_data.copy()
        data["query_text"] = query_value
        
         # Make the HTTP POST request
        response = httpx.post(url, headers=headers, json=data)
        
        # Print the response
        print(response.json()["query"])
        for result in response.json()["results"]:
            chunk_text = result["_source"]["chunk_text"]
            if len(chunk_text) <= 20:
                print("Chunk Text: {} and length: {}".format(chunk_text, len(chunk_text)))