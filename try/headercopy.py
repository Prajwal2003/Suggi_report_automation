import gspread
from oauth2client.service_account import ServiceAccountCredentials

def copy_row_in_sheet():
    # Define the scope and authenticate using the service account credentials
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
    client = gspread.authorize(credentials)

    # Open the Google Sheets document
    spreadsheet = client.open("Weekly Report Siri Suggi(1)")
    
    # Select the worksheet
    worksheet = spreadsheet.worksheet("Sheet 1")

    # Define the source and destination row indices
    source_row_index = 1  # Index of the row you want to copy (1-based index)
    destination_row_index = []  # Index of the destination row (1-based index)
    destination_row_index.append(int(15))
    destination_row_index.append(int(29))
    # Fetch the values from the source row
    source_row = worksheet.row_values(source_row_index)

    for x in destination_row_index :
        if source_row:
            # Set the range for the destination row
            num_cols = len(source_row)
            destination_range = f'A{x}:{chr(64 + num_cols)}{x}'
            
            # Update the destination row in the worksheet
            worksheet.update(destination_range, [source_row])

            print(f"Copied row {source_row_index} to row {x}")
        else:
            print(f"No data found in row {source_row_index}")

# Call the function
copy_row_in_sheet()
