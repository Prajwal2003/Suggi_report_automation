import gspread
from oauth2client.service_account import ServiceAccountCredentials
import glob
import csv

def update_in_sheets():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
    print(credentials)
    client = gspread.authorize(credentials)

    spreadsheet = client.open("Weekly Report Siri Suggi(1)")

    pattern4 = "*Report.csv"
    matching_files = glob.glob(pattern4)
    sheet_names = ['Suggi_second_report', 'Suggi_third_report', 'Suggi_fourth_report', 'Suggi_fifth_report',
                   'Suggi_sixth_report', 'Suggi_seventh_report', 'Suggi_eigthth_report', 'Suggi_ninth_report']
    matching_files.sort()

    for i, file in enumerate(matching_files):
        with open(file, 'r') as file_obj:
            reader = csv.reader(file_obj)
            content = list(reader)
            
            print("---------------------------------")
            print(f"Updating {file} to {sheet_names[i]}")
            
            worksheet = spreadsheet.worksheet(sheet_names[i])
            
            worksheet.batch_clear(["B2:ZZ1000"])
            
            if content:
                num_rows = len(content)
                num_cols = len(content[0]) if content else 0

                # Initialize the updated content list
                updated_content = []

                for index, row in enumerate(content):
                    updated_content.append([float(item) for item in row])

                    # Insert a blank row after the 12th row for the second worksheet
                    if i == 1 and (index == 12 or index == 25):
                        updated_content.append([''] * num_cols)

                update_range = f'B2:{chr(65 + num_cols)}{len(updated_content) + 1}'
                worksheet.update(values=updated_content, range_name=update_range)
                
                print(f"Updated {len(updated_content)} rows and {num_cols} columns in {sheet_names[i]}")
                print("---------------------------------")
            else:
                print(f"No data to update in {file}")
    
    worksheet = spreadsheet.worksheet("Suggi_third_report")
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

    print("All files uploaded")
    print("---------------------------------")

update_in_sheets()
