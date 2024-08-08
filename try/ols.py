import csv
import glob
import os
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

directory = "/Users/ashwyn/hello/input"
file_names = os.listdir(directory)

pattern1 = "first*.csv"
pattern2 = "second*.csv"
pattern3 = "third*.csv"
pattern4 = "fourth*.csv"

matching_files1 = glob.glob(pattern1)
matching_files2 = glob.glob(pattern2)
matching_files3 = glob.glob(pattern3)
matching_files4 = glob.glob(pattern4)

print("-----------------------")
print(matching_files1)
print(matching_files2)
print(matching_files3)
print(matching_files4)
print("-----------------------")

def read_csv(file_path):
    with open(file_path, mode='r') as file:
        reader = csv.DictReader(file)
        return [row for row in reader]

def write_csv(file_path, data, headers):
    with open(file_path, mode='w') as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        writer.writerows(data)
        
def Suggi_first(input_files):
    data = []
    for file in input_files:
        print(file)
        data.append(read_csv(file))

    print(data)
    output_data = []
    for i in range(0):
        target_15_21 = float(data[1][i]['Sales Target (₹)']) / 100000
        achieved_15_21 = float(data[1][i]['Total Sales (₹)']) / 100000
        achieved_percent = (achieved_15_21 / target_15_21) * 100 if target_15_21 != 0 else 0
        margin_15_21 = float(data[1][i]['Gross Margin']) / 100000
        
        mtd_target = float(data[0][i]['Sales Target (₹)']) / 100000
        mtd_achieved = float(data[0][i]['Total Sales (₹)']) / 100000
        mtd_achieved_percent = (mtd_achieved / mtd_target) * 100 if mtd_target != 0 else 0
        mtd_margin = float(data[0][i]['Gross Margin']) / 100000
        
        target_ytd = float(data[2][i]['Sales Target (₹)']) / 100000
        achieved_ytd = float(data[2][i]['Total Sales (₹)']) / 100000
        achieved_ytd_percent = (achieved_ytd / target_ytd) * 100 if target_ytd != 0 else 0
        ytd_margin = float(data[2][i]['Gross Margin']) / 100000
        
        output_data.append({
            '15-21 Target​': target_15_21,
            '15-21 Achieved​': achieved_15_21,
            'Achieved % ​': achieved_percent,
            '15-21 July Margin​': margin_15_21,
            'MTD Target July-24​': mtd_target,
            'MTD Achieved June-24​': mtd_achieved,
            'Achievement for Month %​': mtd_achieved_percent,
            'MTD Margin​': mtd_margin,
            'Target YTD​': target_ytd,
            'Achieve YTD​': achieved_ytd,
            'Achievement YTD %​': achieved_ytd_percent,
            'YTD Margin​': ytd_margin
        })

    x = "15-21"
    y = "June-24"
    output_file = 'Suggi_1 Report.csv'
    headers = [
         x + ' Target​', x + ' Achieved​', 'Achieved % ​', x + ' July Margin​',
        'MTD Target July-24​', 'MTD Achieved June-24​', 'Achievement for Month %​', 'MTD Margin​',
        'Target YTD​', 'Achieve YTD​', 'Achievement YTD %​', 'YTD Margin​'
    ]
    write_csv(output_file, output_data, headers)
        
def Suggi_second(matching_files2):
    data = []
    for file in matching_files2:
        print(file)
        data.append(read_csv(file))

    output_data = []
    for i in range(len(data[0])):
        target_15_21 = float(data[1][i]['Sales Target (₹)']) / 100000
        achieved_15_21 = float(data[1][i]['Total Sales (₹)']) / 100000
        achieved_percent = (achieved_15_21 / target_15_21) * 100 if target_15_21 != 0 else 0
        margin_15_21 = float(data[1][i]['Gross Margin']) / 100000
        
        mtd_target = float(data[0][i]['Sales Target (₹)']) / 100000
        mtd_achieved = float(data[0][i]['Total Sales (₹)']) / 100000
        mtd_achieved_percent = (mtd_achieved / mtd_target) * 100 if mtd_target != 0 else 0
        mtd_margin = float(data[0][i]['Gross Margin']) / 100000
        
        target_ytd = float(data[2][i]['Sales Target (₹)']) / 100000
        achieved_ytd = float(data[2][i]['Total Sales (₹)']) / 100000
        achieved_ytd_percent = (achieved_ytd / target_ytd) * 100 if target_ytd != 0 else 0
        ytd_margin = float(data[2][i]['Gross Margin']) / 100000
        
        output_data.append({
            '15-21 Target​': target_15_21,
            '15-21 Achieved​': achieved_15_21,
            'Achieved % ​': achieved_percent,
            '15-21 July Margin​': margin_15_21,
            'MTD Target July-24​': mtd_target,
            'MTD Achieved June-24​': mtd_achieved,
            'Achievement for Month %​': mtd_achieved_percent,
            'MTD Margin​': mtd_margin,
            'Target YTD​': target_ytd,
            'Achieve YTD​': achieved_ytd,
            'Achievement YTD %​': achieved_ytd_percent,
            'YTD Margin​': ytd_margin
        })

    x = "15-21"
    y = "June-24"
    output_file = 'Suggi_2 Report.csv'
    headers = [
         x + ' Target​', x + ' Achieved​', 'Achieved % ​', x + ' July Margin​',
        'MTD Target July-24​', 'MTD Achieved June-24​', 'Achievement for Month %​', 'MTD Margin​',
        'Target YTD​', 'Achieve YTD​', 'Achievement YTD %​', 'YTD Margin​'
    ]
    write_csv(output_file, output_data, headers)
    
def Suggi_third(matching_files3):
    data = []
    for file in matching_files3:
        print(file)
        data.append(read_csv(file))

    output_data = []
    for i in range(len(data[0])):
        Sales_Target = float(data[1][i]['Sales Target (₹)']) / 100000
        Total_Sales = float(data[1][i]['Total Sales (₹)']) / 100000
        achieved_percent = (Total_Sales / Sales_Target) *100 if Sales_Target != 0 else 0
        
        mtd_target = float(data[0][i]['Sales Target (₹)']) / 100000
        mtd_achieved = float(data[0][i]['Total Sales (₹)']) / 100000
        mtd_achieved_percent = (mtd_achieved / mtd_target) * 100 if mtd_target != 0 else 0
        
        target_ytd = float(data[2][i]['Sales Target (₹)']) / 100000
        achieved_ytd = float(data[2][i]['Total Sales (₹)']) / 100000
        achieved_ytd_percent = (achieved_ytd / target_ytd) * 100 if target_ytd != 0 else 0
        ytd_margin = float(data[2][i]['Gross Margin']) / 100000
        
        output_data.append({
            'Sales Target​': Sales_Target,
            'Total Sales​': Total_Sales,
            'Achieved % ​': achieved_percent,
            'MTD Target July-24​': mtd_target,
            'MTD Achieved June-24​': mtd_achieved,
            'Achievement for Month %​': mtd_achieved_percent,
            'Target YTD​': target_ytd,
            'Achieve YTD​': achieved_ytd,
            'Achievement YTD %​': achieved_ytd_percent,
            'YTD Margin​': ytd_margin
        })

    output_file = 'Suggi_3 Report.csv'
    headers = [
         'Sales Target​', 'Total Sales​', 'Achieved % ​',
        'MTD Target July-24​', 'MTD Achieved June-24​', 'Achievement for Month %​',
        'Target YTD​', 'Achieve YTD​', 'Achievement YTD %​', 'YTD Margin​'
    ]
    write_csv(output_file, output_data, headers)

def Suggi_fourth(input_files):
    data = []
    for file in input_files:
        print(file)
        data.append(read_csv(file))

    output_data = []
    for i in range(len(data[0])):
        target_15_21 = float(data[1][i]['Sales Target (₹)']) / 100000
        achieved_15_21 = float(data[1][i]['Total Sales (₹)']) / 100000
        achieved_percent = (achieved_15_21 / target_15_21) * 100 if target_15_21 != 0 else 0
        margin_15_21 = float(data[1][i]['Gross Margin']) / 100000
        
        mtd_target = float(data[0][i]['Sales Target (₹)']) / 100000
        mtd_achieved = float(data[0][i]['Total Sales (₹)']) / 100000
        mtd_achieved_percent = (mtd_achieved / mtd_target) * 100 if mtd_target != 0 else 0
        mtd_margin = float(data[0][i]['Gross Margin']) / 100000
        
        target_ytd = float(data[2][i]['Sales Target (₹)']) / 100000
        achieved_ytd = float(data[2][i]['Total Sales (₹)']) / 100000
        achieved_ytd_percent = (achieved_ytd / target_ytd) * 100 if target_ytd != 0 else 0
        ytd_margin = float(data[2][i]['Gross Margin']) / 100000
        
        output_data.append({
            '15-21 Target​': target_15_21,
            '15-21 Achieved​': achieved_15_21,
            'Achieved % ​': achieved_percent,
            '15-21 July Margin​': margin_15_21,
            'MTD Target July-24​': mtd_target,
            'MTD Achieved June-24​': mtd_achieved,
            'Achievement for Month %​': mtd_achieved_percent,
            'MTD Margin​': mtd_margin,
            'Target YTD​': target_ytd,
            'Achieve YTD​': achieved_ytd,
            'Achievement YTD %​': achieved_ytd_percent,
            'YTD Margin​': ytd_margin
        })

    x = "15-21"
    y = "June-24"
    output_file = 'Suggi_4 Report.csv'
    headers = [
         x + ' Target​', x + ' Achieved​', 'Achieved % ​', x + ' July Margin​',
        'MTD Target July-24​', 'MTD Achieved June-24​', 'Achievement for Month %​', 'MTD Margin​',
        'Target YTD​', 'Achieve YTD​', 'Achievement YTD %​', 'YTD Margin​'
    ]
    write_csv(output_file, output_data, headers)
    
Suggi_first(matching_files1)
print("-----------------------")
Suggi_second(matching_files2)
print("-----------------------")
Suggi_third(matching_files3)
print("-----------------------")
Suggi_fourth(matching_files4)
print("-----------------------")
print("Reports Generated")
print("-----------------------")


scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
    "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
print(credentials)
client = gspread.authorize(credentials)

spreadsheet = client.open("Weekly Report Siri Suggi(1)")

pattern4 = "*Report.csv"
matching_files = glob.glob(pattern4)
sheet_names = ['Suggi_first_report','Suggi_second_report','Suggi_third_report','Suggi_fourth_report']
i = 0
matching_files.sort()
print(matching_files)

for i, file in enumerate(matching_files):
    with open(file, "r") as file_obj:
        data = list(csv.DictReader(file_obj))
    worksheet = spreadsheet.worksheet(sheet_names[i])
    cell_list = []
    for row_num, row in enumerate(data):
        for col_num, value in enumerate(row):
            if row_num > 0 and col_num > 0:  
                cell_list.append(worksheet.cell(row=row_num + 2, col=col_num + 1)) 
                cell_list[-1].value = value

    if cell_list:
        worksheet.update_cells(cell_list)
        print(f"Updated {len(cell_list)} cells in {sheet_names[i]}")
    else:
        print(f"No cells to update in {sheet_names[i]}")

    print(f"Uploaded {file} (excluding row 1 and column 1)!")

print("All files uploaded!")