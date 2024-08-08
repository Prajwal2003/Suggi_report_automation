import csv
import glob
import os
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

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
        
def Suggi_second(matching_files2):
    data = []
    for file in matching_files2:
        print(file)
        data.append(read_csv(file))

    output_data = []
    column_sums = {'15-21 Target​': 0, '15-21 Achieved​': 0}
    
    for i in range(len(data[0])):
        target_15_21 = round(float(data[1][i]['Sales Target (₹)']) / 100000, 2)
        achieved_15_21 = round(float(data[1][i]['Total Sales (₹)']) / 100000, 2)
        
        output_data.append({
            '15-21 Target​': target_15_21,
            '15-21 Achieved​': achieved_15_21,
        })
        
        column_sums['15-21 Target​'] += target_15_21
        column_sums['15-21 Achieved​'] += achieved_15_21

    column_sums['15-21 Target​'] = column_sums['15-21 Target​']/len(data[0])
    output_data.append({
        '15-21 Target​': f"{column_sums['15-21 Target​']:.2f}",
        '15-21 Achieved​': f"{column_sums['15-21 Achieved​']:.2f}"
    })

    x = "15-21"
    y = "June-24"
    output_file = 'totlasd Report.csv'
    headers = [
         x + ' Target​', x + ' Achieved​'
    ]
    write_csv(output_file, output_data, headers)
    
Suggi_second(matching_files2)