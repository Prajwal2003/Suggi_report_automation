import csv
import glob
import os
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import itertools
import requests
import json
from datetime import datetime, timedelta

pattern2 = "TBD*.csv"
pattern3 = "SBD*.csv"
pattern4 = "CBD*.csv"
pattern5 = "CI_*.csv"
pattern6 = "CIB_E*.csv"
pattern7 = "SI_*.csv"
pattern9 = "SIB_E*.csv"
pattern8 = "TWC*.csv"

matching_files2 = glob.glob(pattern2)
matching_files3 = glob.glob(pattern3)
matching_files4 = glob.glob(pattern4)
matching_files5 = glob.glob(pattern5)
matching_files6 = glob.glob(pattern6)
matching_files7 = glob.glob(pattern7)
matching_files8 = glob.glob(pattern8)
matching_files9 = glob.glob(pattern9)

matching_files2.sort()
matching_files3.sort()
matching_files4.sort()
matching_files5.sort()
matching_files7.sort()
matching_files6.sort()
matching_files8.sort()
matching_files9.sort()
print(matching_files2)
print(matching_files3)
print(matching_files4)
print(matching_files5)
print(matching_files6)
print(matching_files7)
print(matching_files8)
print(matching_files9)

def get_days_in_month(month):
    if month in [1, 3, 5, 7, 8, 10, 12]:
        return 31
    elif month in [4, 6, 9, 11]:
        return 30
    elif month == 2:
        return 29
    else:
        raise ValueError("Invalid month number. Must be between 1 and 12.")

def read_csv_columns(file_path, columns_to_extract, new_headers, fixed_values):
    data = []
    fixed_cycle = itertools.cycle(fixed_values)
    with open(file_path, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            new_row = {}
            for old_header, new_header in zip(columns_to_extract, new_headers):
                if old_header in row:
                    new_row[new_header] = row[old_header]
                else:
                    print(f"Column '{old_header}' not found in the CSV file.")
                    return None
            new_row["FixedColumn"] = next(fixed_cycle)
            data.append(new_row)
    return data

def read_csv(file_path):
    with open(file_path, mode='r') as file:
        reader = csv.DictReader(file)
        return [row for row in reader]

def read_csv_header(file_path, columns):
    with open(file_path, mode='r') as file:
        reader = csv.DictReader(file)
        return [{col: row[col] for col in columns if col in row} for row in reader]

def write_csv_edit(data, output_file):
    with open(output_file, mode='w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)

def write_csv(file_path, data, headers):
    with open(file_path, mode='w') as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        writer.writerows(data)
       
def write_csv_multiple(file_path, data, headers):
    with open(file_path, mode='w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        for i, row in enumerate(data):
            if i % 13 == 0:
                writer.writeheader()
            writer.writerow(row)
        
def read_and_process_file(file_path,header):
    data = {}
    with open(file_path, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            territory = row[header].strip()
            purchase_inventory = float(row['Purchase Inventory (₹)'].replace(',', '').strip())
            if territory not in data:
                data[territory] = 0
            data[territory] += purchase_inventory
    return data

def Suggi_first(matching_files4):
    url = "http://13.126.125.132:4000"

    payload = "{\"query\":\"query suggi_GetSaleWiseProfit { suggi_getSaleWiseProfit { _id customerDetails { name onboarding_date  village { name }  pincode { code } cust_type GSTIN iscouponapplicable address  { city street  pincode } phone customer_uid} storeDetails { name address {   pincode } vertical territory { name zone { name } } } userDetails { name address {   pincode } } invoiceno invoicedate product { soldqty servicecharge extradiscount sellingprice gst purchaseProductDetails {   name   category {     name   }   manufacturer {     name   }   sub_category {     name   } } lotDetails {  discount rate id sellingprice invoice { supplier_ref invoiceno createddate invoicedate } cnStockproduct subqty transportCharges otherCharges landingrate } supplierDetails { name } storeTargetPerProduct } payment { card cash upi } grosstotal storeCost } } \",\"variables\":{}}"
    headers = {
    'Content-Type': 'application/json'
    }
    
    matching_files4.sort()
    file_name = matching_files2[1]
    filename = os.path.basename(file_name)
    parts = filename.split("_")
    date_part = parts[2]
    parts = date_part.split("-")
    month_part = parts[1]
    date_part = parts[0]
    year = str(2024)
    start_date_str = "2024-" + month_part + "-" + date_part
    date_part = str(int(date_part) + 6)
    days = get_days_in_month(int(month_part))
    if int(date_part) > days:
        date_part = str(int(date_part) - days)
        month_part = str(int(month_part) + 1)
        if (int(month_part) == 13):
            month_part = str(int(month_part) - 12)
            year = str(year + 1)
    
    end_date_str = year + "-" + month_part + "-" + date_part
    
    response = requests.post(url, headers=headers, data=payload)
    data = response.json()["data"]["suggi_getSaleWiseProfit"]
    df = pd.DataFrame(data)


    def achieved_revenue(df, start_date_str, end_date_str):

        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        month = start_date.month
        year = start_date.year

        start_date_week = start_date
        start_date_month = datetime(2024, month, 1)
        start_date_year = datetime(year, 4, 1)

        start_date_week_str = start_date_week.strftime("%Y-%m-%dT00:00:00Z")
        start_date_month_str = start_date_month.strftime("%Y-%m-%dT00:00:00Z")
        start_date_year_str = start_date_year.strftime("%Y-%m-%dT00:00:00Z")

        end_date = (datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)).strftime("%Y-%m-%dT23:59:59Z")

        start_dates = [start_date_week_str, start_date_month_str, start_date_year_str]  

        revenue = []    
        for date_str in start_dates:    
            df['invoicedate'] = pd.to_datetime(df['invoicedate'])
            
            start_date = pd.to_datetime(date_str)
            end_date = pd.to_datetime(end_date)
            
            
            filtered_df = df[(df['invoicedate'] >= start_date) & (df['invoicedate'] <= end_date)]
            df1 = filtered_df.explode(['product'])
            df1['sellingprice_unit'] = df1['product'].apply(lambda x: x['sellingprice'])
            df1['servicecharge_unit'] = df1['product'].apply(lambda x: x['servicecharge'])
            df1['soldqty'] = df1['product'].apply(lambda x: x['soldqty'])

            df1['gst%_unit'] = df1['product'].apply(lambda x: x['gst'])
            df1['sellingprice_unit_xgst'] = df1['sellingprice_unit'] / (1 + (df1['gst%_unit'] / 100))
            df1['servicecharge_unit_xgst'] = df1['servicecharge_unit'] / 1.18
            df1['sale_value_xgst'] = df1['soldqty'] * (df1['sellingprice_unit_xgst'] + df1['servicecharge_unit_xgst'])
            sum = (df1['sale_value_xgst'].sum())
            revenue.append(sum)
            
        revenue_float = [round((float(x))/100000, 2) for x in revenue]    
        return(revenue_float)

    def target_revenue(start_date_str, end_date_str):
        url = "http://13.126.125.132:4000"

        payload = "{\"query\":\"query Suggi_StoreTarget {\\n  suggi_StoreTarget {\\n    Store\\n    TM\\n    Category\\n    Month\\n    Year\\n    Target\\n    Date\\n    Daily_Target\\n  }\\n}\",\"variables\":{}}"
        headers = {
            'Content-Type': 'application/json'
        }

        response = requests.post(url, headers=headers, data=payload)
        data = response.json()["data"]["suggi_StoreTarget"]
        df = pd.DataFrame(data)

        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        month = start_date.month
        year = start_date.year

        start_date_week = start_date
        start_date_month = datetime(2024, month, 1)
        start_date_year = datetime(year, 4, 1)

        start_date_week_str = start_date_week.strftime("%Y-%m-%dT00:00:00Z")
        start_date_month_str = start_date_month.strftime("%Y-%m-%dT00:00:00Z")
        start_date_year_str = start_date_year.strftime("%Y-%m-%dT00:00:00Z")

        end_date = (datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)).strftime("%Y-%m-%dT23:59:59Z")

        start_dates = [start_date_week_str, start_date_month_str, start_date_year_str]  

        revenue = []    
        for date_str in start_dates:    
            df['Date'] = pd.to_datetime(df['Date'])
            
            start_date = pd.to_datetime(date_str)
            end_date = pd.to_datetime(end_date)
            
            
            filtered_df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
            sum = (filtered_df['Daily_Target'].sum())
            revenue.append(sum)
            
        revenue_float = [round((float(x))/100000, 2) for x in revenue]
        return(revenue_float)

    def actual_margin(df, start_date_str, end_date_str):

        df['invoicedate'] = pd.to_datetime(df['invoicedate'])

        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        month = start_date.month
        year = start_date.year

        start_date_week = start_date
        start_date_month = datetime(2024, month, 1)
        start_date_year = datetime(year, 4, 1)

        start_date_week_str = start_date_week.strftime("%Y-%m-%dT00:00:00Z")
        start_date_month_str = start_date_month.strftime("%Y-%m-%dT00:00:00Z")
        start_date_year_str = start_date_year.strftime("%Y-%m-%dT00:00:00Z")

        end_date = (datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)).strftime("%Y-%m-%dT23:59:59Z")

        start_dates = [start_date_week_str, start_date_month_str, start_date_year_str]

        actual_margins = []

        for date_str in start_dates:
            start_date = pd.to_datetime(date_str)
            end_date = pd.to_datetime(end_date)

            filtered_df = df[(df['invoicedate'] >= start_date) & (df['invoicedate'] <= end_date)]
            
            df_exploded = filtered_df.explode('product')
            
            df_product = pd.json_normalize(df_exploded['product'])
            
            df_product['sellingprice_unit_xgst'] = df_product['sellingprice'] / (1 + (df_product['gst'] / 100))
            df_product['servicecharge_unit_xgst'] = df_product['servicecharge'] / (1 + (df_product['gst'] / 100))
            df_product['sale_value_xgst'] = df_product['soldqty'] * (df_product['sellingprice_unit_xgst'] + df_product['servicecharge_unit_xgst'])
            
            df_product['purchasingprice_unit'] = df_product['lotDetails.rate']
            df_product['lotDetails.discount'] = df_product['lotDetails.discount']
            df_product['purchase_value_xgst'] = df_product['soldqty'] * (df_product['purchasingprice_unit'] - df_product['lotDetails.discount'])
            
            df_product['actual_margin'] = df_product['sale_value_xgst'] - (
                df_product['purchase_value_xgst'] -
                (df_product['soldqty'] * (df_product['lotDetails.cnStockproduct'] / df_product['lotDetails.subqty'])) +
                (df_product['soldqty'] * (df_product['lotDetails.transportCharges'] / df_product['lotDetails.subqty'])) +
                (df_product['soldqty'] * (df_product['lotDetails.otherCharges'] / df_product['lotDetails.subqty']))
            )
            
            actual_margins.append(df_product['actual_margin'].sum())

        actual_margins_float = [round((float(x))/100000, 2) for x in actual_margins]
        percentages = [str(num) + "%" for num in actual_margins_float]
        return(percentages)
  
    def achieved_purchase(df, start_date_str, end_date_str):
      df['invoicedate'] = pd.to_datetime(df['invoicedate'])

      start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
      month = start_date.month
      year = start_date.year

      start_date_week = start_date
      start_date_month = datetime(2024, month, 1)
      start_date_year = datetime(year, 4, 1)
      
      start_date_week_str = start_date_week.strftime("%Y-%m-%dT00:00:00Z")
      start_date_month_str = start_date_month.strftime("%Y-%m-%dT00:00:00Z")
      start_date_year_str = start_date_year.strftime("%Y-%m-%dT00:00:00Z")
      
      end_date = (datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)).strftime("%Y-%m-%dT23:59:59Z")
      start_dates = [start_date_week_str, start_date_month_str, start_date_year_str]

      achieve_purchase = []

      for date_str in start_dates:
        start_date = pd.to_datetime(date_str)
        end_date = pd.to_datetime(end_date)

        filtered_df = df[(df['invoicedate'] >= start_date) & (df['invoicedate'] <= end_date)]
        df_exploded = filtered_df.explode('product')
        df_product = pd.json_normalize(df_exploded['product'])
        
        df_product['achived_purchase'] = df_product['soldqty'] * (df_product['lotDetails.rate'])
        achieve_purchase.append(df_product['achived_purchase'].sum())
      
      achieve_purchase_float = [round((float(x))/100000, 2) for x in achieve_purchase]
      return(achieve_purchase_float)
    
    ach_rev = achieved_revenue(df,start_date_str, end_date_str)
    tar_rev = target_revenue(start_date_str, end_date_str)
    act_mar = actual_margin(df,start_date_str, end_date_str)
    ach_pur = achieved_purchase(df,start_date_str, end_date_str)

    print(ach_rev)
    print(tar_rev)
    print(act_mar)
    print(ach_pur)

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)

    client = gspread.authorize(creds)

    spreadsheet_name = 'Weekly Report Siri Suggi(1)'

    sheet = client.open(spreadsheet_name)
    worksheet = sheet.get_worksheet(16)

    cells = ['C2', 'F2', 'I2']
    for x,i in zip(cells, ach_rev):
            worksheet.update(values=[[i]], range_name=x)

    cells = ['B2', 'E2', 'H2']
    for x,i in zip(cells, tar_rev):
            worksheet.update(values=[[i]], range_name=x)

    cells = ['C5', 'F5', 'I5']    
    for x,i in zip(cells, act_mar):
            worksheet.update(values=[[i]], range_name=x)
            
    cells = ['C3', 'F3', 'I3']    
    for x,i in zip(cells, ach_pur):
            worksheet.update(values=[[i]], range_name=x)
    
    filename = os.path.basename(file_name)
    parts = filename.split("_")
    num = parts[2].split("-")
    x = ( num[0] + "-" + str(int(num[0])+6))
    header = []
    header.append(x + ' Target')
    header.append(x + ' Achieved')
    cells = ['B1', 'C1']    
    for x,i in zip(cells, header):
            worksheet.update(values=[[i]], range_name=x)    
        
def Suggi_second(matching_files2):
    matching_files2.sort()
    data = []
    for file in matching_files2:
        print(file)
        data.append(read_csv(file))

    output_data = []
    column_sums = {
        '15-21 Target​': 0,
        '15-21 Achieved​': 0,
        'Achieved % ​': 0,
        '15-21 July Margin​': 0,
        'MTD Target July-24​': 0,
        'MTD Achieved June-24​': 0,
        'Achievement for Month %​': 0,
        'MTD Margin​': 0,
        'Target YTD​': 0,
        'Achieve YTD​': 0,
        'Achievement YTD %​': 0,
        'YTD Margin​': 0
    }

    for i in range(len(data[0])):
        target_15_21 = round(float(data[1][i]['Sales Target (₹)']) / 100000, 2)
        achieved_15_21 = round(float(data[1][i]['Total Sales (₹)']) / 100000, 2)
        achieved_percent = round((achieved_15_21 / target_15_21) * 100 if target_15_21 != 0 else 0, 2)
        margin_15_21 = round(float(data[1][i]['Gross Margin']) / 100000, 2)
        
        mtd_target = round(float(data[0][i]['Sales Target (₹)']) / 100000, 2)
        mtd_achieved = round(float(data[0][i]['Total Sales (₹)']) / 100000, 2)
        mtd_achieved_percent = round((mtd_achieved / mtd_target) * 100 if mtd_target != 0 else 0, 2)
        mtd_margin = round(float(data[0][i]['Gross Margin']) / 100000, 2)
        
        target_ytd = round(float(data[2][i]['Sales Target (₹)']) / 100000, 2)
        achieved_ytd = round(float(data[2][i]['Total Sales (₹)']) / 100000, 2)
        achieved_ytd_percent = round((achieved_ytd / target_ytd) * 100 if target_ytd != 0 else 0, 2)
        ytd_margin = round(float(data[2][i]['Gross Margin']) / 100000, 2)
        
        row_data = {
            '15-21 Target​': target_15_21,
            '15-21 Achieved​': achieved_15_21,
            'Achieved % ​': achieved_percent,
            '15-21 July Margin​': margin_15_21,
            'MTD Target July-24​': mtd_target,
            'MTD Achieved June-24​': mtd_achieved,
            'Achievement for Month %​': mtd_achieved_percent,
            'MTD Margin​': mtd_margin,
            'Target YTD​': target_ytd,
            'Achieve YTD​': achieved_ytd,
            'Achievement YTD %​': achieved_ytd_percent,
            'YTD Margin​': ytd_margin
        }

        output_data.append(row_data)

        for key in column_sums:
            column_sums[key] += row_data[key]

    column_sums['Achieved % ​'] = float(column_sums['15-21 Achieved​'])/float(column_sums['15-21 Target​'])
    column_sums['Achievement for Month %​'] = float(column_sums['MTD Achieved June-24​'])/float(column_sums['MTD Target July-24​'])
    column_sums['Achievement YTD %​'] = float(column_sums['Achieve YTD​'])/float(column_sums['Target YTD​'])
    sum_row = {key: f"{value:.2f}" for key, value in column_sums.items()}
    output_data.append(sum_row)

    x = "15-21"
    y = "June-24"
    output_file = 'Suggi_2 Report.csv'
    headers = [
         x + ' Target​', x + ' Achieved​', 'Achieved % ​', x + ' July Margin​',
        'MTD Target July-24​', 'MTD Achieved June-24​', 'Achievement for Month %​', 'MTD Margin​',
        'Target YTD​', 'Achieve YTD​', 'Achievement YTD %​', 'YTD Margin​'
    ]
    write_csv(output_file, output_data, headers)
    
def Suggi_third(matching_files3):
    matching_files3.sort()
    data = []
    for file in matching_files3:
        print(file)
        data.append(read_csv(file))

    output_data = []
    for i in range(len(data[0])):
        Sales_Target = round(float(data[1][i]['Sales Target (₹)']) / 100000, 2)
        Total_Sales = round(float(data[1][i]['Total Sales (₹)']) / 100000, 2)
        achieved_percent = round((Total_Sales / Sales_Target) *100 if Sales_Target != 0 else 0, 2)
        
        mtd_target = round(float(data[0][i]['Sales Target (₹)']) / 100000, 2)
        mtd_achieved = round(float(data[0][i]['Total Sales (₹)']) / 100000, 2)
        mtd_achieved_percent = round((mtd_achieved / mtd_target) * 100 if mtd_target != 0 else 0, 2)
        
        target_ytd = round(float(data[2][i]['Sales Target (₹)']) / 100000, 2)
        achieved_ytd = round(float(data[2][i]['Total Sales (₹)']) / 100000, 2)
        achieved_ytd_percent = round((achieved_ytd / target_ytd) * 100 if target_ytd != 0 else 0, 2)
        ytd_margin = round(float(data[2][i]['Gross Margin']) / 100000, 2)
        
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
        if (len(output_data) == 13 or len(output_data) == 27):
            output_data.append({
            'Sales Target​': 0,
            'Total Sales​': 0,
            'Achieved % ​': 0,
            'MTD Target July-24​': 0,
            'MTD Achieved June-24​': 0,
            'Achievement for Month %​': 0,
            'Target YTD​': 0,
            'Achieve YTD​': 0,
            'Achievement YTD %​': 0,
            'YTD Margin​': 0
        })
        
    headers = [
         'Sales Target​', 'Total Sales​', 'Achieved % ​',
        'MTD Target July-24​', 'MTD Achieved June-24​', 'Achievement for Month %​',
        'Target YTD​', 'Achieve YTD​', 'Achievement YTD %​', 'YTD Margin​'
    ]
    output_file = 'Suggi_3 Report.csv'
    write_csv(output_file, output_data, headers)

def Suggi_fourth(matching_files4):
    matching_files4.sort()
    data = []
    for file in matching_files4:
        print(file)
        data.append(read_csv(file))

    output_data = []
    column_sums = {
        '15-21 Target​': 0, '15-21 Achieved​': 0, 'Achieved % ​': 0, '15-21 July Margin​': 0,
        'MTD Target July-24​': 0, 'MTD Achieved June-24​': 0, 'Achievement for Month %​': 0, 'MTD Margin​': 0,
        'Target YTD​': 0, 'Achieve YTD​': 0, 'Achievement YTD %​': 0, 'YTD Margin​': 0
    }
    for i in range(len(data[0])):
        
        row_data = {
            '15-21 Target​': round(float(data[0][i]['Sales Target (₹)']) / 100000, 2),
            '15-21 Achieved​': round(float(data[0][i]['Total Sales (₹)']) / 100000, 2),
            'Achieved % ​': round((float(data[0][i]['Total Sales (₹)']) / float(data[0][i]['Sales Target (₹)'])) * 100 if float(data[0][i]['Sales Target (₹)']) != 0 else 0, 2),
            '15-21 July Margin​': round(float(data[0][i]['Gross Margin']) / 100000, 2),
            'MTD Target July-24​': round(float(data[2][i]['Sales Target (₹)']) / 100000, 2),
            'MTD Achieved June-24​': round(float(data[2][i]['Total Sales (₹)']) / 100000, 2),
            'Achievement for Month %​': round((float(data[2][i]['Total Sales (₹)']) / float(data[2][i]['Sales Target (₹)'])) * 100 if float(data[2][i]['Sales Target (₹)']) != 0 else 0, 2),
            'MTD Margin​': round(float(data[2][i]['Gross Margin']) / 100000, 2),
            'Target YTD​': round(float(data[1][i]['Sales Target (₹)']) / 100000, 2),
            'Achieve YTD​': round(float(data[1][i]['Total Sales (₹)']) / 100000, 2),
            'Achievement YTD %​': round((float(data[1][i]['Total Sales (₹)']) / float(data[1][i]['Sales Target (₹)'])) * 100 if float(data[1][i]['Sales Target (₹)']) != 0 else 0, 2),
            'YTD Margin​': round(float(data[1][i]['Gross Margin']) / 100000, 2)
        }
        output_data.append(row_data)
        
        for key in column_sums:
            column_sums[key] += row_data[key]

    column_sums['Achieved % ​'] = float(column_sums['15-21 Achieved​'])/float(column_sums['15-21 Target​'])
    column_sums['Achievement for Month %​'] = float(column_sums['MTD Achieved June-24​'])/float(column_sums['MTD Target July-24​'])
    column_sums['Achievement YTD %​'] = float(column_sums['Achieve YTD​'])/float(column_sums['Target YTD​'])
    sum_row = {key: f"{value:.2f}" for key, value in column_sums.items()}
    output_data.append(sum_row)

    x = "15-21"
    y = "June-24"
    output_file = 'Suggi_4 Report.csv'
    headers = [
        '15-21 Target​', '15-21 Achieved​', 'Achieved % ​', '15-21 July Margin​',
        'MTD Target July-24​', 'MTD Achieved June-24​', 'Achievement for Month %​', 'MTD Margin​',
        'Target YTD​', 'Achieve YTD​', 'Achievement YTD %​', 'YTD Margin​'
    ]
    write_csv(output_file, output_data, headers)
    
def Suggi_fifth(matching_files5):
    matching_files5.sort()
    for file in matching_files5:
        print(file)
        
    header = "\ufeffCategory"
    category_data_list = [read_and_process_file(file,header) for file in matching_files5]

    all_categories = set()
    for category_data in category_data_list:
        all_categories.update(category_data.keys())

    output_data = []
    all_categories = sorted(all_categories)
    for category in all_categories:
        row = [round(category_data.get(category, 0) / 100000, 2) for category_data in category_data_list]
        output_data.append(row)

    column_sums = [sum(row[i] for row in output_data) for i in range(0, len(matching_files5))]

    total_row = column_sums
    output_data.append(total_row)
    
    for row in output_data:
        row_sum = sum(row[0:])
        row.append(row_sum)
        last_value = row[3]
        last_but_one_value = row[4]
        ratio = round(((last_value / last_but_one_value) ),2) if last_but_one_value != 0 else 0
        row.append(ratio)

    output_file_path = "Suggi_5 Report.csv"
    with open(output_file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(output_data)
    
def Suggi_sixth(matching_files6):
    matching_files6.sort()
    data = []
    for file in matching_files6:
        print(file)
        data.append(read_csv(file))

    output_data = []
    column_sums = {
        'Purchase Inventory (₹)': 0
    }
    for i in range(len(data[0])):
        
        row_data = {
            'Purchase Inventory (₹)': round(float(data[0][i]['Purchase Inventory (₹)']) / 100000, 2),
        }
        output_data.append(row_data)
        
        for key in column_sums:
            column_sums[key] += row_data[key]
            
    sum_row = {key: f"{value:.2f}" for key, value in column_sums.items()}
    output_data.append(sum_row)

    output_file = 'Suggi_6 Report.csv'
    headers = [
        'Purchase Inventory (₹)'
    ]
    write_csv(output_file, output_data, headers)
    
def Suggi_seventh(matching_files7):
    matching_files7.sort()
    for file in matching_files7:
        print(file)
    header = "\ufeffterritory"
    territory_data_list = [read_and_process_file(file,header) for file in matching_files7]

    all_territories = set()
    for territory_data in territory_data_list:
        all_territories.update(territory_data.keys())

    output_data = []
    all_territories = sorted(all_territories)
    for territory in all_territories:
        row = [round(territory_data.get(territory, 0) / 100000, 2) for territory_data in territory_data_list]
        output_data.append(row)

    column_sums = [sum(row[i] for row in output_data) for i in range(0, len(matching_files7))]

    total_row = column_sums
    output_data.append(total_row)
    
    for row in output_data:
        row_sum = sum(row[0:])
        row.append(row_sum)
        last_value = row[3]
        last_but_one_value = row[4]
        ratio = round(((last_value / last_but_one_value) ),2) if last_but_one_value != 0 else 0
        row.append(ratio)

    output_file_path = "Suggi_7 Report.csv"
    with open(output_file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(output_data)

def Suggi_eight(matching_files8):
    matching_files8.sort()
    print(matching_files8)
    csv_file = matching_files8[0] 
    columns_to_extract = ["\ufeffcustomer", "territory"]  
    new_headers = ["Customer", "Territory"]  
    fixed_values = ['5819', '4367', '5852', '4303']

    file_path = os.path.join(csv_file)

    if os.path.exists(file_path):
        extracted_data = read_csv_columns(file_path, columns_to_extract, new_headers, fixed_values)
        if extracted_data:
            write_csv_edit(extracted_data, "extracted_data.csv")
        else:
            print(f"File not found: {file_path}")
            
    data = []
    file = "extracted_data.csv"
    if os.path.exists(file):
        with open(file, mode='r', encoding='utf-8-sig') as f:
            csv_reader = csv.DictReader(f)
            for row in csv_reader:
                data.append(row)
    
    output_data = []
    for i in range(len(data)):
        customer = float(data[i]['Customer'])
        target = float(data[i]['FixedColumn'])
        achieved_percentage = (customer / target)
        output_data.append({
            'Target': target,
            'Customer': customer,
            'Achieved %': achieved_percentage
        })
    
    output_file_path = "Suggi_8 Report.csv"
    with open(output_file_path, mode='w', newline='') as file:
        fieldnames = ['Target', 'Customer', 'Achieved %']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writerows(output_data)
     
def Suggi_ninth(matching_files9):
    matching_files9.sort()
    data = []
    for file in matching_files9:
        print(file)
        data.append(read_csv(file))

    output_data = []
    column_sums = {
        'Purchase Inventory (₹)': 0
    }
    for i in range(len(data[0])):
        
        row_data = {
            'Purchase Inventory (₹)': round(float(data[0][i]['Purchase Inventory (₹)']) / 100000, 2),
        }
        output_data.append(row_data)
        
        for key in column_sums:
            column_sums[key] += row_data[key]
            
    sum_row = {key: f"{value:.2f}" for key, value in column_sums.items()}
    output_data.append(sum_row)

    output_file = 'Suggi_9 Report.csv'
    headers = [
        'Purchase Inventory (₹)'
    ]
    write_csv(output_file, output_data, headers)

print("---------------------------------")
Suggi_first(matching_files4)
print("---------------------------------")
Suggi_second(matching_files2)
print("---------------------------------")
Suggi_third(matching_files3)
print("---------------------------------")
Suggi_fourth(matching_files4)
print("---------------------------------")
Suggi_fifth(matching_files5)
print("---------------------------------")
Suggi_sixth(matching_files6)
print("---------------------------------")
Suggi_seventh(matching_files7)
print("---------------------------------")
Suggi_eight(matching_files8)
print("---------------------------------")
Suggi_ninth(matching_files9)
print("---------------------------------")
print("Reports Generated")
print("---------------------------------")
    
def get_date_range_for_header():
    matching_files4.sort()
    file_path = matching_files4[1]
    filename = os.path.basename(file_path)
    parts = filename.split("_")
    num = parts[2].split("-")
    end_date = str(int(num[0]) + 6)
    days = get_days_in_month(int(num[1]))
    if int(end_date) > days:
        end_date = str(int(end_date) - days)
    date_range = ( num[0] + " - " + end_date)
    return(date_range)
    
def update_in_sheets():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
    client = gspread.authorize(credentials)

    spreadsheet = client.open("Weekly Report Siri Suggi(1)")

    pattern_for_report = "*Report.csv"
    matching_files = glob.glob(pattern_for_report)
    sheet_names = ['Suggi_second_report', 'Suggi_third_report', 'Suggi_fourth_report', 'Suggi_fifth_report',
                   'Suggi_sixth_report', 'Suggi_seventh_report', 'Suggi_eigthth_report', 'Suggi_ninth_report']
    matching_files.sort()

    print("---------------------------------")
    print(f"Updated Suggi_first_report values to Suggi_first_report")
    print("---------------------------------")
    for i, file in enumerate(matching_files):
        with open(file, 'r') as file_obj:
            reader = csv.reader(file_obj)
            content = list(reader)
            
            print("---------------------------------")
            print(f"Updating {file} to {sheet_names[i]}")
            
            worksheet = spreadsheet.worksheet(sheet_names[i])
            
            worksheet.batch_clear(["B2:ZZ1000"])
            date_range = get_date_range_for_header()
            header = []
            header.append(date_range + ' Target')
            header.append(date_range + ' Achieved')
            if i < 3:    
                cells = ['B1', 'C1']    
                for x,i in zip(cells, header):
                        worksheet.update(values=[[i]], range_name=x)
            
            if content:
                num_rows = len(content)
                num_cols = len(content[0]) if content else 0

                updated_content = []

                for index, row in enumerate(content):
                    if i == 1 and (index == 12 or index == 25):
                        updated_content.append([''] * num_cols)
                    updated_content.append([float(item) for item in row])

                update_range = f'B2:{chr(65 + num_cols)}{len(updated_content) + 1}'
                worksheet.update(values=updated_content, range_name=update_range)
                
                print(f"Updated {len(updated_content)} rows and {num_cols} columns")
                print("---------------------------------")
            else:
                print(f"{file} is a empty file")
    
    worksheet = spreadsheet.worksheet("Suggi_third_report")
    source_row_index = 1 
    destination_row_index = [] 
    destination_row_index.append(int(15))
    destination_row_index.append(int(29))
    
    source_row = worksheet.row_values(source_row_index)

    for x in destination_row_index :
        if source_row:
            
            num_cols = len(source_row)
            destination_range = f'A{x}:{chr(64 + num_cols)}{x}'
            
            worksheet.update(destination_range, [source_row])
        else:
            print(f"Header in third report are not added {source_row_index}")

    print("---------------------------------")
    print("Report is updated successfully")
    print("---------------------------------")

update_in_sheets()