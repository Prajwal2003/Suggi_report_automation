import os
import csv

def read_csv_columns(file_path, columns_to_extract, new_headers, fixed_values):
    data = []
    with open(file_path, mode='r', encoding='utf-8-sig') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            extracted_row = {new_headers[i]: row[columns_to_extract[i]] for i in range(len(columns_to_extract))}
            data.append(extracted_row)
    for i, fixed_value in enumerate(fixed_values):
        data[i]['FixedColumn'] = fixed_value
    return data

def write_csv_edit(data, output_file_path):
    if data:
        with open(output_file_path, mode='w', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=data[0].keys())
            writer.writerows(data)

def Suggi_eight(matching_files8):
    matching_files8.sort()
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
        achieved_percentage = (customer / target) * 100
        output_data.append({
            'Target': target,
            'Customer': customer,
            'Achieved %': achieved_percentage
        })
    
    print(output_data)
    output_file_path = "1trying.csv"
    with open(output_file_path, mode='w', newline='') as file:
        fieldnames = ['Target', 'Customer', 'Achieved %']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writerows(output_data)

matching_files8 = ['input_data.csv']
Suggi_eight(matching_files8)
