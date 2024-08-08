import csv
import glob

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

matching_files5 = glob.glob("fifth*.csv")
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
grand_total = sum(column_sums)
total_row.append(grand_total)
    
for row in output_data:
    row_sum = sum(row[0:])
    row.append(row_sum)
    last_value = row[3]
    last_but_one_value = row[4]
    ratio = round(((last_value / last_but_one_value) ),2) if last_but_one_value != 0 else 0
    row.append(ratio)

grand_ratio = round(grand_total / column_sums[-1], 2) if column_sums[-1] != 0 else 0
total_row.append(grand_ratio)
output_data.append(total_row)

output_file_path = "Suggi_5 Report.csv"
with open(output_file_path, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerows(output_data)