import csv

# Helper function to read CSV file and process the data
def read_and_process_file(file_path):
    data = {}
    with open(file_path, mode='r') as file:
        reader = csv.DictReader(file)
        headers = reader.fieldnames  # Get the headers
        print(f"Headers in {file_path}: {headers}")  # Print headers for debugging
        for row in reader:
            territory = row['\ufeffCategory'].strip()
            purchase_inventory = float(row['Purchase Inventory (₹)'].replace(',', '').strip())
            if territory not in data:
                data[territory] = 0
            data[territory] += purchase_inventory
    return data

input_files = ["fifth1.csv", "fifth2.csv", "fifth3.csv", "fifth4.csv"]
territory_data_list = [read_and_process_file(file) for file in input_files]

all_territories = set()
for territory_data in territory_data_list:
    all_territories.update(territory_data.keys())

output_data = []
for territory in all_territories:
    row = [round(territory_data.get(territory, 0) / 100000, 2) for territory_data in territory_data_list]
    output_data.append(row)

column_sums = [sum(row[i] for row in output_data) for i in range(0, len(input_files))]

for row in output_data:
    row_sum = sum(row[0:])
    row.append(row_sum)
    last_value = row[3]
    last_but_one_value = row[4]
    ratio = round(((last_value / last_but_one_value)* 100 ),2) if last_but_one_value != 0 else 0
    row.append(ratio)

total_row = column_sums
grand_total = sum(column_sums)
total_row.append(grand_total)
grand_ratio = round(grand_total / column_sums[-1], 2) if column_sums[-1] != 0 else 0
total_row.append(grand_ratio)
output_data.append(total_row)

output_file_path = "output.csv"
with open(output_file_path, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerows(output_data)

print("Output CSV file generated successfully.")
