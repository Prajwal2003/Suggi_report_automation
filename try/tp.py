import csv

from dataclasses import dataclass, field

from typing import List, Tuple




@dataclass

class InventoryItem:

    Category: str

    Purchase_Inventory: float = field(metadata={"alias": "Purchase Inventory (₹)"})

    Quantity: int




    def __post_init__(self):

        # Validate Category

        if not self.Category:

            raise ValueError("Category cannot be empty")

        

        # Validate Purchase_Inventory

        if self.Purchase_Inventory < 0:

            raise ValueError("Purchase Inventory cannot be negative")

        

        # Validate Quantity

        if self.Quantity < 0:

            raise ValueError("Quantity cannot be negative")




def read_csv(file_path: str) -> List[InventoryItem]:

    items = []

    errors = []


    with open(file_path, mode='r', newline='') as file:

        reader = csv.DictReader(file)

        for row in reader:

            try:

                row["Purchase Inventory (₹)"] = float(row["Purchase Inventory (₹)"])

                row["Quantity"] = int(row["Quantity"])

                item = InventoryItem(

                    Category=row["Category"],

                    Purchase_Inventory=row["Purchase Inventory (₹)"],

                    Quantity=row["Quantity"]

                )

                items.append(item)

            except ValueError as e:

                errors.append((row, str(e)))




    return items, errors




# Read and validate the CSV file

file_path = 'Crop_Inventory_1-45.csv'

validated_data, validation_errors = read_csv(file_path)




# Print validation errors if any

if validation_errors:

    for error in validation_errors:

        print(f"Error in row {error[0]}: {error[1]}")

else:

    print("All data validated successfully!")




# For demonstration, print validated data

for validated_item in validated_data:

    print(validated_item)