import csv
from dataclasses import dataclass, field
from typing import List, Tuple
import pandas as pd

@dataclass
class SalesData:
    PBIT_Percentage: float = field(metadata={"alias": "PBIT %"})
    Target_Achieved_Percentage: float = field(metadata={"alias": "Target Achieved %"})
    Sales_Target: float = field(metadata={"alias": "Sales Target (₹)"})
    Total_Sales: float = field(metadata={"alias": "Total Sales (₹)"})
    Total_Purchase: float = field(metadata={"alias": "Total Purchase (₹)"})
    Transactions: int
    Sold_Qty: int
    PBIT: float = field(metadata={"alias": "PBIT (₹)"})
    Gross_Margin: float
    Gross_Margin_Percentage: float = field(metadata={"alias": "Gross Margin %"})
    Territory: str

    def __post_init__(self):

        if not -100 <= self.PBIT_Percentage <= 100:
            raise ValueError("PBIT % must be between -100 and 100")
        
        if not 0 <= self.Target_Achieved_Percentage <= 100:
            raise ValueError("Target Achieved % must be between 0 and 100")
        
        if self.Sales_Target < 0:
            raise ValueError("Sales Target (₹) cannot be negative")
        
        if self.Total_Sales < 0:
            raise ValueError("Total Sales (₹) cannot be negative")
        

        if self.Total_Purchase < 0:
            raise ValueError("Total Purchase (₹) cannot be negative")
        

        if self.Transactions < 0:
            raise ValueError("Transactions cannot be negative")
        
        if self.Sold_Qty < 0:
            raise ValueError("Sold Qty cannot be negative")
        
        if self.PBIT < 0:
            raise ValueError("PBIT (₹) cannot be negative")
        
        if self.Gross_Margin < 0:
            raise ValueError("Gross Margin cannot be negative")
        
        if not 0 <= self.Gross_Margin_Percentage <= 100:
            raise ValueError("Gross Margin % must be between 0 and 100")
        
        if not self.Territory:
            raise ValueError("Territory cannot be empty")

def read_csv(file_path: str) -> Tuple[List[SalesData], List[Tuple[dict, str]]]:
    items = []
    errors = []

    with open(file_path, mode='r', newline='') as file:
        reader = csv.DictReader(file)

        print("Column names:", reader.fieldnames)
        
        for row in reader:
            try:
                row = {k.strip(): v.strip() for k, v in row.items()}  # Strip spaces from keys and values
                row["PBIT %"] = float(row["PBIT %"].strip('%'))
                row["Target Achieved %"] = float(row["Target Achieved %"].strip('%'))
                row["Sales Target (₹)"] = float(row["Sales Target (₹)"])
                row["Total Sales (₹)"] = float(row["Total Sales (₹)"])
                row["Total Purchase (₹)"] = float(row["Total Purchase (₹)"])
                row["Transactions"] = int(row["Transactions"])
                row["Sold Qty"] = int(row["Sold Qty"])
                row["PBIT (₹)"] = float(row["PBIT (₹)"])
                row["Gross Margin"] = float(row["Gross Margin"])
                row["Gross Margin %"] = float(row["Gross Margin %"].strip('%'))
                row["Territory"] = row["Territory"]

                item = SalesData(
                    PBIT_Percentage=row["PBIT %"],
                    Target_Achieved_Percentage=row["Target Achieved %"],
                    Sales_Target=row["Sales Target (₹)"],
                    Total_Sales=row["Total Sales (₹)"],
                    Total_Purchase=row["Total Purchase (₹)"],
                    Transactions=row["Transactions"],
                    Sold_Qty=row["Sold Qty"],
                    PBIT=row["PBIT (₹)"],
                    Gross_Margin=row["Gross Margin"],
                    Gross_Margin_Percentage=row["Gross Margin %"],
                    Territory=row["Territory"]
                )

                items.append(item)
            except ValueError as e:
                errors.append((row, str(e)))
            except KeyError as e:
                errors.append((row, f"Missing column: {e}"))

    return items, errors

file_path = 'SalesData.csv'
validated_data, validation_errors = read_csv(file_path)

if validation_errors:
    for error in validation_errors:
        print(f"Error in row {error[0]}: {error[1]}")
else:
    print("All data validated successfully!")


for validated_item in validated_data:
    print(validated_item)
    
validated_data.to_excel('topic_posts.xlsx', index=False)
topic.to_excel('topic_posts.xlsx', index=False)
print('Spreadsheet saved.')

