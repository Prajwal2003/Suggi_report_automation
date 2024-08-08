import csv
def read(x):
    with open(x, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            print(row)

read("data.csv")