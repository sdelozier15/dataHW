import csv
import os

cereal_csv = os.path.join('pythonwork', 'cereal.csv') # ../Resources/cereal.csv

with open(cereal_csv, newline='') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')

    csv_header = next(csvfile, None)
    print(f"CSV: {csv_header}")

    for row in csvreader:

        if float(row[7]) >= 5:
            print(row)


