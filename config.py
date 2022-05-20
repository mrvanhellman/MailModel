import pandas as pd
import csv

# Import and clean the database for use
stores = pd.read_excel('UnidadesRech.xlsx', usecols="A,C,D,F:H")
first_row = list(stores.iloc[0])
stores = stores.drop(0)
stores = stores.drop(4)
stores = stores.drop(5)
stores = stores.drop(48)
stores = stores.drop(49)
stores.columns = first_row

# Save the cleaned db in a csv file
with open('stores.csv', 'w', newline='') as csvfile:
    row_writer = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
    for index, rows in stores.iterrows():
        listRows = [rows[0], rows[1], rows[2], rows[3], rows[4], rows[5]]
        row_writer.writerow(listRows)

print(stores)
print("It's done")
