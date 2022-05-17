import pandas as pd
import csv

stores = pd.read_excel('UnidadesRech.xlsx', usecols="A,C,D,F:H")
first_row = list(stores.iloc[0])
stores = stores.drop(0)
stores = stores.drop(21)
stores = stores.drop(22)
stores.columns = first_row

with open('stores.csv', 'w', newline='') as csvfile:
    row_writer = csv.writer(csvfile, delimiter=' ', quotechar='|', quoting=csv.QUOTE_NONNUMERIC)
    for index, rows in stores.iterrows():
        listRows = [rows[0], rows[1], rows[2], rows[3], rows[4], rows[5]]
        row_writer.writerow(listRows)

print('Feito')
