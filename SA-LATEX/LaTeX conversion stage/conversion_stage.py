import xlrd

# open the Workbook
workbook = xlrd.open_workbook("opendatazno2019.xlsx")

# open the worksheet
worksheet = workbook.sheet_by_index(0)

# iterate the rows and columns
for i in range(0, 5):
    for j in range(0, 3):
        # print the cell values with a tab space
        print(worksheet.cell_value(i, j), end="\t")
    print("")

import pandas as pd
import xlrd 
import csv
 
# open workbook by sheet index, optional - sheet_by_index()
sheet = xlrd.open_workbook("opendatazno2019.xlsx").sheet_by_index(0)
  
# writer object is created
column = csv.writer(open("opendatazno2019.csv", "w", newline=""))
  
# write the data into csv file
for row in range(sheet.nrows):
    # row by row write operation
    column.writerow(sheet.row_values(row))
  
# read csv file and convert into a dataframe object
df = pd.DataFrame(pd.read_csv("opendatazno2019.csv", dtype="unicode"))

# -------------------------- CLEANING DATA --------------------------

df = pd.DataFrame(pd.read_csv("opendatazno2019.csv", dtype="unicode"))

# convert string to float
df["UkrBall100"] = df["UkrBall100"].astype("float")

# remove all rows with NULL values:
cleaned_df = df.dropna(subset=["UkrBall100"])

# resetting indexes after removing rows from dataframe
cleaned_df.reset_index(drop=True, inplace=True)

# remove all rows with "0.0" values:
cleaned_df = cleaned_df.loc[cleaned_df["UkrBall100"] != 0.0]
cleaned_df.reset_index(drop=True, inplace=True)

import random

sample_size = 500

# create a list of all possible indexes
index = [i for i in range(0,cleaned_df.last_valid_index()+1)]

# create an empty dictionary
random_elements = {}

# add one column to a dictionary
elements = []
random_elements.update({"UkrBall100": elements})

for j in range(sample_size):
    random_index = random.choice(index)
    elements.append(cleaned_df.loc[random_index, "UkrBall100"])
    index.remove(random_index)

random_selected_df = pd.DataFrame(random_elements)