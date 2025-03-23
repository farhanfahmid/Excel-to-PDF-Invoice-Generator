import pandas as pd

import glob #for loading multiple filepaths

filepaths = glob.glob("invoices/*.xlsx") #get the filepaths in a python list called "filepaths"
print(filepaths)

#now read each filepath from the list filepaths
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    print("\n")
