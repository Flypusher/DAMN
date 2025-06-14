"""
DAMN

Drosphila
Activity
Monitor (in excel)
Navigator"""

import openpyxl
import pandas as pd
from openpyxl import Workbook, load_workbook
import random


"""Importing DAM CSV to Excel"""
CSV_file = input("Type in CSV file name: ")
#Copy File URI, not name or path i.e. should like like file://D:"name".txt
excel_file = input("Type in what you want to name the excel sheet as: ")
excel_name = excel_file + ".xlsx"
df = pd.read_csv(CSV_file, delimiter='\t', header = None)
#makes csv file into a dataframe
# "\t" for tab separated data or "," for CSV for delimiter
###df.to_excel(excel_name, index = False, header = None)
#makes dataframe into excel with the excel name you made


"""Cutting out all data before the DAMs are warmed up i.e.  51 to 1"""
first_index = df[df.iloc[:, 3] == 1].index[0]
#iloc[:,3] = fourth column i.e. DAM warm up column D
df = df.iloc[first_index - 1:]
#only keeps data of row and below, with 0 as first row


"""Moves D and J // On/Off and Lights column to end"""
cols = df.columns.tolist()
columns_to_move = [cols[3], cols[9]] #i.e. columns D and E as index values
new_order = [col for col in cols if col not in columns_to_move] + columns_to_move
df = df[new_order]



"""deletes 5th column i.e. E and the 4 following ones including I"""
df.drop(columns=df.columns[3:8], inplace=True)


"""header"""
header = (
    ["Number", "Date", "Time"] + 
    [f'M{i}' for i in range(1, 33)] + 
    ["On/Off", "Light"])
df.columns = header
df.to_excel(excel_name, index=False, header = True)



#consequtive finder
result_data = {
    "Monitor": [],
    "Time of 3rd >0": []
}

for monitor in [f"M{i}" for i in range(1,33)]:
    series = df[monitor]
    for i in range(len(series) - 2):
        if series.iloc[i] > 0 and series.iloc[i+1] > 0 and series.iloc[i+2] >0:
            result_data["Monitor"].append(monitor)
            result_data["Time of 3rd >0"].append(df["Number"].iloc[i+2])
            break
    else:
        result_data["Monitor"].append(monitor)
        result_data["Time of 3rd >0"].append("Not Found")
        
result_df = pd.DataFrame(result_data)
with pd.ExcelWriter(excel_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    result_df.to_excel(writer, sheet_name="ThresholdTimes", index=False)




    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


chance = random.random()
if chance < .1:
    print("""\n
          \n
    For God so loved the world, \n
      that he gave his only begotten Son, \n
      that whosoever believeth in him \n
      should not perish, \n
      but have everlasting life.
      \n - John 3:16""")