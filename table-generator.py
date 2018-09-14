#This automatically generate a table to a Postgre
#table from an Excel file as a data source.
#This also appends two extra columns at the 
#end containing the date of file uploaded
#and the name of the file uploaded
# -Zird Triztan E. Driz

import xlrd
import psycopg2

import sys
import re

#Imports the necessary packages for loading local .env files
import os
from dotenv import load_dotenv
load_dotenv()

#Checks that syntax or arguments have been met
def isSyntaxCorrect(arr):

    if (len(arr) < 2):
        return -1

    return 0

#Checks the data type
def isIntOrFloat(sheet_val):

    #check if string
    if(not isinstance(sheet_val, str)):

        #check if float
        if((sheet_val % 1) == 0.0):
            return "INT"
        else:
            return "NUMERIC"
    else:
            return "VARCHAR"

def main():

    if(isSyntaxCorrect(sys.argv[1:])==-1):
        print("ERROR!\n")
        print("Please indicate the <filename> and <sheetname> in your argument\n")

        print("e.g.\n")
        print("python table-generator.py some-file some-sheet")
    
    srcFileName = sys.argv[1] + ".xlsx"
    srcSheetName = sys.argv[2]
    tableName = sys.argv[3]

    book = xlrd.open_workbook("../Data/" + srcFileName)
    sheet = book.sheet_by_name(srcSheetName)

    query = "CREATE TABLE {}(".format(tableName)

    db = psycopg2.connect (database = os.getenv("db"), user=os.getenv("user"), password=os.getenv("password"),host=os.getenv("host"),port=os.getenv("port"))

    cursor = db.cursor()

    c = 0
    c_length = sheet.ncols
    while(c<c_length):
        print(f"Column # {c+1} : {sheet.cell(0,c).value}")

        #Special character cleansing
        column = str(sheet.cell(0,c).value)

        #Replaces space, forward slash, parantheses, 
        #and dash with a single underscore
        #also removes periods and commas
        column = re.sub(r"[/() -]","_",column)
        column = re.sub(r"[.,]","",column)
        column = re.sub(r"_(\w+)\1+","_",column)
        

        column = column.replace("%","pct")
        column = column.replace("#","nbr")
        column = column.replace("&","and")

        column = str.lower(column)

        if(sheet.cell_type(1,c)!= 3):
            columntype = isIntOrFloat(sheet.cell(1,c).value)
        else:
            columntype = "VARCHAR"

        query += "{} {},".format(column, columntype) 

        c += 1

    query += "{} VARCHAR, {} VARCHAR)".format("date_file_uploaded", "file_name")
    print(query)

    cursor.execute(query)

    cursor.close()

    db.commit()

    db.close()

if __name__ == '__main__':
    print(sys.argv)
    main()
