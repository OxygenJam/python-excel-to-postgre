import xlrd
import psycopg2

import sys
import datetime
import re

#Imports the necessary packages for loading local .env files
import os
from dotenv import load_dotenv
load_dotenv()

def isSyntaxCorrect(arr):

    if (len(arr) < 2):
        return -1

    return 0

def isIntOrFloat(sheet_val):

    #check if string
    if(not isinstance(sheet_val, str)):

        #check if float
        if((sheet_val % 1) == 0.0):
            return int(sheet_val)
        else:
            return float(sheet_val)
    else:
            return str(sheet_val)


def main():

    print_break = "============================="

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

    #Calls the psycopg factory function to return a Connection class instance
    db = psycopg2.connect (database = os.getenv("db"), user=os.getenv("user"), password=os.getenv("password"),host=os.getenv("host"),port=os.getenv("port"))

    cursor = db.cursor()
    print(f"Connected to the {os.getenv('db')} database")
    print(f"\n{print_break}\n")
    print("Preparing query...")
    
    #Truncating of table
    query = "TRUNCATE {}".format(tableName)

    cursor.execute(query)

    #Query pre-processing

    #Take note table name must be same as sheetname
    query = "INSERT INTO {}(".format(tableName)

    #Retrieval of xlsx column header as db column header
    c = 0
    c_length = sheet.ncols
    while(c<c_length):
        print(f"Column # {c+1} : {sheet.cell(4,c).value}")

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

        query += "{} ,".format(column) 

        c += 1

    #Preparation for string interpolation
    c=0
    query += "date_file_uploaded, file_name) VALUES {}"


    print(query)
    print("Finished preparing query")
    print(f"\n{print_break}\n")


    with db.cursor() as cursor:

        for r in range(1,sheet.nrows):
            arr = []

            for c in range(0,sheet.ncols):
                val = sheet.cell(r,c).value

                if(sheet.cell_type(r,c) == 0):
                    
                    arr.append("NULL")
                else:

                    if(sheet.cell_type(r,c) != 3):

                        arr.append(isIntOrFloat(val))
                    else:
                        arr.append(str(xlrd.xldate.xldate_as_datetime(val, book.datemode)))
            
            arr.append(str(datetime.datetime.now().strftime("%Y-%m-%d")))
            arr.append(srcFileName)

            values = tuple(arr)

            formatted_query = query.format(values)

            #Format any NULL values
            formatted_query = formatted_query.replace("'NULL'","NULL")

            cursor.execute(formatted_query)
        


    db.commit()

    db.close()

    print(str(sheet.ncols))
    print(str(sheet.nrows))

if __name__ == '__main__':
    print(sys.argv)
    main()