# This automatically generate a table to a Postgre
# table from an Excel file as a data source.
# This also appends two extra columns at the 
# end containing the date of file uploaded
# and the name of the file uploaded
#
# Partner file to this script
# is the excel-to-postgre .py files
# -Zird Triztan E. Driz

import sys
import re
import os
import datetime

# Core package dependencies
import xlrd
import psycopg2

# Package for loading local .env files
from dotenv import load_dotenv

#Package for UI Color
from colorama import init
from termcolor import colored

# Use Colorama to make Termcolor work on Windows too
init()

# Loads the env file
load_dotenv()

# Colored printing
def logPrint(word):
    sys.stdout.write(colored(">> ","green"))
    sys.stdout.write(word)
    sys.stdout.flush()
    
def errPrint(word):
    sys.stdout.write(colored("ERR: ","red"))
    sys.stdout.flush()
    print(word)

def gPrint(word):
    print(colored(word,"green"))

def rPrint(word):
    print(colored(word,"red"))

def bPrint(word):
    print(colored(word,"blue"))

# Line break function
def lineBreak():
    print_break = "============================="
    print(colored(f"\n{print_break}\n", "green"))

# Checks that syntax or arguments have been met
def isSyntaxCorrect(arr):

    if (len(arr) < 2):
        return -1

    return 0

# Checks the data type
def isIntOrFloat(sheet_val):

    # check if string
    if(not isinstance(sheet_val, str)):

        return "NUMERIC"
    else:
            return "VARCHAR"

def main():
    start = datetime.datetime.now()

    if(isSyntaxCorrect(sys.argv[1:])==-1):
        
        errPrint("Please indicate the <filename> and <sheetname> in your argument\n")

        rPrint("e.g.\n")
        rPrint("python table-generator.py <some-file> <some-sheet> <table-name> <row-to-compare>")
        return -1
    
    srcFileName = sys.argv[1] + ".xlsx"
    srcSheetName = sys.argv[2]
    tableName = sys.argv[3]
    rowToCompare = int(sys.argv[4])

    logPrint("Loading the worksheet...\n")
    book = xlrd.open_workbook("../Data/" + srcFileName)
    sheet = book.sheet_by_name(srcSheetName)
    logPrint("Finished loading the worksheet!\n")

    query = "CREATE TABLE {}(".format(tableName)

    try:

        db = psycopg2.connect (database = os.getenv("db"), user=os.getenv("user"), password=os.getenv("password"),host=os.getenv("host"),port=os.getenv("port"))
        cursor = db.cursor()
    except Exception as e:

        lineBreak()
        errPrint(str(e) +"\n")
        lineBreak()
        
        return -1

    c = 0
    c_length = sheet.ncols
    while(c<c_length):
        logPrint(f"Column #  {c+1} :")
        bPrint(f"{sheet.cell(0,c).value}")

        # Special character cleansing
        column = str(sheet.cell(0,c).value)

        # Replaces space, forward slash, parantheses, 
        # and dash with a single underscore
        # also removes periods and commas
        column = re.sub(r"[/() -]","_",column)
        column = re.sub(r"[.,]","",column)
        column = re.sub(r"_(\w+)\1+","_",column)
        

        column = column.replace("%","pct")
        column = column.replace("#","nbr")
        column = column.replace("&","and")

        column = str.lower(column)

        if(sheet.cell_type(1,c)!= 3):
            columntype = isIntOrFloat(sheet.cell(rowToCompare,c).value)
        else:
            columntype = "VARCHAR"

        query += "{} {},".format(column, columntype) 

        c += 1

    query += "date_file_uploaded VARCHAR, file_name VARCHAR)"
    logPrint("Finished preparing query\n")
    lineBreak()
    print("QUERY STRING: \n\n")
    print(colored(query,'green'))

    try:

        cursor.execute(query)
    except Exception as e:

        lineBreak()
        errPrint(str(e) + "\n")

        end = datetime.datetime.now()


        elapsed = end - start
        elapsed_mileseconds = elapsed.microseconds//1000
        elapsed = list(divmod(elapsed.days * 86400 + elapsed.seconds, 60))
        elapsed.append(elapsed_mileseconds)

        logPrint(f"Time elapsed: {elapsed[0]} minutes, {elapsed[1]} seconds, and {elapsed[2]} milleseconds")
        lineBreak()

        return -1

    try:

        cursor.close()

        db.commit()
        db.close()
    except Exception as e:

        lineBreak()
        errPrint(str(e) +"\n")
        lineBreak()
        
        return -1

    lineBreak()

    end = datetime.datetime.now()

    elapsed = end - start
    elapsed_mileseconds = elapsed.microseconds//1000
    elapsed = list(divmod(elapsed.days * 86400 + elapsed.seconds, 60))
    elapsed.append(elapsed_mileseconds)

    logPrint(f"Time elapsed: {elapsed[0]} minutes, {elapsed[1]} seconds, and {elapsed[2]} milleseconds")
    lineBreak()

    return 0

if __name__ == '__main__':
    bPrint(str(sys.argv) + "\n\n")
    main()
