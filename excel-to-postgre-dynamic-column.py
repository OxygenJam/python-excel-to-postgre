# This automatically inserts data from an excel
# sheet to a specified Postgre DB table
# made by the table-generator.py script
# -Zird Triztan E. Driz

import sys
import datetime
import re
import os

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

    # Check if string
    if(not isinstance(sheet_val, str)):

        # Check if float
        if((sheet_val % 1) == 0.0):
            return int(sheet_val)
        else:
            return float(sheet_val)
    else:
            new_val = str(sheet_val)

            # Check for apostrophes remove quotations
            new_val = new_val.replace("'","''")
            new_val = new_val.replace('"',"")
            return str(new_val)


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

    logPrint("Loading the worksheet...\n")
    book = xlrd.open_workbook("../Data/" + srcFileName)
    sheet = book.sheet_by_name(srcSheetName)
    logPrint("Finished loading the worksheet!")

    # Calls the psycopg factory function to return a Connection class instance
    try:

        db = psycopg2.connect (database = os.getenv("db"), user=os.getenv("user"), password=os.getenv("password"),host=os.getenv("host"),port=os.getenv("port"))
        cursor = db.cursor()
    except Exception as e:

        lineBreak()
        errPrint(str(e) +"\n")
        lineBreak()
        
        return -1

    logPrint("Connected to the ")
    bPrint(f"{os.getenv('db')} database")
    lineBreak()
    logPrint("Preparing query...\n")
    
    # Truncating of table
    query = "TRUNCATE {}".format(tableName)

    try:

        cursor.execute(query)
    except Exception as e:

        lineBreak()
        errPrint(str(e) +"\n")
        lineBreak()
        
        return -1

    # Query pre-processing

    # Take note table name must be same as sheetname
    query = "INSERT INTO {}(".format(tableName)

    # Retrieval of xlsx column header as db column header
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

        query += "{} ,".format(column) 

        c += 1

    # Preparation for string interpolation
    c=0
    query += "date_file_uploaded, file_name) VALUES {}"

    logPrint("Finished preparing query\n")
    lineBreak()
    print("QUERY STRING: \n\n")
    print(colored(query,'green'))
    


    with db.cursor() as cursor:

        for r in range(1,sheet.nrows):
            arr = []

            # Formatting info is not implemented in .xlsx
            #print(f"Row # {r} is {sheet.rowinfo_map[]}")

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
            # Format any NULL values
            formatted_query = formatted_query.replace("'NULL'","NULL")
            formatted_query = formatted_query.replace('"',"'")

            try:

                cursor.execute(formatted_query)
            except Exception as e:

                lineBreak()
                errPrint(str(e) + "\n")
                lineBreak()
                print("Error on excel data row {}\n".format(r+1))
                print("Tuple data: \n {}\n".format(values))
                print("Query data: \n {}\n".format(formatted_query))

                end = datetime.datetime.now()

                
                elapsed = end - start
                elapsed_mileseconds = elapsed.microseconds//1000
                elapsed = list(divmod(elapsed.days * 86400 + elapsed.seconds, 60))
                elapsed.append(elapsed_mileseconds)

                logPrint(f"Time elapsed: {elapsed[0]} minutes, {elapsed[1]} seconds, and {elapsed[2]} milleseconds")
                lineBreak()
                            
                return -1

    try:   

        db.commit()
        db.close()
    except Exception as e:

        lineBreak()
        errPrint(str(e) +"\n")
        lineBreak()
        
        return -1

    lineBreak()

    logPrint(str(sheet.nrows) + " number of rows / or data inserted to table\n")
    logPrint(str(sheet.ncols) + " number of columns / number of data per row\n\n")

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