import xlrd
import psycopg2

#Imports the necessary packages for loading local .env files
import os
from dotenv import load_dotenv
load_dotenv()
#.env file variables
#db
#user
#password
#host
#port

srcfilename = "Certified and trained List_MDA Program_30th June"
srcfilename +=".xlsx"
sheetname = "Certified List"

print_break = "============================="

book = xlrd.open_workbook("../Data/"+srcfilename)
sheet = book.sheet_by_name(sheetname)

#Calls the psycopg factory function to return a Connection class instance
db = psycopg2.connect (database = os.getenv("db"), user=os.getenv("user"), password=os.getenv("password"),host=os.getenv("host"),port=os.getenv("port"))

#The returns the cursor instance that handles all the db functions
cursor = db.cursor()
print(f"Connected to the {os.getenv('db')} database")

query = """INSERT INTO "Certified List" ("Learner e-mail", "Employee Id", "Geographic Area", "Geographic unit", "Country", "Career Level", "Enrollment Status", "Certification Name") VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""

#With is a clause in Python to dispose or close any existing
#filestream connection when the last line is met
with db.cursor() as cursor:

    for r in range(1,sheet.nrows):
        le = sheet.cell(r,0).value
        eid = sheet.cell(r,1).value
        ga = sheet.cell(r,2).value
        ge = sheet.cell(r,3).value
        c = sheet.cell(r,4).value
        cl = sheet.cell(r,5).value
        e = sheet.cell(r,6).value
        crt = sheet.cell(r,7).value

        values = (le,eid,ga,ge,c,cl,e,crt)
        cursor.execute(query, values)
    
#Commits all pending changes to the main db
db.commit()

#Closes the db connection
db.close()

print(str(sheet.ncols))
print(str(sheet.nrows))
print("I just imported Excel into postgreSQL")