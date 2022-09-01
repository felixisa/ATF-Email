import xlrd
import sqlite3

def readxlemails(): 
    conn = sqlite3.connect('ATFplacement.db')
    cur = conn.cursor()
    cur.execute(''' DROP TABLE IF EXISTS Emails''')

    SQL = '''CREATE TABLE Emails (Name TEXT, Email TEXT)'''
    cur.execute(SQL)
    
    loc = ("email_list.xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for i in range(sheet.nrows):
        n = sheet.cell_value(i, 0)
        e = sheet.cell_value(i, 1)
        cur.execute('''INSERT INTO Emails (Name, Email) VALUES(?, ?)''',(n, e))
    conn.commit()
    conn.close()

def readxlclasses(): 
    conn = sqlite3.connect('ATFplacement.db')
    cur = conn.cursor()
    cur.execute(''' DROP TABLE IF EXISTS Classes''')

    SQL = '''CREATE TABLE Classes (ClassName TEXT, Members TEXT)'''
    cur.execute(SQL)
    
    loc = ("do it like this.xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for i in range(sheet.nrows):
        n = sheet.cell_value(i, 0)
        e = sheet.cell_value(i, 1)
        cur.execute('''INSERT INTO Classes (ClassName, Members) VALUES(?, ?)''',(n, e))
    conn.commit()
    conn.close()

def main():
    readxlemails()
    readxlclasses()

main()
