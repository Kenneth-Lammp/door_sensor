import openpyxl
import pyodbc
# Load the Excel workbook and select the active worksheet
workbook = openpyxl.load_workbook("/C::path/to/excel/file")
worksheet = workbook.active
# Connect to the SQL Server database
connection = pyodbc.connect('Driver={SQL Server};'
                            'Server=localhost;'
                            'Database=Expert-Only;'
                            'Trusted_Connection=yes;')
cursor = connection.cursor()
# Read data from the Excel file and insert it into the SQL Server table
for row in range(2, worksheet.max_row + 1):
    index = worksheet.cell(row=row, column=1).value
    time = worksheet.cell(row=row, column=2).value
    milliseconds = worksheet.cell(row=row, column=3).value
    people = worksheet.cell(row=row, column=4).value
    cursor.execute("INSERT INTO 'table name' (Index, Time, Milliseconds, People) VALUES (?, ?, ?, ?)",
                   index, time, milliseconds, people)
# Commit the changes and close the connection
connection.commit()
connection.close()
