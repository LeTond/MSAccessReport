# import pyodbc
#
# msa_drivers = [x for x in pyodbc.drivers() if 'ACCESS' in x.upper()]
# print(f'MS-Access Drivers : {msa_drivers}')
#
# conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=/home/lg/Documents/DataAlmazDomen/Statistika/Base_2020.accdb;')
#
# cursor = conn.cursor()
# # cursor.execute('select * from Журнал WHERE Дата исследования > CONVERT(datetime, 2020, 9, 12, 0, 0)')
# cursor.execute('select * from Журнал WHERE  date between 2020-9-13 and 2020-9-30')
# # cursor.execute('select * from Журнал')
#
#
# for row in cursor.fetchall():
#     print(row)
#
#
# """
# pandas_access
# docx
# PyQt5
# pip install pandas
# pip install openpyxl xlsxwriter xlrd
# """
from pprint import pprint

if __name__ == '__main__':
    diction = {
        "amb": {
            "with contrast": {},
            "without contrast": {}
        },
        "inpatient": {
            "with contrast": {},
            "without contrast": {}
        }
    }

    diction["amb"].update({"with_contrast": {"oms": 10}})

    pprint(diction)

