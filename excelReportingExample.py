

import xlrd # read files in python script
import os


# current working folder
folder = os.path.dirname(os.path.abspath(__file__))

excelSheet1 = os.path.join(folder, 'testSheet1.xlsx')
excelSheet2 = os.path.join(folder, 'testSheet2.xlsx')

print(excelSheet1)

def excelFun(excelSheet1, excelSheet2):
    ''' takes in 2 excel files as strings and does things '''

    # open the workbooks
    b1 = xlrd.open_workbook(excelSheet1)
    b2 = xlrd.open_workbook(excelSheet2)

    # open the sheets in the workbooks
    sht1 = b1.sheet_by_index(0)
    sht2 = b2.sheet_by_index(0)

    # pull out information about a specific row
    a = sht1.row_values(0)
    print(a)

    # the number of rows
    nRow1 = sht1.nrows
    nRow2 = sht2.nrows
    print(f'number of rows in testSheet1 is {nRow1} and in testSheet2 is {nRow2}')

    # loop through the rows doing some process
    sumRow = 0
    for row in range(1, nRow1):        
        sumRow = sumRow + sht1.cell(row,2).value
        # print('the value is ',  round(sht1.cell(row,2).value, 2))
    print('the total (3dp) is', round(sumRow, 2))
    print('the average (3dp) is', round(sumRow / nRow1, 3))

print('done')


def loopThrough(wksht, n):
    pass

if __name__ == "__main__":
    excelFun(excelSheet1, excelSheet2)
    pass