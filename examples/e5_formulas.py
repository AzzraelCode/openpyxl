import openpyxl

def example():
    """
    Оформление таблиц (колонок, строк, ячеек)
    :return:
    """
    filename="test.xlsx"
    book = openpyxl.load_workbook(filename)
    sheet = book.active

    sheet['E2'].value = '=SUM(C1:C999)'
    sheet['E3'].value = '=AVERAGE(C1:C999)'

    book.save(filename)