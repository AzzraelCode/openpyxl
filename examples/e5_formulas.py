import openpyxl

def example():
    """
    Оформление таблиц (колонок, строк, ячеек)
    :return:
    """
    filename="test.xlsx"
    book = openpyxl.load_workbook(filename)
    sheet = book["Коллеги"]
    book.active = sheet

    sheet['D1'].value = '=SUM(C1:C999)'
    sheet['D2'].value = '=AVERAGE(C1:C999)'

    book.save(filename)