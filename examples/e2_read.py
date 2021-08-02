import openpyxl


def example():
    """
    Один из способов чтения значений
    :return:
    """
    book = openpyxl.load_workbook(filename="test.xlsx")
    # sheet = book.active
    # sheet = book.worksheets[1]
    sheet = book["Коллеги"]

    # for row in sheet.values:
    #     for cell in row:
    #         print(cell)

    for row in sheet.iter_rows():
        for cell in row:
            print(cell.value)