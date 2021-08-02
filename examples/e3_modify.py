import openpyxl
from openpyxl.worksheet import worksheet


def example():
    """
    Редактирую таблицу
    - двигаю данные
    - добавляю колонки и строки
    - записываю данные в ячейки
    https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html
    :return:
    """
    filename="test.xlsx"
    book = openpyxl.load_workbook(filename=filename)
    sheet : worksheet = book["Коллеги"]

    sheet.insert_rows(0)
    sheet["A1"].value = "Имя"
    sheet["B1"].value = "Телефон"
    sheet["C1"].value = "Деньги"

    # sheet.delete_rows(0)
    # sheet.move_range("A1:B200", cols=-22)

    # автофильтры и сортировка https://openpyxl.readthedocs.io/en/stable/filters.html
    sheet.auto_filter.ref = "A1:C999"
    # sheet.auto_filter.add_filter_column(2, ["0"]) # добавит условие, но не применит
    # sheet.auto_filter.add_sort_condition("B1:B999")

    book.save(filename)
