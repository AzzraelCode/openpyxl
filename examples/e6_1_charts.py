import openpyxl
from openpyxl.chart import Reference, BarChart, LineChart


def example(filename="test.xlsx"):
    """
    Добавление графиков в Excel с пакетом OpenPyXl
    https://openpyxl.readthedocs.io/en/stable/charts/introduction.html
    Этот код для видео https://youtu.be/WQqxA8R8YaQ
    :return:
    """
    book = openpyxl.load_workbook(filename)
    sheet = book.active

    chart = LineChart()

    chart.anchor="J5"
    chart.width=15 # in cm
    chart.height=5 # in cm

    data = Reference(sheet, min_col=3, max_col=4, min_row=2, max_row=99)
    chart.add_data(data)

    sheet.add_chart(chart)
    book.save(filename)