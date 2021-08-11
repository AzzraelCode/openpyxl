import openpyxl
from openpyxl.chart import Reference, BarChart, LineChart


def example(filename="test.xlsx"):
    """
    https://openpyxl.readthedocs.io/en/stable/charts/introduction.html
    :return:
    """
    book = openpyxl.load_workbook(filename)
    sheet = book.active

    chart = LineChart()
    # chart = BarChart()

    # chart.add_data("Коллеги!C2:C99")

    # data = Reference(sheet, min_col=3, max_col=4, min_row=2, max_row=99)
    # print(data)
    # data = Reference(sheet, range_string="Коллеги!C2:C99")
    # print([data, data.min_col, data.max_col, data.min_row, data.max_row])
    # chart.add_data(data)

    # chart.anchor="J5"
    # chart.width=15 # in cm
    # chart.height=5 # in cm

    # By default the top-left corner of a chart is anchored to cell E15 and the size is 15 x 7.5 cm (approximately 5 columns by 14 rows)
    # активной таблицы
    sheet.add_chart(chart)

    # см. class ChartBase
    # sheet.add_chart(chart, "D10")

    book.save(filename)