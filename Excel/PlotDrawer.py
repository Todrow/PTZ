from openpyxl.chart import LineChart, BarChart, PieChart, Reference, ScatterChart


def add_chart_to_sheet(sheet, chart_type, chart_data_x=(1, None, 2, 7), chart_data_y=(2, 3, 1, 7), title="Title",
                       x_axis_title="X", y_axis_title="Y", cell="A10"):
    """
    Функция добавления диаграммы (графика) в файл Excel
    @param sheet: лист Excel, в который следует добавить диаграмму
    @param chart_type: тип диаграммы (например, 'bar', 'line', 'pie')
    @param chart_data_x: tuple(min_col, max_col, min_row, max_row), диапазон данных для оси X
    @param chart_data_y: tuple(min_col, max_col, min_row, max_row), диапазон данных для оси Y
    @param title: заголовок диаграммы
    @param x_axis_title: заголовок оси X
    @param y_axis_title: заголовок оси Y
    @param cell: имя ячейки, в которую следует добавить диаграмму
    @return: объект диаграммы openpyxl
    """
    # Создать диаграмму
    if chart_type == 'line':
        chart = LineChart()
    elif chart_type == 'bar':
        chart = BarChart()
        # установим тип - `вертикальные столбцы`
        chart.type = "col"
    elif chart_type == 'pie':
        chart = PieChart()
    else:
        chart = ScatterChart

    # установим стиль диаграммы (цветовая схема)
    chart.style = 10

    # Установить заголовок и заголовки осей
    chart.title = title
    chart.x_axis.title = x_axis_title
    chart.y_axis.title = y_axis_title

    # выберем 2 столбца с данными для оси `y`
    data = Reference(sheet, min_col=chart_data_y[0], max_col=chart_data_y[1],
                     min_row=chart_data_y[2], max_row=chart_data_y[3])
    # теперь выберем категорию для оси `x`
    categor = Reference(sheet, min_col=chart_data_x[0], max_col=chart_data_x[1],
                        min_row=chart_data_x[2], max_row=chart_data_x[3])

    # добавляем данные в объект диаграммы
    chart.add_data(data, titles_from_data=True)
    # установим метки на объект диаграммы
    chart.set_categories(categor)

    # Установить данные диаграммы
    chart.add_data(data, titles_from_data=True)
    # установим метки на объект диаграммы
    chart.set_categories(categor)

    # Добавить диаграмму в лист
    sheet.add_chart(chart, cell)

    return chart
