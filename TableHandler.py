from PyQt5.QtWidgets import QTableWidgetItem

def draw_table(table, df):
    table.clear()

    table.setRowCount(df.shape[0])
    table.setColumnCount(df.shape[1])
    
    # adding each cell
    for row_ind, row in df.iterrows():
        for col_index, value in enumerate(row):
            if isinstance(value, (float, int)):
                value = "{0:0,.0f}".format(value)
            tableItem = QTableWidgetItem(str(value))
            table.setItem(row_ind, col_index, tableItem)


def clear_table(table):
    table.clear()
    table.setRowCount(0)
    table.setColumnCount(0)
