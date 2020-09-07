from PyQt5.QtWidgets import (QWidget,QScrollArea, QTableWidget, QVBoxLayout,QTableWidgetItem,
                             QApplication, QTableView, QAbstractScrollArea, QSizePolicy,
                             QHeaderView, QLabel, QDesktopWidget, QMainWindow, QTabWidget,
                             QGridLayout, QCalendarWidget, QComboBox, QPushButton)
from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QTextEdit
from numpy import isnan
import pandas as pd
import sys


if __name__ == '__main__':

    app = QApplication(sys.argv)
    win = QWidget()
    tab = QTabWidget()

    def get_suffix(day):
        if day >= 11 and day <= 20:
            return 'th'
        if day % 10 == 1:
            return 'st'
        if day % 10 == 2:
            return 'nd'
        if day % 10 == 3:
            return 'rd'
        return 'th'

    def showDate(date_selected):
        print(date_selected)
        date_selected = date_selected.addDays(-(date_selected.dayOfWeek() - 1))
        jj = 1
        for i, col in enumerate(df.columns):
            if col.startswith('Unnamed'):
                continue
            date_end = date_selected.addDays(6)
            start_day = date_selected.day()
            end_day = date_end.day()
            print("##########################")
            week_str = (date_selected.shortMonthName(date_selected.month())
                        + " " + str(start_day) + get_suffix(start_day)
                        + " - " + date_end.shortMonthName(date_end.month())
                        + " " + str(end_day) + get_suffix(end_day))
            print(week_str)
            week_item = table.item(0, jj)
            week_item.setText(week_str)
            date_selected = date_selected.addDays(7)
            jj += 1


        cal.setVisible(False)


    cal = QCalendarWidget()
    cal.setGridVisible(True)
    cal.move(20, 20)
    cal.clicked.connect(showDate)
    cal.setVisible(False)

    def open_calendar(args):
        cal.setVisible(True)

    set_data_button = QPushButton("Set First Week")
    set_data_button.setFixedSize(200, 50)
    set_data_button.clicked.connect(open_calendar)


    layout = QGridLayout()
    table = QTableWidget()
    table.verticalHeader().setVisible(False)
    table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
    table.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
    tab.addTab(table, 'Prudents')
    layout.addWidget(set_data_button, 0, 0)
    layout.addWidget(cal, 0, 1)
    layout.addWidget(tab, 1, 0, -1, 0)
    win.setLayout(layout)

    # dfinput = pd.read_csv(r'F:\python excel code\support_roaster_format.xlsx', header=None)
    df = pd.read_excel(r'F:\python excel code\support_roaster_format.xlsx', sheet_name='Sheet1', index_col=0)
    df = df.fillna("")
    col_names = [col for col in df.columns if not col.startswith('Unnamed')]
    index_names = df.index.fillna('')
    table.setColumnCount(len(col_names) + 1)
    table.setHorizontalHeaderLabels([""] + col_names)
    table.setVerticalHeaderLabels(index_names)
    table.setRowCount(len(df.index))


    for i in range(len(df.index)):
        t = QLabel()
        t.setText(str('' if str(df.index[i]) == 'nan' else df.index[i]))
        table.setIndexWidget(table.model().index(i, 0), t)
        j_ori = 1
        for j in range(len(df.columns)):
            if df.columns[j].startswith('Unnamed'):
                continue
            table.setItem(i, j_ori, QTableWidgetItem(str(df.iloc[i, j])))
            j_ori += 1

    hh = table.item(0, 1)
    hh.setText('Heeellllo')
    win.show()
    app.exec_()