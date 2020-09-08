from PyQt5.QtWidgets import (QWidget,QScrollArea, QTableWidget, QVBoxLayout,QTableWidgetItem,
                             QApplication, QTableView, QAbstractScrollArea, QSizePolicy,
                             QHeaderView, QLabel, QDesktopWidget, QMainWindow, QTabWidget,
                             QGridLayout, QCalendarWidget, QComboBox, QPushButton, QMessageBox,
                             QInputDialog)
from PyQt5 import QtCore
from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QTextEdit
from numpy import isnan
import pandas as pd
import sys


class ExcelTable(object):
    UNNAMED = 'Unnamed'
    ADDTABLE = 'Add Team +'

    def __init__(self, widget, calendar, excel_file_path, tab_widget,
                sheet_name='Sheet1'):
        self.main_widget = widget
        self.calendar = calendar
        self.excel_file_path = excel_file_path
        self.sheet_name = sheet_name
        self.tab = tab_widget
        self.template_df = None
        self.col_names = []
        self.index_names = []
        self.tables = []
        self.populate_dataframe()
        self.add_table_to_tab('Prudents')
        self.add_table_to_tab(self.ADDTABLE, True)
        self.tab.tabBarClicked.connect(self.create_new_tab)

    def get_table_by_index(self, index):
        if 0 <= index < len(self.tables):
            return self.tables[index]
        return None

    def populate_dataframe(self):
        self.template_df = pd.read_excel(self.excel_file_path,
                                         sheet_name=self.sheet_name,
                                         index_col=0)
        self.template_df = self.template_df.fillna("")
        self.col_names = [col for col in self.template_df.columns
                          if not col.startswith(self.UNNAMED)]
        self.index_names = self.template_df.index.fillna('')

    def get_table_widget(self):
        table = QTableWidget()
        table.verticalHeader().setVisible(False)
        table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        table.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tables.append(table)
        return table

    def populate_table(self, table):
        table.setColumnCount(len(self.col_names) + 1)
        table.setHorizontalHeaderLabels([""] + self.col_names)
        table.setVerticalHeaderLabels(self.index_names)
        table.setRowCount(len(self.template_df.index))
        for i in range(len(self.template_df.index)):
            t = QLabel()
            t.setText(str('' if str(self.template_df.index[i]) == 'nan'
                          else self.template_df.index[i]))
            table.setIndexWidget(table.model().index(i, 0), t)
            j_ori = 1
            for j in range(len(self.template_df.columns)):
                if self.template_df.columns[j].startswith(self.UNNAMED):
                    continue
                table.setItem(i, j_ori,
                              QTableWidgetItem(
                                  str(self.template_df.iloc[i, j])))
                j_ori += 1

    def add_table_to_tab(self, name, empty_table=False):
        table = self.get_table_widget()
        if not empty_table:
            self.populate_table(table)
        self.tab.addTab(table, name)

    def create_new_tab(self, table_index):
        if table_index != (len(self.tables) - 1):
            return
        table_name, ok = QInputDialog.getText(self.main_widget, 'Team Name',
                                              'Enter Team Name')
        if table_name == '' or ok is False:
            return
        table = self.tables[table_index]
        self.tab.setTabText(table_index, table_name)
        self.populate_table(table)
        self.add_table_to_tab(self.ADDTABLE, True)

    @staticmethod
    def get_suffix(day):
        if 11 <= day <= 20:
            return 'th'
        if day % 10 == 1:
            return 'st'
        if day % 10 == 2:
            return 'nd'
        if day % 10 == 3:
            return 'rd'
        return 'th'

    def show_date(self, date_selected):
        print(date_selected)
        date_selected = date_selected.addDays(-(date_selected.dayOfWeek() - 1))
        jj = 1
        for i, col in enumerate(self.template_df.columns):
            if col.startswith('Unnamed'):
                continue
            date_end = date_selected.addDays(6)
            start_day = date_selected.day()
            end_day = date_end.day()
            week_str = (date_selected.shortMonthName(date_selected.month())
                        + " " + str(start_day) + self.get_suffix(start_day)
                        + " - " + date_end.shortMonthName(date_end.month())
                        + " " + str(end_day) + self.get_suffix(end_day))
            for table in self.tables[:-1]:
                week_item = table.item(0, jj)
                week_item.setText(week_str)
            date_selected = date_selected.addDays(7)
            jj += 1
        self.calendar.setVisible(False)


def main_widget():
    win = QWidget()
    tab = QTabWidget()

    cal = QCalendarWidget()
    cal.setGridVisible(True)
    cal.move(20, 20)
    cal.setVisible(False)
    excel_path = r'F:\python excel code\support_roaster_format.xlsx'
    et = ExcelTable(win, cal, excel_path, tab)
    cal.clicked.connect(et.show_date)

    def open_calendar(args):
        cal.setVisible(True)
    set_data_button = QPushButton("Set First Week")
    set_data_button.setFixedSize(200, 50)
    set_data_button.clicked.connect(open_calendar)

    layout = QGridLayout()
    layout.addWidget(set_data_button, 0, 0)
    layout.addWidget(cal, 0, 1)
    layout.addWidget(tab, 1, 0, -1, 0)
    win.setLayout(layout)
    return win, cal, et


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win, cal, et = main_widget()
    cal.clicked.connect(et.show_date)
    win.show()
    app.exec_()

