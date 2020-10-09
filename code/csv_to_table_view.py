import sys

import pandas as pd
from numpy import array
from PyQt5 import QtCore
from PyQt5.QtWidgets import (QWidget, QTableWidget, QTableWidgetItem,
                             QApplication, QAbstractScrollArea, QSizePolicy,
                             QHeaderView, QLabel, QTabWidget,
                             QGridLayout, QCalendarWidget, QPushButton,
                             QMessageBox,
                             QInputDialog, QFileDialog)
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import numbers
import re
from collections import defaultdict


class MappingExcelSheet(object):
    def __init__(self, sheet, df):
        self.excelformulas = ['SUM', 'COUNT', 'IF']
        self.min_cell = None
        self.min_df = None
        self.excel_sheet = sheet
        self.df = df
        self.set_map_excel()

    def set_map_excel(self):
        for i in range(1, self.excel_sheet.max_column):
            for j in range(1, self.excel_sheet.max_row):
                if self.excel_sheet.cell(i, j).value is None:
                    continue
                self.min_cell = [i, j]
                # self.min_df = self._find_minimum_df_cell(
                #     self.excel_sheet.cell(i, j).value)
                return

    def parse_formula(self, formula):
        formula = formula[1:]  # lose the equal sign
        if formula.contains('('):
            formula_func = formula.split("(")[0]  # Get sum or count or if
            if formula_func in self.excelformulas:
                within_func = formula.split("(")[1].split(')')[0]
                individual_cells = within_func.split(',')
        individual_cells = re.split('; |, |\+|\n|\*|-|/', formula)

    def convert_cells_to_df_locations(self, cells):
        pass


class ExcelTable(object):
    UNNAMED = 'Unnamed'
    ADDTABLE = 'Add Team +'

    def __init__(self, widget, calendar, excel_file_path, tab_widget,
                 summary_sheet='Summary', sheet_name='Sheet1'):
        self.main_widget = widget
        self.calendar = calendar
        self.tab = tab_widget
        self.excel_file_path = excel_file_path
        self._activated_table = None
        self.summary_sheet = summary_sheet
        self.sheet_name = sheet_name
        self.index_names = []
        self.tables = []
        self.sheet_mapper = {}
        self.map_names = defaultdict(list)
        self.wb = load_workbook(self.excel_file_path)
        self.summary_df = self.populate_dataframe(sheet_name=self.summary_sheet)
        self.template_df = self.populate_dataframe()
        self.add_table_to_tab(self.summary_sheet)
        self.add_table_to_tab(self.sheet_name)
        self.add_table_to_tab(self.ADDTABLE, True)
        # self.tab.currentChanged.connect(self.create_new_tab)
        self.tab.tabBarClicked.connect(self.create_new_tab)
        self.map_all_sheet_name()

    def get_table_by_index(self, index):
        if 0 <= index < len(self.tables):
            return self.tables[index]
        return None

    def populate_dataframe(self, sheet_name=None):
        if sheet_name is None:
            sheet_name = self.sheet_name
        sheet_index = self.wb.sheetnames.index(sheet_name)
        sheet = self.wb.worksheets[sheet_index]
        template_df = pd.read_excel(self.excel_file_path,
                                    sheet_name=sheet_name,
                                    index_col=None, header=None)
        template_df = template_df.fillna("")
        mp_obj = MappingExcelSheet(sheet, template_df)
        self.sheet_mapper[sheet_name] = mp_obj
        return template_df

    def get_table_widget(self):
        table = QTableWidget()
        table.verticalHeader().setVisible(False)
        table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        table.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.horizontalHeader().setDefaultAlignment(
            QtCore.Qt.AlignHCenter | QtCore.Qt.Alignment(QtCore.Qt.TextWordWrap))
        table.verticalHeader().setDefaultAlignment(
            QtCore.Qt.AlignHCenter | QtCore.Qt.Alignment(QtCore.Qt.TextWordWrap))
        self.tables.append(table)
        return table

    def populate_table(self, table, sheet_name):
        if sheet_name not in self.sheet_mapper:
            sheet = self.wb.copy_worksheet(self.wb[self.sheet_name])
            mp_obj = MappingExcelSheet(sheet, None)
            self.sheet_mapper[sheet_name] = mp_obj
        cur_sheet_map = self.sheet_mapper[sheet_name]
        sheet = cur_sheet_map.excel_sheet
        values = array([list(v) for v in sheet.values])
        min_i, min_j = cur_sheet_map.min_cell
        table.setColumnCount(len(values[0]) - min_j + 1)
        table.setHorizontalHeaderLabels(values[min_i - 1, min_j - 1:])
        table.setVerticalHeaderLabels(values[min_i - 1:, min_j - 1])
        table.setRowCount(len(values) - min_i + 1)
        for i in range(cur_sheet_map.min_cell[0], len(values)):
            t = QLabel()
            t.setAlignment(QtCore.Qt.AlignCenter)
            t.setWordWrap(True)
            t.setText(str('' if values[i, min_j - 1] is None else values[i, min_j - 1]))
            table.setIndexWidget(table.model().index(i - min_i, 0), t)
            for j in range(cur_sheet_map.min_cell[1], len(values[i])):
                text = str('' if values[i, j] is None else values[i, j])
                tw = QTableWidgetItem(text)
                table.setItem(i - min_i, j - min_j + 1, tw)
        table.cellChanged.connect(self.update_aggregation_df)
        table.cellClicked.connect(self.cell_activated)

    def map_all_sheet_name(self, sheet_names=None):
        if sheet_names is None:
            sheet_names = self.sheet_mapper.keys()
        for sheet_name in sheet_names:
            cur_sheet_map = self.sheet_mapper[sheet_name]
            values = array([list(v) for v in cur_sheet_map.excel_sheet.values])
            for i in range(cur_sheet_map.min_cell[0] - 1, len(values)):
                for j in range(cur_sheet_map.min_cell[1] - 1, len(values[i])):
                    value = '' if values[i, j] is None else str(values[i, j])
                    if sheet_name != self.summary_sheet:
                        if sheet_name in value:
                            self.map_names[sheet_name].append([sheet_name, i + 1, j + 1])
                        continue
                    for key in self.sheet_mapper.keys():
                        if key in value:
                            self.map_names[sheet_name].append([key, i + 1, j + 1])
        print(self.map_names)

    def add_table_to_tab(self, name, empty_table=False, df=None):
        table = self.get_table_widget()
        if not empty_table:
            self.populate_table(table, name)
            self.update_aggregation_df(None, [1, 2, 3, 4], -1)
        self.tab.addTab(table, name)

    def create_new_tab(self, table_index):
        if table_index != (len(self.tables) - 1) and table_index != self.tab.currentIndex():
            return
        table_name, ok = QInputDialog.getText(self.main_widget, 'Team Name',
                                              'Enter Team Name')
        if table_name == '' or ok is False:
            if table_index == (len(self.tables) - 1):
                self.tab.setCurrentIndex(0)
            return
        if table_index == self.tab.currentIndex():
            self.update_sheet_names(self.tab.tabText(table_index), table_name)
        if table_index != (len(self.tables) - 1):
            return
        self._activated_table = None
        table = self.tables[table_index]
        self.populate_table(table, sheet_name=table_name)
        self.update_aggregation_df(None, [1, 2, 3, 4], -1)
        self.tab.setTabText(table_index, table_name)
        self.add_table_to_tab(self.ADDTABLE, True)
        # self.add_to_summary(table_name)

    def update_sheet_names(self, old_name, new_name):
        new_map = defaultdict(list)
        for sheet_key in self.map_names:
            cur_sheet_map = self.sheet_mapper[sheet_key]
            sheet = cur_sheet_map.excel_sheet
            m_i, m_j = cur_sheet_map.min_cell
            map_name = new_name if sheet_key == old_name else old_name
            new_map[map_name] = []
            for key, i, j in self.map_names[sheet_key]:
                value = old_name
                if key == old_name:
                    table = [t for i, t in enumerate(self.tables)
                             if self.tab.tabText(i) == sheet_key][0]
                    if i == m_i:
                        t_value = table.horizontalHeaderItem(j - m_j).text().replace(old_name, new_name)
                        table.horizontalHeaderItem(j - m_j).setText(t_value)
                    elif j == m_j:
                        t_value = table.indexWidget(table.model().index(
                            i - m_i - 1, 0)).text().replace(old_name, new_name)
                        table.indexWidget(table.model().index(
                            i - m_i - 1, 0)).setText(t_value)
                    else:
                        t_value = table.item(i - m_i - 1, j - m_j).text().replace(old_name, new_name)
                        table.item(i - m_i - 1, j - m_j).setText(t_value)
                    value = new_name
                    s_value = sheet.cell(i, j).value.replace(old_name, new_name)
                    sheet.cell(i, j).value = s_value
                new_map[map_name].append([value, i, j])
        del self.map_names
        self.map_names = new_map
        old_sheet_mapper_value = self.sheet_mapper.pop(old_name)
        self.sheet_mapper[new_name] = old_sheet_mapper_value
        sheet_map = self.sheet_mapper[new_name]
        sheet_map.title = new_name

    def cell_activated(self, row, col):
        self._activated_table = self.tab.currentIndex()

    def update_aggregation_df(self, row, cols, table_index=None):
        return
        if table_index is None:
            table_index = (self._activated_table
                           if self._activated_table is not None
                           else self.tab.currentIndex())
        if not isinstance(cols, list):
            cols = [cols]
        table = self.tables[table_index]
        for col in cols:
            sumc = 0
            for irow in [2, 3]:
                tw = table.item(irow, col)
                if tw.text() == '':
                    tw.setText('0')
                sumc += int(tw.text())
            sum_row = table.item(4, col)
            sum_row.setText(str(sumc))
        self._activated_table = None

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
        stagnant_row = 2
        col_maps = {}
        for i, (sheet_name, cur_sheet_mapper) in enumerate(self.sheet_mapper.items()):
            if sheet_name == 'Summary':
                continue
            excel_sheet = cur_sheet_mapper.excel_sheet
            min_i, min_j = cur_sheet_mapper.min_cell
            for col in range(1, excel_sheet.max_column + 1):
                value = excel_sheet.cell(stagnant_row, col).value
                if not value:
                    continue
                if col not in col_maps:
                    date_end = date_selected.addDays(6)
                    start_day = date_selected.day()
                    end_day = date_end.day()
                    week_str = (date_selected.shortMonthName(date_selected.month())
                                + " " + str(start_day) + self.get_suffix(start_day)
                                + " - " + date_end.shortMonthName(date_end.month())
                                + " " + str(end_day) + self.get_suffix(end_day))
                    col_maps[col] = week_str
                week_str = col_maps.get(col)
                excel_sheet.cell(stagnant_row, col).value = week_str
                for table in self.tables[1:-1]:
                    week_item = table.item(stagnant_row - 1 - min_i, col - min_j)
                    week_item.setText(week_str)
                if col not in col_maps:
                    date_selected = date_selected.addDays(7)

        # for i, col in enumerate(range(7)):
        #     if col % 2 == 1:
        #         continue
        #     date_end = date_selected.addDays(6)
        #     start_day = date_selected.day()
        #     end_day = date_end.day()
        #     week_str = (date_selected.shortMonthName(date_selected.month())
        #                 + " " + str(start_day) + self.get_suffix(start_day)
        #                 + " - " + date_end.shortMonthName(date_end.month())
        #                 + " " + str(end_day) + self.get_suffix(end_day))
        #     for table in self.tables[1:-1]:
        #         week_item = table.item(0, col)
        #         if not week_item:
        #             continue
        #         week_item.setText(week_str)
        #
        #     date_selected = date_selected.addDays(7)
        #     jj += 1
        self.calendar.setVisible(False)

    def _get_dtyped_test(self, value):
        try:
            if float(value) == int(value):
                return int(value), numbers.FORMAT_NUMBER
            else:
                return float(value), numbers.FORMAT_NUMBER
        except Exception as ex:
            return value, None
        return value, None

    def clear_all(self):
        self.tab.clear()
        self.tables = []
        del self.template_df
        self.template_df = None
        self._activated_table = None
        self.excel_file_path = r''
        self.sheet_name = ''

    def import_from_excel(self, file_path):
        wb = load_workbook(file_path)
        response = QMessageBox.question(self.main_widget, 'Data Clear', "Are you sure to clear the existing data")
        if response == QMessageBox.Yes:
            self.clear_all()
            self.tab.currentChanged.disconnect(self.create_new_tab)
            self.excel_file_path = file_path
            for sheet in wb.sheetnames:
                df = self.populate_dataframe(sheet)
                if self.template_df is None:
                    self.template_df = df
                self.add_table_to_tab(sheet, df=df)
            self.add_table_to_tab(self.ADDTABLE, True)
            self.tab.currentChanged.connect(self.create_new_tab)

    def export_to_excel(self, file_path):
        if file_path == '':
            QMessageBox.Warning("Not a valid path")
        wb = load_workbook(self.excel_file_path)
        first_sheet_name = wb.sheetnames[0]
        sheet = wb[first_sheet_name]
        for tab_index in range(self.tab.count() - 1):
            title = self.tab.tabBar().tabText(tab_index)
            if tab_index != 0:
                sheet = wb.copy_worksheet(wb[wb.sheetnames[0]])
            sheet.title = title
            sheet.cell(1, 1).value = 'Support Activity - ' + str(title)
            table = self.tables[tab_index]
            for row in range(table.rowCount()):
                for col in range(1, table.columnCount()):
                    cell = sheet.cell(row + 2, col + col + 1)
                    if (row + 2) == 6:
                        formula = f'=SUM({cell.column_letter}4:{cell.column_letter}5)'
                        cell.value = formula
                        continue
                    txt, formt = self._get_dtyped_test(table.item(row, col).text())
                    cell.value = txt
                    if formt is not None:
                        cell.number_format = formt

        wb.save(file_path)


def main_widget():
    win = QWidget()
    tab = QTabWidget()

    cal = QCalendarWidget()
    cal.setGridVisible(True)
    cal.move(20, 20)
    cal.setVisible(False)
    excel_path = r'C:\Users\akhil\PycharmProjects\ExcelExtractor\support_roaster_format.xlsx'
    et = ExcelTable(win, cal, excel_path, tab, summary_sheet='Summary',
                    sheet_name='Prudents')
    cal.clicked.connect(et.show_date)

    def open_calendar(args):
        cal.setVisible(True)

    def file_save_as(args):
        filePath, _ = QFileDialog.getSaveFileName(win, "Export to Excel", "",
                                                  "Excel Workbook (*.xlsx);;Excel 97-2003 Workbook (*.xls);;CSV UTF-8 (Comma delimited) (*.csv)")
        et.export_to_excel(filePath)

    def file_open(args):
        filePath, _ = QFileDialog.getOpenFileName(win, "Import from Excel", "",
                                                  "Excel Workbook (*.xlsx);;Excel 97-2003 Workbook (*.xls);;CSV UTF-8 (Comma delimited) (*.csv)")
        et.import_from_excel(filePath)

    set_data_button = QPushButton("Set First Week")
    set_data_button.setFixedSize(200, 50)
    set_data_button.clicked.connect(open_calendar)

    load_existing = QPushButton("Load Existing")
    load_existing.setFixedSize(200, 50)
    load_existing.clicked.connect(file_open)

    save_data_button = QPushButton("Export to Excel")
    save_data_button.setFixedSize(200, 50)
    save_data_button.clicked.connect(file_save_as)

    layout = QGridLayout()
    layout.addWidget(set_data_button, 0, 0)
    layout.addWidget(cal, 0, 1)
    layout.addWidget(load_existing, 0, 2)
    layout.addWidget(tab, 1, 0, 1, 0)
    layout.addWidget(save_data_button, 2, 0)
    win.setLayout(layout)
    return win, cal, et


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win, cal, et = main_widget()
    cal.clicked.connect(et.show_date)
    win.show()
    app.exec_()
