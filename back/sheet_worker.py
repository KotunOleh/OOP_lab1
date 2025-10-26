from PySide6.QtWidgets import (QTabWidget, QTableWidget, QTableWidgetItem, 
                               QHeaderView, QMessageBox, QInputDialog, QMenu)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QPoint
from openpyxl.utils import get_column_letter

from utils.config import DEFAULT_ROWS, DEFAULT_COLS

class SheetManager:
    def __init__(self, tab_widget: QTabWidget, main_window):
        self.tab_widget = tab_widget
        self.main_window = main_window
        self._setup_signals()

    def _setup_signals(self):
        self.tab_widget.currentChanged.connect(self.main_window.on_tab_changed)
        self.tab_widget.tabBar().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tab_widget.tabBar().customContextMenuRequested.connect(self.main_window.show_tab_context_menu)
        
    def get_current_table(self) -> QTableWidget | None:
        return self.tab_widget.currentWidget()

    def get_current_sheet_name(self) -> str | None:
        idx = self.tab_widget.currentIndex()
        if idx != -1:
            return self.tab_widget.tabText(idx)
        return None

    def clear_tabs(self):
        self.tab_widget.blockSignals(True)
        self.tab_widget.clear()
        self.tab_widget.blockSignals(False)

    def add_sheet_tab(self, sheet_name: str, sheet_data = None) -> QTableWidget:
        table_widget = self.create_new_table_widget()
        self.populate_table(table_widget, sheet_data) # sheet_data може бути аркушем openpyxl
        index = self.tab_widget.addTab(table_widget, sheet_name)
        self.tab_widget.setCurrentIndex(index)
        self.update_column_headers(table_widget)
        return table_widget
        
    def create_new_table_widget(self) -> QTableWidget:
        table_widget = QTableWidget()
        table_widget.setToolTip("Клацніть правою кнопкою миші для опцій")
        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        
        table_widget.customContextMenuRequested.connect(self.main_window.show_context_menu)
        table_widget.itemChanged.connect(self.main_window.on_item_changed)
        table_widget.itemDoubleClicked.connect(self.main_window.on_item_double_clicked)
        return table_widget

    def populate_table(self, table_widget: QTableWidget, sheet) -> None:
        table_widget.blockSignals(True)
        
        max_row, max_col = DEFAULT_ROWS, DEFAULT_COLS
        if sheet:
            max_row = sheet.max_row
            max_col = sheet.max_column
            if max_row == 1 and max_col == 1 and sheet.cell(1, 1).value is None:
                 max_row, max_col = DEFAULT_ROWS, DEFAULT_COLS

        table_widget.setRowCount(max_row)
        table_widget.setColumnCount(max_col)
        self.update_column_headers(table_widget)

        if sheet:
            for row_idx, row in enumerate(sheet.iter_rows(max_row=max_row, max_col=max_col)):
                for col_idx, cell in enumerate(row):
                    cell_value = cell.value
                    cell_value_str = str(cell_value) if cell_value is not None else ""
                    item = QTableWidgetItem()
                    if cell_value_str.startswith("="):
                        item.setData(Qt.ItemDataRole.UserRole, cell_value_str)
                        item.setText("...") 
                    else:
                        item.setText(cell_value_str)
                    table_widget.setItem(row_idx, col_idx, item)
        else:
             for r in range(max_row):
                 for c in range(max_col):
                     table_widget.setItem(r, c, QTableWidgetItem(""))

        table_widget.blockSignals(False)
        
    def update_column_headers(self, table_widget: QTableWidget | None = None) -> None:
        if table_widget is None:
            table_widget = self.get_current_table()
        if not table_widget: return
        col_count = table_widget.columnCount()
        headers = [get_column_letter(i + 1) for i in range(col_count)]
        table_widget.setHorizontalHeaderLabels(headers)

    def add_row(self) -> None:
        table_widget = self.get_current_table()
        if not table_widget: return
        current_row = table_widget.currentRow()
        if current_row < 0: current_row = table_widget.rowCount()
        table_widget.insertRow(current_row)
        self.main_window.set_dirty(True)

    def delete_row(self) -> None:
        table_widget = self.get_current_table()
        if not table_widget: return
        current_row = table_widget.currentRow()
        if current_row >= 0: 
            table_widget.removeRow(current_row)
            self.main_window.set_dirty(True)

    def add_column(self) -> None:
        table_widget = self.get_current_table()
        if not table_widget: return
        current_col = table_widget.currentColumn()
        if current_col < 0: current_col = table_widget.columnCount()
        table_widget.insertColumn(current_col)
        self.update_column_headers(table_widget)
        self.main_window.set_dirty(True)

    def delete_column(self) -> None:
        table_widget = self.get_current_table()
        if not table_widget: return
        current_col = table_widget.currentColumn()
        if current_col >= 0:
            table_widget.removeColumn(current_col)
            self.update_column_headers(table_widget)
            self.main_window.set_dirty(True)
            
    def update_sheet_from_table(self, sheet, table_widget):
        if sheet.max_row > 0:
            sheet.delete_rows(1, sheet.max_row)
        
        for r in range(table_widget.rowCount()):
            for c in range(table_widget.columnCount()):
                item = table_widget.item(r, c)
                value_to_save = None 
                if item:
                    formula = item.data(Qt.ItemDataRole.UserRole)
                    if formula:
                        value_to_save = formula
                    else:
                        value_to_save = item.text()
                
                if value_to_save:
                    sheet.cell(row=r + 1, column=c + 1, value=value_to_save)

    def update_workbook_from_all_tabs(self, workbook):
        if not workbook: return
        for idx in range(self.tab_widget.count()):
            sheet_name = self.tab_widget.tabText(idx)
            if sheet_name not in workbook.sheetnames:
                continue 
            sheet = workbook[sheet_name]
            table_widget = self.tab_widget.widget(idx)
            self.update_sheet_from_table(sheet, table_widget)