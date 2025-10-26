import sys
import os
import io

from PySide6.QtWidgets import (QMainWindow, QMessageBox, QTableWidgetItem, 
                               QMenu, QTabWidget, QPushButton, QInputDialog)
from PySide6.QtGui import QCloseEvent
from PySide6.QtCore import Qt, QTimer, QPoint

from back.calculator import FormulaCalculator
from back.file_worker import FileWorker
from back.google_drive import GoogleDriveManager
from back.sheet_worker import SheetWorker
from ui.ui_dispatcher import UIRenderer
from utils.config import APP_NAME, DEFAULT_SHEET_NAME

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setGeometry(100, 100, 800, 600)

        self.current_workbook = None
        self.current_filepath = None
        self.is_dirty = False
        self.is_formula_view = False
        self.is_calculating = False

        #Back end managers
        self.calculator = FormulaCalculator()
        self.file_manager = FileWorker(self)
        self.google_manager = GoogleDriveManager(self)
        
        # UI
        self.tab_widget = QTabWidget()
        self.tab_widget.setDocumentMode(True)
        self.tab_widget.setTabsClosable(False)
        self.setCentralWidget(self.tab_widget)
        self.add_sheet_button = QPushButton("+Новий аркуш")
        self.add_sheet_button.setToolTip("Додати новий аркуш")
        self.add_sheet_button.setEnabled(False) 
        self.tab_widget.setCornerWidget(self.add_sheet_button, Qt.Corner.TopRightCorner)

        self.sheet_manager = SheetWorker(self.tab_widget, self)
        self.ui_manager = UIRenderer(self)

        # Init UI
        self.ui_manager.setup_toolbar()
        self.ui_manager.setup_context_menus()
        self.add_sheet_button.clicked.connect(self.add_new_sheet_action)

        self.reset_app()

    def set_dirty(self, dirty: bool) -> None:
        if self.is_dirty == dirty:
            return
        self.is_dirty = dirty
        title = self.windowTitle().replace(" - ", f" ({APP_NAME}) - ", 1)
        if dirty and not title.endswith("*"):
            self.setWindowTitle(title + "*")
        elif not dirty and title.endswith("*"):
            self.setWindowTitle(title[:-1])

    #Events handlers
    def on_tab_changed(self, index: int):
        self.recalculate_all_cells()
        
    def toggle_formula_view(self, checked: bool):
        self.is_formula_view = checked
        self.recalculate_all_cells()

    def on_item_double_clicked(self, item: QTableWidgetItem):
        if self.is_calculating: return
        formula = item.data(Qt.ItemDataRole.UserRole)
        if formula:
            self.is_calculating = True 
            item.setText(formula)    
            self.is_calculating = False

    def on_item_changed(self, item: QTableWidgetItem):
        if self.is_calculating: return 
        self.set_dirty(True)
        table_widget = item.tableWidget()
        if not table_widget: return

        user_text = item.text()
        self.is_calculating = True 
        if user_text.startswith("="):
            item.setData(Qt.ItemDataRole.UserRole, user_text)
            if self.is_formula_view:
                item.setText(user_text)
            else:
                result = self.calculator.parse_and_calculate(user_text, table_widget)
                item.setText(result)
        else:
            item.setData(Qt.ItemDataRole.UserRole, None)
            item.setText(user_text) 
        self.is_calculating = False 
        QTimer.singleShot(0, self.recalculate_all_cells)

    def show_context_menu(self, position: QPoint):
        self.sheet_manager.show_context_menu(position)

    def show_tab_context_menu(self, position: QPoint):
        self.sheet_manager.show_tab_context_menu(position)

    def closeEvent(self, event: QCloseEvent):
        if self.is_dirty:
            if not self.prompt_save_changes():
                event.ignore()
                return
        event.accept()

    def prompt_save_changes(self) -> bool:
        reply = QMessageBox.question(self, "Зберегти зміни?",
            "У вас є незбережені зміни. Бажаєте зберегти їх?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel)

        if reply == QMessageBox.StandardButton.Yes:
            return self.save_file()
        elif reply == QMessageBox.StandardButton.No:
            return True
        else:
            return False

    #Calculations
    def recalculate_all_cells(self):
        table_widget = self.sheet_manager.get_current_table()
        if not table_widget: return
        if self.is_calculating: return
            
        self.is_calculating = True
        try:
            for _ in range(2):
                for r in range(table_widget.rowCount()):
                    for c in range(table_widget.columnCount()):
                        item = table_widget.item(r, c)
                        if not item: continue
                        formula = item.data(Qt.ItemDataRole.UserRole)
                        if formula and formula.startswith("="):
                            if self.is_formula_view:
                                if item.text() != formula:
                                    item.setText(formula)
                            else:
                                result = self.calculator.parse_and_calculate(formula, table_widget)
                                if item.text() != result:
                                    item.setText(str(result))
        finally:
            self.is_calculating = False

    #Files managing
    def new_file(self):
        if self.is_dirty:
            if not self.prompt_save_changes(): return
                
        workbook = self.file_manager.create_new_workbook()
        if workbook:
            self.current_workbook = workbook
            self.current_filepath = None
            self.sheet_manager.clear_tabs()
            sheet_name = self.current_workbook.sheetnames[0]
            self.sheet_manager.add_sheet_tab(sheet_name) 
            self._update_ui_state(is_file_open=True)
            self.setWindowTitle(f"{APP_NAME} - Новий файл")
            self.set_dirty(False)

    def open_file(self):
        if self.is_dirty:
            if not self.prompt_save_changes(): return

        workbook, filepath = self.file_manager.open_local_workbook()
        if workbook and filepath:
            self.current_workbook = workbook
            self.current_filepath = filepath
            self.sheet_manager.populate_all_tabs(self.current_workbook)
            self._update_ui_state(is_file_open=True)
            self.setWindowTitle(f"{APP_NAME} - {filepath}")
            self.set_dirty(False)

    def save_file(self) -> bool:
        if not self.current_workbook: return False
        self.sheet_manager.update_workbook_from_all_tabs(self.current_workbook)
        saved, new_path = self.file_manager.save_local_workbook(self.current_workbook, self.current_filepath)
        if saved:
            self.current_filepath = new_path
            self.set_dirty(False)
            self.setWindowTitle(f"{APP_NAME} - {new_path}")
        return saved

    #Drive managing
    def authenticate_google(self):
         if self.google_manager.authenticate():
              self.ui_manager.set_action_enabled("google_login", False)
              self.ui_manager.set_action_enabled("google_logout", True)
              self.ui_manager.set_action_enabled("select_from_drive", True)
              if self.current_workbook:
                   self.ui_manager.set_action_enabled("save_to_drive", True)

    def logout_google(self):
        self.google_manager.logout()
        self.ui_manager.set_action_enabled("google_login", True)
        self.ui_manager.set_action_enabled("google_logout", False)
        self.ui_manager.set_action_enabled("select_from_drive", False)
        self.ui_manager.set_action_enabled("save_to_drive", False)
        
    def select_from_drive(self):
        if self.is_dirty:
            if not self.prompt_save_changes(): return
            
        files, error = self.google_manager.list_spreadsheets()
        if error:
            QMessageBox.critical(self, "Помилка Google Drive", error)
            return
        if not files:
            QMessageBox.information(self, "Google Drive", "На вашому диску не знайдено таблиць.")
            return

        file_map = {f['name']: (f['id'], f['mimeType']) for f in files}
        item_name, ok = QInputDialog.getItem(self, "Обрати файл", 
            "Оберіть файл з Google Drive:", file_map.keys(), 0, False)

        if ok and item_name:
            file_id, mime_type = file_map[item_name]
            buffer, error = self.google_manager.download_file(file_id, mime_type)
            if error:
                 QMessageBox.critical(self, "Помилка Google Drive", error)
                 return
            if buffer:
                 workbook = self.file_manager.load_workbook_from_buffer(buffer)
                 if workbook:
                    self.current_workbook = workbook
                    self.current_filepath = None
                    self.sheet_manager.populate_all_tabs(self.current_workbook)
                    self._update_ui_state(is_file_open=True)
                    self.setWindowTitle(f"{APP_NAME} - {item_name} (Google Drive)")
                    self.set_dirty(False)

    def save_to_drive(self):
        if not self.current_workbook: return
        
        file_name_suggestion = os.path.basename(self.current_filepath) if self.current_filepath else "Untitled.xlsx"
        file_name, ok = QInputDialog.getText(self, "Зберегти на Google Drive", 
            "Введіть ім'я файлу:", text=file_name_suggestion)

        if ok and file_name:
            if not file_name.endswith(".xlsx"): file_name += ".xlsx"
            self.sheet_manager.update_workbook_from_all_tabs(self.current_workbook)
            buffer = self.file_manager.save_workbook_to_buffer(self.current_workbook)
            if buffer:
                file_id, link, error = self.google_manager.upload_file(file_name, buffer)
                if error:
                     QMessageBox.critical(self, "Помилка Google Drive", error)
                else:
                     QMessageBox.information(self, "Успіх", 
                         f"Файл успішно збережено на Google Drive.\nПосилання: {link}")
                     self.set_dirty(False)

    # Sheets managing
    def add_new_sheet_action(self):
        self.sheet_manager.add_new_sheet_action()

    def delete_current_sheet_action(self):
         self.sheet_manager.delete_current_sheet_action()

    def rename_current_sheet_action(self):
         self.sheet_manager.rename_current_sheet_action()
         
    #Rows and cols managing
    def add_row(self): self.sheet_manager.add_row()
    def delete_row(self): self.sheet_manager.delete_row()
    def add_column(self): self.sheet_manager.add_column()
    def delete_column(self): self.sheet_manager.delete_column()

    #Reset
    def reset_app(self):
        self.current_workbook = None
        self.current_filepath = None
        self.sheet_manager.clear_tabs()
        self._update_ui_state(is_file_open=False)
        self.setWindowTitle(APP_NAME)
        self.set_dirty(False)
        
        is_logged_in = os.path.exists('token.json')
        if is_logged_in:
             try:
                 creds = Credentials.from_authorized_user_file('token.json', SCOPES)
                 if creds.valid or (creds.expired and creds.refresh_token):
                     if creds.expired: creds.refresh(Request())
                     self.google_manager.service = build('drive', 'v3', credentials=creds)
                 else:
                      is_logged_in = False
             except Exception:
                  is_logged_in = False
        
        self.ui_manager.set_action_enabled("google_login", not is_logged_in)
        self.ui_manager.set_action_enabled("google_logout", is_logged_in)
        self.ui_manager.set_action_enabled("select_from_drive", is_logged_in)


    def _update_ui_state(self, is_file_open: bool):
        self.ui_manager.set_action_enabled("save", is_file_open)
        self.ui_manager.set_action_enabled("show_formulas", is_file_open)
        self.add_sheet_button.setEnabled(is_file_open)
        
        self.sheet_manager.set_edit_actions_enabled(is_file_open)
        
        google_logged_in = self.google_manager.service is not None
        self.ui_manager.set_action_enabled("save_to_drive", is_file_open and google_logged_in)