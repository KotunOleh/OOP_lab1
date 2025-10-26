from PySide6.QtWidgets import QToolBar, QStyle, QMenu, QPushButton
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt

class UIManager:
    def __init__(self, main_window):
        self.window = main_window
        self.actions = {}

    def setup_toolbar(self):
        toolbar = QToolBar("Головна панель")
        toolbar.setMovable(False)
        self.window.addToolBar(toolbar)
        
        style = self.window.style()
        
        #Actions
        self._add_action(toolbar, "new", "Створити", "Створити новий файл (Ctrl + N)", "Ctrl+N", 
                         style.standardIcon(QStyle.StandardPixmap.SP_FileIcon), self.window.new_file)
        self._add_action(toolbar, "open", "Відкрити", "Відкрити файл (Ctrl + O)", "Ctrl+O", 
                         style.standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton), self.window.open_file)
        self._add_action(toolbar, "save", "Зберегти", "Зберегти файл (Ctrl + S)", "Ctrl+S", 
                         style.standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton), self.window.save_file, enabled=False)

        toolbar.addSeparator()

        act = self._add_action(toolbar, "show_formulas", "Показати формули", "Режим перегляду формул (Ctrl + F)", "Ctrl+F", 
                             style.standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView), self.window.toggle_formula_view, 
                             enabled=False, checkable=True)
        act.toggled.connect(self.window.toggle_formula_view) 


        toolbar.addSeparator()

       #Gooogle
        icon_google = style.standardIcon(QStyle.StandardPixmap.SP_DriveNetIcon)
        self._add_action(None, "google_login", "Увійти в Google", trigger_slot=self.window.authenticate_google)
        self._add_action(None, "google_logout", "Вийти з Google", trigger_slot=self.window.logout_google, enabled=False)
        self._add_action(None, "select_from_drive", "Обрати файл з Google Drive", trigger_slot=self.window.select_from_drive, enabled=False)
        self._add_action(None, "save_to_drive", "Зберегти на Google Drive", trigger_slot=self.window.save_to_drive, enabled=False)

        google_button = QPushButton(icon_google, " Google Drive", self.window)
        google_button.setToolTip("Інтеграція з Google Drive")
        google_menu = QMenu(self.window)
        google_menu.addAction(self.actions["google_login"])
        google_menu.addAction(self.actions["google_logout"])
        google_menu.addSeparator()
        google_menu.addAction(self.actions["select_from_drive"])
        google_menu.addAction(self.actions["save_to_drive"])
        google_button.setMenu(google_menu)
        toolbar.addWidget(google_button)
        self.window.google_button = google_button

    def setup_context_menus(self):
        # SHEET context menu
        self._add_action(None, "del_sheet", "Видалити аркуш", trigger_slot=self.window.delete_current_sheet_action)
        self._add_action(None, "ren_sheet", "Перейменувати аркуш", trigger_slot=self.window.rename_current_sheet_action)

        # TABLE context menu
        self._add_action(None, "add_row", "Додати рядок", trigger_slot=self.window.add_row)
        self._add_action(None, "del_row", "Видалити поточний рядок", trigger_slot=self.window.delete_row)
        self._add_action(None, "add_col", "Додати стовпець", trigger_slot=self.window.add_column)
        self._add_action(None, "del_col", "Видалити поточний стовпець", trigger_slot=self.window.delete_column)

    def _add_action(self, parent_widget, name, text, tooltip=None, shortcut=None, 
                    icon=None, trigger_slot=None, enabled=True, checkable=False) -> QAction:
        action = QAction(text, self.window)
        if icon:
            action.setIcon(icon)
        if tooltip:
            action.setToolTip(tooltip)
        if shortcut:
            action.setShortcut(shortcut)
        if trigger_slot and not checkable:
            action.triggered.connect(trigger_slot)
        
        action.setEnabled(enabled)
        action.setCheckable(checkable)
        
        self.actions[name] = action
        if parent_widget:
            parent_widget.addAction(action)
        return action

    def get_action(self, name: str) -> QAction | None:
        return self.actions.get(name)

    def set_action_enabled(self, name: str, enabled: bool) -> None:
        action = self.get_action(name)
        if action:
            action.setEnabled(enabled)
            
    def set_action_checked(self, name: str, checked: bool) -> None:
         action = self.get_action(name)
         if action and action.isCheckable():
             action.setChecked(checked)