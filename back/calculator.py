import re
from openpyxl.utils import column_index_from_string, get_column_letter

class FormulaCalculator:
    def __init__(self):
        self._cell_name_cache = {}

    @staticmethod
    def _excel_max(*args) -> float | int:
        if not args: return 0
        try: return max(args)
        except TypeError: return 0 

    @staticmethod
    def _excel_min(*args) -> float | int:
        if not args: return 0
        try: return min(args)
        except TypeError: return 0

    def cell_name_to_indices(self, cell_name: str) -> tuple[int, int] | None:
        if cell_name in self._cell_name_cache:
            return self._cell_name_cache[cell_name]
        match = re.match(r"([A-Z]+)([0-9]+)", cell_name.upper())
        if not match: return None
        col_str, row_str = match.groups()
        try:
            col_idx = column_index_from_string(col_str) - 1
            row_idx = int(row_str) - 1
            result = (row_idx, col_idx)
            self._cell_name_cache[cell_name] = result
            return result
        except Exception:
            return None

    def get_cell_value(self, cell_name: str, table_widget: any) -> float:
        indices = self.cell_name_to_indices(cell_name)
        if indices is None: return 0.0
        row, col = indices
        if row >= table_widget.rowCount() or col >= table_widget.columnCount():
            return 0.0
        item = table_widget.item(row, col)
        if not item or not item.text(): return 0.0
        try:
            return float(item.text())
        except ValueError:
            return 0.0


    def get_range_values(self, range_str: str, table_widget: any) -> list[float]:
        try:
            start_cell, end_cell = range_str.split(':')
            start_indices = self.cell_name_to_indices(start_cell)
            end_indices = self.cell_name_to_indices(end_cell)
            if start_indices is None or end_indices is None: return []
            r1, c1 = start_indices
            r2, c2 = end_indices
            min_row, max_row = min(r1, r2), max(r1, r2)
            min_col, max_col = min(c1, c2), max(c1, c2)
            values = []
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    cell_name = get_column_letter(c + 1) + str(r + 1)
                    values.append(self.get_cell_value(cell_name, table_widget))
            return [v for v in values if isinstance(v, (int, float))]
        except Exception:
            return []


    def parse_and_calculate(self, formula_string: str, table_widget: any) -> str:
        if not formula_string.startswith("="):
            return formula_string
        expression = formula_string[1:].upper()
        try:
            match = re.match(r"SUM\((.*?)\)", expression)
            if match:
                range_str = match.group(1)
                if ":" in range_str:
                    values = self.get_range_values(range_str, table_widget)
                    return str(sum(values))
            match = re.match(r"AVERAGE\((.*?)\)", expression)
            if match:
                range_str = match.group(1)
                if ":" in range_str:
                    values = self.get_range_values(range_str, table_widget)
                    if not values: return "0"
                    return str(sum(values) / len(values))
            match = re.match(r"MAX\((.*?)\)", expression)
            if match:
                range_str = match.group(1)
                if ":" in range_str:
                    values = self.get_range_values(range_str, table_widget)
                    if not values: return "0"
                    return str(max(values))
            match = re.match(r"MIN\((.*?)\)", expression)
            if match:
                range_str = match.group(1)
                if ":" in range_str:
                    values = self.get_range_values(range_str, table_widget)
                    if not values: return "0"
                    return str(min(values))

            cell_names = re.findall(r"([A-Z]+[0-9]+)", expression)
            eval_expression = expression
            for cell_name in set(cell_names):
                value = self.get_cell_value(cell_name, table_widget)
                eval_expression = re.sub(r"\b" + cell_name + r"\b", str(value), eval_expression)

            eval_expression = eval_expression.replace("^", "**")
            eval_expression = eval_expression.replace("=", "==")
            eval_expression = eval_expression.replace("!==", "!=")
            eval_expression = eval_expression.replace("<==", "<=")
            eval_expression = eval_expression.replace(">==", ">=")
            eval_expression = eval_expression.replace("<>", "!=")

            eval_globals = {}
            eval_locals = {
                "MAX": FormulaCalculator._excel_max,
                "MIN": FormulaCalculator._excel_min,
            }
            result = eval(eval_expression, eval_globals, eval_locals)
            return str(result)
        except Exception as e:
            return "#ERROR!"