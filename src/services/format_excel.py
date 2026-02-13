# format_excel.py

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

class PortfolioFormatter:
    def __init__(self, worksheet):
        self.ws = worksheet
    
    # Public Methods
    
    # Formatting that does NOT depend on data
    def static_format(self):
        self._headers()
        self._set_column_widths()
        self._freeze_panes()
        self._add_filters()
    
    # Formatting that does depend on data
    def dynamic_format(self, start_row: int, end_row: int):
        self._currency_columns(start_row, end_row)
        self._percent_columns(start_row, end_row)
        self._gain_loss_conditional_formatting(start_row, end_row)
    
    # Private Helpers
    
    def _headers(self):
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        for cell in self.ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
    
    def _set_column_widths(self):
        widths = {
            "A": 12,  # Ticker
            "B": 12,  # Shares
            "C": 14,  # Avg Cost
            "D": 14,  # Total Cost
            "E": 14,  # Current Price
            "F": 14,  # Total Value
            "G": 14,  # Gain/Loss
            "H": 12,  # % Change
        }
        
        for col, width in widths.items():
            self.ws.column_dimensions[col].width = width
    
    def _freeze_panes(self):
        self.ws.freeze_panes = "C1"
    
    def _add_filters(self):
        self.ws.auto_filter.ref = f"A1:H{self.ws.max_row}"
    
    def _currency_columns(self, start_row, end_row):
        currency_cols = ["C", "D", "E", "F", "G"]
        
        for col in currency_cols:
            for row in range(start_row, end_row +1):
                self.ws[f"{col}{row}"].number_format = "$#,##0.00"
    
    def _percent_columns(self, start_row, end_row):
        percent_cols = ["H"]
        
        for col in percent_cols: 
            for row in range(start_row, end_row + 1):
                self.ws[f"{col}{row}"].number_format = "0.00%"
    
    def _gain_loss_conditional_formatting(self, start_row, end_row):
        gain_loss_col = "G"
        gain_range = f"{gain_loss_col}{start_row}:{gain_loss_col}{end_row}"

        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        self.ws.conditional_formatting.add(
            gain_range,
            CellIsRule(operator="greaterThan", formula=["0"], fill=green_fill)
        )

        self.ws.conditional_formatting.add(
            gain_range,
            CellIsRule(operator="lessThan", formula=["0"], fill=red_fill)
        )