# formulas.py

def percent_change(total_value_cell: str, total_cost_cell: str) -> str:
    return f"=IF({total_cost_cell}=0, 0, ({total_value_cell} - {total_cost_cell}) / {total_cost_cell})"