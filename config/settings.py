template_sheets = ['Problem', 'Detail', 'Summary', 'Mismatch']
summary_loc = {"Summary": {"row": 10, "col": 2}}
overbilled_loc = {"Mismatch": {"row": 4, "col": 2}}
problem_loc = {
    "Empty": {"row": 4, "col": 2},
    "Format": {"row": 4, "col": 5},
    "Military": {"row": 4, "col": 8},
    "ConflictingTime": {"row": 4, "col": 11},
    "Acceptable": {"row": 4, "col": 14}
}

detail_loc = {
    "Empty": {"row": 4, "col": 2},
    "Format": {"row": 4, "col": 8},
    "Military": {"row": 4, "col": 14},
    "ConflictingTime": {"row": 4, "col": 20},
    "Acceptable": {"row": 4, "col": 26}
}

blacklist_charge_codes = ['(PEN00000.00)']
