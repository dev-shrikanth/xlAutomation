from openpyxl import load_workbook


def vlookup(lookup_value, lookup_range, return_column_index, worksheet):
    for row in worksheet[lookup_range]:
        if row[0].value == lookup_value:
            return row[return_column_index - 1].value
    return None
