# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms

def toggle_case(sheet_name):
    if sheet_name.isupper():
        return sheet_name.capitalize()
    else:
        return sheet_name.upper()

def main():
    # Filter for Sheets only
    sheets = forms.select_sheets(title="Select Sheets to Toggle Name Case")
    if not sheets:
        forms.alert("No sheets selected.", title="Sheet Name Case Toggle")
        return

    with revit.Transaction("Toggle Sheet Name Case"):
        for sheet in sheets:
            name = sheet.Name
            new_name = toggle_case(name)
            try:
                sheet.Name = new_name
            except Exception as e:
                forms.alert("Failed to rename sheet '{}': {}".format(name, str(e)), title="Error")

    forms.alert("Sheet name case toggle complete.", title="Success")

if __name__ == '__main__':
    main()
