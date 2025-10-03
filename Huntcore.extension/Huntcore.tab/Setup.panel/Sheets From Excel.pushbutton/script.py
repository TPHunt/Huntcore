# -*- coding: utf-8 -*-
"""
Create or Update Sheets from the Arup Structural Drawing List Builder Excel file.

!!! NOTE: Excel file to be updated to provide full drg number without prefix codes !!!

Reads headers from row 5 (B5-K5) and values from row 6 onward.
Creates new sheets or updates existing ones, populating parameters.
Lets user pick title block family for all new sheets.
Only prompts for existing sheets if Excel data differs from Revit parameters.
Shows column-by-column differences for each existing sheet needing update.

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from Autodesk.Revit.DB import *
from pyrevit import forms, revit
import xlrd
import re

# Headers we expect in Excel
REQUIRED_HEADERS = [
    "Sheet Number",
    "Sheet Name",
    "Admin_Document_Title_Line 1",
    "Admin_Document_Title_Line 2",
    "Admin_Document_Title_Line 3",
    "Admin_Document_Title_Line 4",
    "Admin_Document_Package",
    "Admin_Document_Volume or System_Code"
]

# --- Helpers ---
def safe_str(val):
    """Convert Excel cell value safely to a clean string."""
    if val is None:
        return ""
    try:
        text = str(val)
    except Exception:
        text = str(val).encode('utf-8', errors='ignore').decode('utf-8')
    # Remove characters Revit might not accept
    return re.sub(r'[^\x20-\x7E]', '', text).strip()

def get_titleblock_types(doc):
    """Return available titleblock FamilySymbols."""
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
    return list(collector)

def get_param_value(param):
    """Get a string value from a Revit parameter safely."""
    if not param:
        return ""
    if param.StorageType == StorageType.String:
        return safe_str(param.AsString())
    return safe_str(param.AsValueString())

def get_differences(sheet, excel_data):
    """Return dict of differences between Revit sheet and Excel data."""
    diffs = {}
    # Check Sheet Name
    if safe_str(sheet.Name) != excel_data["Sheet Name"]:
        diffs["Sheet Name"] = (safe_str(sheet.Name), excel_data["Sheet Name"])
    # Check parameters
    for key, val in excel_data.items():
        if key in ["Sheet Number", "Sheet Name"]:
            continue
        param = sheet.LookupParameter(key)
        revit_val = get_param_value(param)
        if revit_val != val:
            diffs[key] = (revit_val, val)
    return diffs

# --- Main ---
doc = revit.doc

# Ask user to pick Excel file
excel_path = forms.pick_file(file_ext='xlsx', multi_file=False, title="Select Excel File with Sheet Data")
if not excel_path:
    forms.alert("No Excel file selected.", exitscript=True)

# Ask user to pick a titleblock family type
titleblocks = get_titleblock_types(doc)
if not titleblocks:
    forms.alert("No title block families found in this project.", exitscript=True)

titleblock_lookup = {tb.FamilyName + " : " + tb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(): tb for tb in titleblocks}
titleblock_choice = forms.SelectFromList.show(titleblock_lookup.keys(), title="Select Title Block Type", multiselect=False)
if not titleblock_choice:
    forms.alert("No title block selected.", exitscript=True)

selected_titleblock = titleblock_lookup[titleblock_choice]

# Read Excel file
wb = xlrd.open_workbook(excel_path)
sheet = wb.sheet_by_index(0)

# Read header row (row 5 = index 4)
headers = [safe_str(sheet.cell_value(4, col)) for col in range(sheet.ncols)]
header_map = {h: i for i, h in enumerate(headers) if h in REQUIRED_HEADERS}

missing = [h for h in REQUIRED_HEADERS if h not in header_map]
if missing:
    forms.alert("Missing required headers: {}".format(", ".join(missing)), exitscript=True)

# Read sheet data (rows start from row 6 = index 5)
sheet_data = []
for row_idx in range(5, sheet.nrows):
    row = {}
    for h in REQUIRED_HEADERS:
        col_idx = header_map[h]
        row[h] = safe_str(sheet.cell_value(row_idx, col_idx))
    # Skip blank or TOTAL rows
    if not row["Sheet Number"] or row["Sheet Number"].strip().lower() == "total":
        continue
    sheet_data.append(row)

# Collect sheet numbers already existing in Revit
existing_sheets = {s.SheetNumber: s for s in FilteredElementCollector(doc).OfClass(ViewSheet)}

# Separate into new and existing
new_sheet_numbers = [d["Sheet Number"] for d in sheet_data if d["Sheet Number"] not in existing_sheets]

# Find existing sheets that differ from Excel
sheets_with_differences = {}
for d in sheet_data:
    num = d["Sheet Number"]
    if num in existing_sheets:
        diffs = get_differences(existing_sheets[num], d)
        if diffs:
            # Build preview string
            diff_lines = []
            for key, (rev_val, xl_val) in diffs.items():
                diff_lines.append("{}: '{}' → '{}'".format(key, rev_val, xl_val))
            preview = "[{}] {}".format(num, "; ".join(diff_lines))
            sheets_with_differences[num] = preview

# Ask user which new sheets to create
chosen_new = forms.SelectFromList.show(new_sheet_numbers, title="Select NEW Sheets to Create", multiselect=True) or []

# Ask user which existing sheets to update (only if differences exist)
if sheets_with_differences:
    chosen_preview = forms.SelectFromList.show(sheets_with_differences.values(), title="Select EXISTING Sheets to Update (with differences)", multiselect=True) or []
    # Map previews back to sheet numbers
    chosen_update = [num for num, prev in sheets_with_differences.items() if prev in chosen_preview]
else:
    chosen_update = []

# --- Tracking ---
created = []
updated = []
skipped = []
errors = []

# Start Revit transaction
t = Transaction(doc, "Create/Update Sheets from Excel")
t.Start()

for data in sheet_data:
    sheet_number = data["Sheet Number"]
    sheet_name = data["Sheet Name"]

    try:
        if sheet_number in existing_sheets:
            if sheet_number in chosen_update:
                vs = existing_sheets[sheet_number]
                # Update name
                if vs.Name != sheet_name:
                    vs.Name = sheet_name
                # Update parameters
                for key, val in data.items():
                    if key in ["Sheet Number", "Sheet Name"]:
                        continue
                    param = vs.LookupParameter(key)
                    if param and not param.IsReadOnly:
                        param.Set(val)
                updated.append(sheet_number)
            else:
                skipped.append(sheet_number)
        else:
            if sheet_number in chosen_new:
                vs = ViewSheet.Create(doc, selected_titleblock.Id)
                vs.SheetNumber = sheet_number
                vs.Name = sheet_name
                # Set parameters
                for key, val in data.items():
                    if key in ["Sheet Number", "Sheet Name"]:
                        continue
                    param = vs.LookupParameter(key)
                    if param and not param.IsReadOnly:
                        param.Set(val)
                created.append(sheet_number)
            else:
                skipped.append(sheet_number)
    except Exception as ex:
        errors.append("{} ({})".format(sheet_number, str(ex)))

t.Commit()

# --- Report back ---
msg = []
if created:
    msg.append("✅ Created ({}):\n{}".format(len(created), "\n".join(created)))
if updated:
    msg.append("🔄 Updated ({}):\n{}".format(len(updated), "\n".join(updated)))
if skipped:
    msg.append("⚠️ Skipped:\n" + "\n".join(skipped))
if errors:
    msg.append("❌ Errors:\n" + "\n".join(errors))

forms.alert("\n\n".join(msg) if msg else "No changes made.")
