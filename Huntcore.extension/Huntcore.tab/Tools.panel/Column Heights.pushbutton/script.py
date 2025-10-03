# -*- coding: utf-8 -*-
"""
Adjust either the Base or Top Height of selected Structural Columns by a user-specified dimension.

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from pyrevit import forms, revit, DB
from pyrevit.forms import WPFWindow
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button

uidoc = revit.uidoc
doc   = revit.doc


# ------------------------------
# Collect user selection
# ------------------------------
sel_ids = uidoc.Selection.GetElementIds()
if not sel_ids:
    forms.alert("Please select one or more structural columns before running the tool.", exitscript=True)

columns = [doc.GetElement(eid) for eid in sel_ids if isinstance(doc.GetElement(eid), DB.FamilyInstance) 
           and doc.GetElement(eid).Category.Name == "Structural Columns"]

if not columns:
    forms.alert("No Structural Columns found in the selection.", exitscript=True)


# ------------------------------
# UI Layout
# ------------------------------
components = [
    Label("Adjust which end of the column?"),
    ComboBox("end_choice", {"Base": "BASE", "Top": "TOP"}),

    Separator(),
    Label("Enter adjustment amount (mm):"),
    TextBox("adjust_value", Text="300"),

    Separator(),
    Button("OK"),
    Button("Cancel")
]

form = FlexForm("Adjust Column Heights", components)
form.show()

if not form.values or "adjust_value" not in form.values:
    script.exit()


# ------------------------------
# Process Input
# ------------------------------
end_choice   = form.values["end_choice"]
adjust_mm    = float(form.values["adjust_value"])
adjust_ft    = adjust_mm / 304.8   # mm -> feet (Revit internal units)


# ------------------------------
# Transaction: Apply Adjustments
# ------------------------------
with revit.Transaction("Adjust Column Heights"):
    for col in columns:
        loc = col.Location
        if not isinstance(loc, DB.LocationPoint):
            continue

        base_param = col.get_Parameter(DB.BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
        top_param  = col.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)

        if end_choice == "BASE" and base_param:
            base_param.Set(base_param.AsDouble() + adjust_ft)

        elif end_choice == "TOP" and top_param:
            top_param.Set(top_param.AsDouble() + adjust_ft)


forms.alert("Adjusted {} Structural Columns successfully.".format(len(columns)), title="Success")
