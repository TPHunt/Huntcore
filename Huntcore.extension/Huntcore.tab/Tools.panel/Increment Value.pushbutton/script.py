# -*- coding: utf-8 -*-
"""
Specify parameter to update.
Specify amount to increment up or down for the last number of the value.
Ie. 4B26 -> 4B27

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms, revit
import re

doc = revit.doc

# ------------------------------------------------------------
# Get user selection
# ------------------------------------------------------------
selection = revit.get_selection()
if not selection:
    TaskDialog.Show("Error", "No elements selected.")
    raise SystemExit

# ------------------------------------------------------------
# Collect available string parameters from selection
# ------------------------------------------------------------
param_names = set()
for el in selection:
    for p in el.Parameters:
        if p.StorageType == StorageType.String:
            param_names.add(p.Definition.Name)

if not param_names:
    TaskDialog.Show("Error", "No string parameters found on selected elements.")
    raise SystemExit

# ------------------------------------------------------------
# Ask user which parameter to update
# ------------------------------------------------------------
param_name = forms.SelectFromList.show(
    sorted(list(param_names)),
    title="Select Parameter to Update",
    button_name="Use Parameter"
)

if not param_name:
    raise SystemExit

# ------------------------------------------------------------
# Ask user for increment AFTER choosing parameter
# ------------------------------------------------------------
try:
    increment = int(forms.ask_for_string(
        prompt="Enter increment (+/-):",
        default="1",
        title="Increment for {}".format(param_name)
    ))
except:
    TaskDialog.Show("Error", "Invalid input. Must be an integer.")
    raise SystemExit

# ------------------------------------------------------------
# Function to calculate new value
# ------------------------------------------------------------
def update_value(old_val, increment):
    if not old_val:
        return old_val
    match = re.search(r"(\d+)(?!.*\d)", old_val)   # last number block
    if not match:
        return old_val
    number = match.group(1)
    start, end = match.span()
    new_number = str(int(number) + increment).zfill(len(number))
    return old_val[:start] + new_number + old_val[end:]

# ------------------------------------------------------------
# Build preview values
# ------------------------------------------------------------
preview_rows = []
element_map = {}   # element -> (old_val, new_val)

for el in selection:
    p = el.LookupParameter(param_name)
    if p and p.HasValue:
        old_val = p.AsString()
        new_val = update_value(old_val, increment)
        element_map[el] = (old_val, new_val)
        preview_rows.append("{}  →  {}".format(old_val, new_val))

if not preview_rows:
    TaskDialog.Show("Notice", "No values to update.")
    raise SystemExit

# ------------------------------------------------------------
# Show preview alert (informational only)
# ------------------------------------------------------------
preview_text = "\n".join(preview_rows)
forms.alert(
    preview_text,
    title="Preview Changes ({} values)".format(param_name),
    ok=True
)

# ------------------------------------------------------------
# Apply changes
# ------------------------------------------------------------
t = Transaction(doc, "Update {}".format(param_name))
t.Start()
count = 0
for el, (old_val, new_val) in element_map.items():
    if new_val and old_val != new_val:
        p = el.LookupParameter(param_name)
        if p:
            p.Set(new_val)
            count += 1
t.Commit()

TaskDialog.Show("Done", "Updated {} element(s).".format(count))
