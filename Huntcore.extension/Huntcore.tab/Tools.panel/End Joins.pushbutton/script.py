# -*- coding: utf-8 -*-
"""
Allow or Disallow structural framing joins of selected framing members.
Option for Start, End, or Both ends.

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType

# --- Helper checks ---------------------------------------------------------
def is_structural_framing(elem):
    """Return True if element is a FamilyInstance in Structural Framing category."""
    try:
        if not isinstance(elem, DB.FamilyInstance):
            return False
        if not elem.Category:
            return False
        return elem.Category.Id == DB.ElementId(DB.BuiltInCategory.OST_StructuralFraming)
    except:
        return False

# --- Get selection / allow user to pick -----------------------------------
selection = revit.get_selection()
elements = [e for e in selection] if selection else []

if not elements:
    res = forms.alert(
        "No elements selected.\n\n"
        "Choose how you want to select structural framing elements:",
        options=["Pick in model", "Cancel"],
        warn_icon=False
    )
    if res != "Pick in model":
        script.exit()
    try:
        refs = revit.uidoc.Selection.PickObjects(ObjectType.Element, "Select structural framing elements")
        elements = [revit.doc.GetElement(r) for r in refs]
    except Exception:
        script.exit()

framing_elems = [e for e in elements if is_structural_framing(e)]

if not framing_elems:
    forms.alert("No structural framing family instances found in selection.", title="Nothing to do")
    script.exit()

# --- Ask user: Enable or Disable ------------------------------------------
mode = forms.alert(
    "Do you want to ENABLE (Allow) or DISABLE (Disallow) joins?",
    options=["Enable", "Disable", "Cancel"],
    title="Framing Joins Mode"
)

if mode == "Cancel" or not mode:
    script.exit()

enable_mode = (mode == "Enable")

# --- Ask user which end(s) -------------------------------------------------
choice = forms.alert(
    "Apply to which end(s)?\n\nStart = start end (0)\nEnd = end end (1)\nBoth = both ends",
    options=["Start", "End", "Both", "Cancel"],
    title="Select Ends"
)

if choice == "Cancel" or not choice:
    script.exit()

if choice == "Start":
    ends = [0]
elif choice == "End":
    ends = [1]
else:
    ends = [0, 1]

# --- Apply inside transaction ---------------------------------------------
processed = 0
failed = []

with revit.Transaction("{} Framing Joins".format("Enable" if enable_mode else "Disable")):
    for el in framing_elems:
        try:
            for end_idx in ends:
                if enable_mode:
                    DB.Structure.StructuralFramingUtils.AllowJoinAtEnd(el, end_idx)
                else:
                    DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(el, end_idx)
            processed += 1
        except Exception as ex:
            try:
                eid = el.Id.IntegerValue
            except:
                eid = str(getattr(el, "Id", "Unknown"))
            failed.append((eid, str(ex)))

# --- Summary ---------------------------------------------------------------
msg = "{0}d joins on {1} framing elements.".format(
    "Enable" if enable_mode else "Disable", processed
)

if failed:
    msg += "\n\nFailed to update {0} elements:\n".format(len(failed))
    for fid, ferr in failed[:10]:
        msg += "- ID {0}: {1}\n".format(fid, ferr)
    if len(failed) > 10:
        msg += "... (and {0} more)".format(len(failed) - 10)

forms.alert(msg, title="Framing Joins — Done")
