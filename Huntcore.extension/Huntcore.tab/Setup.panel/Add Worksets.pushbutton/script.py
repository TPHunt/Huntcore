# -*- coding: utf-8 -*-
"""
Add predefined worksets by package selection.
Renames default worksets and sets 'Visible in all views' = True or False.

Package: Aus STR
    "00_Structures", True
    "01_Foundations", True
    "02_Core Walls", True
    "03_3D Reinforcement", True
    "80_Link RVT_STRUC (Name) Exist Model", True
    "81_Link RVT_ARCH (Name) Exist Model", False
    "81_Link RVT_ARCH (Name) Model", False
    "82_Link RVT_ELEC (Name) Model", False
    "83_Link RVT_FIRE (Name) Model", False
    "84_Link RVT_HYD (Name) Model", False
    "85_Link RVT_MECH (Name) Model", False
    "86_Link RVT_CIVIL (Name) Model", False
    "87_Link RVT_LANDSCAPE (Name) Model", False
    "88_Link DWG", False
    "89_Link IFC", False
    "90_Scopeboxes", False
    "99_Shared Levels and Grids", True

Package: Aus MEP
    Matches Aus STR but with 'MEP-' prefix

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from Autodesk.Revit.DB import (
    Workset, FilteredWorksetCollector, WorksetKind,
    Transaction, WorksetDefaultVisibilitySettings, RelinquishOptions, TransactWithCentralOptions, SynchronizeWithCentralOptions
)
from pyrevit import forms, revit

doc = revit.doc
uiapp = __revit__
app = uiapp.Application

# --- SETTINGS ---
AUTO = False   # set to True to skip popup and just create all

# Ensure model is workshared
if not doc.IsWorkshared:
    forms.alert("This model is not workshared.\nWorksets can only be added in a workshared model.", exitscript=True)

# --- PREDEFINED PACKAGES ---

AUS_STR = [
    ("00_Structures", True),
    ("01_Foundations", True),
    ("02_Core Walls", True),
    ("03_3D Reinforcement", True),
    ("80_Link RVT_STRUC (Name) Exist Model", True),
    ("81_Link RVT_ARCH (Name) Exist Model", False),
    ("81_Link RVT_ARCH (Name) Model", False),
    ("82_Link RVT_ELEC (Name) Model", False),
    ("83_Link RVT_FIRE (Name) Model", False),
    ("84_Link RVT_HYD (Name) Model", False),
    ("85_Link RVT_MECH (Name) Model", False),
    ("86_Link RVT_CIVIL (Name) Model", False),
    ("87_Link RVT_LANDSCAPE (Name) Model", False),
    ("88_Link DWG", False),
    ("89_Link IFC", False),
    ("90_Scopeboxes", False),
    ("99_Shared Levels and Grids", True),
]

# Generate Aus MEP from Aus STR
AUS_MEP = [("MEP-" + name, vis) for name, vis in AUS_STR]

PACKAGES = {
    "Aus STR": AUS_STR,
    "Aus MEP": AUS_MEP
}

# --- SELECT PACKAGE ---
if AUTO:
    selected_package_name = "Aus STR"
else:
    selected_package_name = forms.SelectFromList.show(
        sorted(PACKAGES.keys()),
        title="Select Workset Package",
        button_name="Select Package"
    )

if not selected_package_name:
    forms.alert("No package selected. Cancelled.", exitscript=True)

PREDEFINED_WORKSETS = PACKAGES[selected_package_name]

# --- SELECT WORKSETS FROM PACKAGE ---
choices = [ws[0] for ws in PREDEFINED_WORKSETS]
if AUTO:
    selected = choices
else:
    selected = forms.SelectFromList.show(
        choices,
        multiselect=True,
        title="Select Worksets to Create ({})".format(selected_package_name),
        button_name="Create Worksets",
        default=choices
    )

if not selected:
    forms.alert("No worksets selected. Cancelled.", exitscript=True)

# --- EXISTING WORKSETS ---
collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
existing_worksets = {ws.Name: ws for ws in collector}

predef_dict = dict(PREDEFINED_WORKSETS)
created, skipped, renamed, errors = [], [], [], []


# --- HELPER FUNCTIONS ---
def rename_workset_safely(workset_table, workset_id, new_name):
    """Handle RenameWorkset API differences across Revit versions"""
    try:
        workset_table.RenameWorkset(workset_id, new_name)  # Revit 2023 and earlier
        return True
    except Exception as e:
        if "takes exactly 3 arguments" in str(e):
            try:
                workset_table.RenameWorkset(doc, workset_id, new_name)  # Revit 2024+
                return True
            except Exception as e2:
                errors.append("Failed to rename workset to '{}': {}".format(new_name, e2))
                return False
        else:
            errors.append("Failed to rename workset to '{}': {}".format(new_name, e))
            return False


# --- MAIN TRANSACTION ---
t = Transaction(doc, "Manage Worksets")
try:
    t.Start()
    ws_table = doc.GetWorksetTable()

    # Rename defaults
    for ws in collector:
        if ws.Name == "Workset1":
            if rename_workset_safely(ws_table, ws.Id, "00_Structures"):
                WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc).SetWorksetVisibility(ws.Id, True)
                renamed.append("Workset1 -> 00_Structures")
        elif ws.Name == "Shared Levels and Grids":
            if rename_workset_safely(ws_table, ws.Id, "99_Shared Levels and Grids"):
                WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc).SetWorksetVisibility(ws.Id, True)
                renamed.append("Shared Levels and Grids -> 99_Shared Levels and Grids")

    # Create or update worksets
    for name in selected:
        if name in existing_worksets:
            skipped.append(name)
            wid = existing_worksets[name].Id
            WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc).SetWorksetVisibility(wid, predef_dict[name])
        else:
            try:
                new_ws = Workset.Create(doc, name)
                WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc).SetWorksetVisibility(new_ws.Id, predef_dict[name])
                created.append(name)
            except Exception as e:
                errors.append("Failed to create '{}': {}".format(name, e))

    t.Commit()

except Exception as e:
    if t.HasStarted():
        t.RollBack()
    forms.alert("Error: {}".format(e), exitscript=True)

# --- RELINQUISH ALL ---
try:
    relinq_opts = RelinquishOptions(True)
    relinq_opts.StandardWorksets = True
    relinq_opts.ViewWorksets = True
    relinq_opts.FamilyWorksets = True
    relinq_opts.UserWorksets = True
    relinq_opts.CheckedOutElements = True

    sync_opts = SynchronizeWithCentralOptions()
    sync_opts.SetRelinquishOptions(relinq_opts)
    sync_opts.Comment = "Relinquished by Huntcore Add Standard Worksets Tool"

    twc_opts = TransactWithCentralOptions()
    doc.SynchronizeWithCentral(twc_opts, sync_opts)
except Exception as e:
    errors.append("Relinquish failed: {}".format(e))


# --- REPORT BACK ---
msg = []
if created:
    msg.append("✅ Created:\n" + "\n".join(created))
if skipped:
    msg.append("⚠️ Already existed (visibility updated):\n" + "\n".join(skipped))
if renamed:
    msg.append("✏️ Renamed:\n" + "\n".join(renamed))
if errors:
    msg.append("❌ Errors:\n" + "\n".join(errors))

forms.alert("\n\n".join(msg) if msg else "No changes made.")
