# -*- coding: utf-8 -*-
"""
Add Predefined or Custom Worksets
- Adds predefined worksets by package selection
- Renames Workset1 to name containing "00"
- Adds prefix "99_" to Shared Levels and Grids
- Supports user-created custom packages (saved locally)
- Single-line, comma-separated or multi-line input with visibility toggle per workset
- Default visibility flag option for all new worksets

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from Autodesk.Revit.DB import (
    Workset, FilteredWorksetCollector, WorksetKind, Transaction,
    WorksetDefaultVisibilitySettings, RelinquishOptions,
    TransactWithCentralOptions, SynchronizeWithCentralOptions
)
from pyrevit import forms, revit
import os, json, sys, re
import System.Windows.Forms as WinForms

doc = revit.doc

# --- VERIFY WORKSHARING ---
if not doc.IsWorkshared:
    forms.alert(
        "This model is not workshared.\nWorksets can only be added in a workshared model.",
        exitscript=True
    )

# --- BUILT-IN PACKAGES ---
AUS_STR = [
    ("00_Structures", True),
    ("01_Foundations", True),
    ("02_Core Walls", True),
    ("03_3D Reinforcement", True),
    ("80_Link RVT_STRUC (Name) Exist Model", True),
    ("81_Link RVT_ARCH (Name) Exist Model", False),
    ("81_Link RVT_ARCH (Name) Model", False),
    ("82_Link RVT_MEP Combined Services (Name) Model", False),
    ("86_Link RVT_CIVIL (Name) Model", False),
    ("87_Link RVT_LANDSCAPE (Name) Model", False),
    ("88_Link DWG", False),
    ("89_Link IFC", False),
    ("90_Scopeboxes", False),
    ("99_Shared Levels and Grids", True),
]

AUS_MEP = [
    ("00_MEP Combined Services", True),
    ("80_Link RVT_STRUC (Name) Exist Model", True),
    ("81_Link RVT_ARCH (Name) Exist Model", False),
    ("81_Link RVT_ARCH (Name) Model", False),
    ("82_Link RVT_MEP Combined Services (Name) Exist Model", True),
    ("82_Link RVT_MEP Combined Services (Name) Model", True),
    ("86_Link RVT_CIVIL (Name) Model", False),
    ("87_Link RVT_LANDSCAPE (Name) Model", False),
    ("88_Link DWG", False),
    ("89_Link IFC", False),
    ("90_Scopeboxes", False),
    ("99_Shared Levels and Grids", True),
]

SYSTEM_PACKAGES = {"Aus STR": AUS_STR, "Aus MEP": AUS_MEP}

# --- USER CONFIG LOCATION ---
user_config_dir = os.path.join(os.getenv("APPDATA"), "pyRevit", "Huntcore")
user_config_file = os.path.join(user_config_dir, "user_workset_packages.json")
if not os.path.exists(user_config_dir):
    os.makedirs(user_config_dir)

# --- LOAD USER PACKAGES ---
USER_PACKAGES = {}
if os.path.exists(user_config_file):
    try:
        with open(user_config_file, "r") as f:
            USER_PACKAGES = json.load(f)
    except Exception as e:
        forms.alert("⚠️ Failed to load user packages:\n{}".format(e))

# --- MERGE ALL PACKAGES ---
def get_all_packages():
    pkgs = {}
    pkgs.update(SYSTEM_PACKAGES)
    pkgs.update(USER_PACKAGES)
    return pkgs

ALL_PACKAGES = get_all_packages()

# --- MULTI-LINE INPUT FOR IRONPYTHON ---
def ask_for_multiline_string(prompt="", default=""):
    """Shows a Windows Form for multi-line input, returns string or None."""
    form = WinForms.Form()
    form.Text = "Enter Worksets"
    form.Width = 500
    form.Height = 400
    form.StartPosition = WinForms.FormStartPosition.CenterScreen

    label = WinForms.Label()
    label.Text = prompt
    label.Top = 10
    label.Left = 10
    label.Width = 460
    label.Height = 40
    form.Controls.Add(label)

    textbox = WinForms.TextBox()
    textbox.Multiline = True
    textbox.ScrollBars = WinForms.ScrollBars.Vertical
    textbox.WordWrap = True
    textbox.Top = 60
    textbox.Left = 10
    textbox.Width = 460
    textbox.Height = 260
    textbox.Text = default
    form.Controls.Add(textbox)

    ok_button = WinForms.Button()
    ok_button.Text = "OK"
    ok_button.Top = 330
    ok_button.Left = 300
    ok_button.DialogResult = WinForms.DialogResult.OK
    form.Controls.Add(ok_button)

    cancel_button = WinForms.Button()
    cancel_button.Text = "Cancel"
    cancel_button.Top = 330
    cancel_button.Left = 380
    cancel_button.DialogResult = WinForms.DialogResult.Cancel
    form.Controls.Add(cancel_button)

    form.AcceptButton = ok_button
    form.CancelButton = cancel_button
    result = form.ShowDialog()

    if result == WinForms.DialogResult.OK:
        return textbox.Text
    return None

# ===============================================================
# ========== CUSTOM PACKAGE MANAGER =============================
# ===============================================================
def manage_custom_packages():
    """Interactive UI to create, edit, or delete user-defined workset packages."""
    global USER_PACKAGES, ALL_PACKAGES

    action = forms.SelectFromList.show(
        ["➕ Create New Package", "✏️ Edit Existing Package", "❌ Delete Package"],
        title="Manage Custom Packages",
        multiselect=False,
        button_name="Continue"
    )
    if not action:
        return

    # --- CREATE NEW PACKAGE ---
    if "Create" in action:
        name = forms.ask_for_string(
            prompt="Enter a name for your new package:",
            default="My Custom Package"
        )
        if not name:
            return

        default_vis = forms.alert(
            "Set all new worksets as 'Visible in all views'?", yes=True, no=True
        )
        if default_vis is None:
            return
        default_vis_flag = True if default_vis else False

        worksets_str = ask_for_multiline_string(
            prompt="Enter workset names, separated by commas or newlines:\n"
                   "Prefix + (visible) or - (not visible), e.g. +00_Structures,-01_Foundations",
            default="00_Custom,01_MyWorkset"
        )
        if not worksets_str:
            return

        worksets = []
        for item in re.split(r'[,\n]+', worksets_str):
            line = item.strip()
            if not line:
                continue
            if line.startswith("+"):
                ws_name = line[1:].strip()
                ws_vis = True
            elif line.startswith("-"):
                ws_name = line[1:].strip()
                ws_vis = False
            else:
                ws_name = line
                ws_vis = default_vis_flag
            worksets.append((ws_name, ws_vis))

        USER_PACKAGES[name] = worksets

    # --- EDIT EXISTING PACKAGE ---
    elif "Edit" in action:
        if not USER_PACKAGES:
            forms.alert("No custom packages available to edit.")
            return
        selected = forms.SelectFromList.show(
            sorted(USER_PACKAGES.keys()),
            title="Select package to edit"
        )
        if not selected:
            return

        pkg = USER_PACKAGES[selected]
        parts = []
        for n, v in pkg:
            prefix = "+" if v else "-"
            parts.append("{}{}".format(prefix, n))
        worksets_str = ", ".join(parts)

        edited_str = ask_for_multiline_string(
            prompt="Edit workset names, separated by commas or newlines:\nUse + or - to control visibility",
            default=worksets_str
        )
        if not edited_str:
            return

        new_pkg = []
        for item in re.split(r'[,\n]+', edited_str):
            line = item.strip()
            if not line:
                continue
            if line.startswith("+"):
                ws_name = line[1:].strip()
                ws_vis = True
            elif line.startswith("-"):
                ws_name = line[1:].strip()
                ws_vis = False
            else:
                ws_name = line
                ws_vis = True
            new_pkg.append((ws_name, ws_vis))

        USER_PACKAGES[selected] = new_pkg

    # --- DELETE PACKAGE ---
    elif "Delete" in action:
        if not USER_PACKAGES:
            forms.alert("No custom packages available to delete.")
            return
        selected = forms.SelectFromList.show(
            sorted(USER_PACKAGES.keys()),
            title="Select package to delete"
        )
        if not selected:
            return
        del USER_PACKAGES[selected]
        forms.alert("🗑️ Package '{}' deleted.".format(selected))

    # --- SAVE CHANGES ---
    try:
        with open(user_config_file, "w") as f:
            json.dump(USER_PACKAGES, f, indent=4)
        forms.alert("✅ Custom packages saved successfully.\nReloading tool...")
    except Exception as e:
        forms.alert("Failed to save packages:\n{}".format(e))

    # ✅ AUTO-RERUN SCRIPT
    try:
        execfile(__file__)
    except Exception as e:
        forms.alert("Failed to reload tool: {}".format(e))
    sys.exit()

# ===============================================================
# ========== MAIN WORKFLOW ======================================
# ===============================================================
menu_choices = sorted(ALL_PACKAGES.keys()) + ["🧩 Manage Custom Packages"]
selected_package_name = forms.SelectFromList.show(
    menu_choices,
    title="Select Workset Package",
    button_name="Select Package"
)
if not selected_package_name:
    forms.alert("No package selected. Cancelled.", exitscript=True)

if "Manage Custom" in selected_package_name:
    manage_custom_packages()
    sys.exit()

PREDEFINED_WORKSETS = ALL_PACKAGES[selected_package_name]

# --- DETERMINE NEW NAME FOR Workset1 ---
if "STR" in selected_package_name:
    workset1_newname = "00_Structures"
elif "MEP" in selected_package_name:
    workset1_newname = "00_MEP Combined Services"
else:
    workset1_newname = PREDEFINED_WORKSETS[0][0] if PREDEFINED_WORKSETS else None

# --- SELECT WORKSETS FROM PACKAGE ---
choices = [ws[0] for ws in PREDEFINED_WORKSETS]
default_indices = list(range(len(choices)))  # Pre-check all worksets

selected = forms.SelectFromList.show(
    choices,
    multiselect=True,
    title="Select Worksets to Create ({})".format(selected_package_name),
    button_name="Create Worksets",
    default=default_indices
)
if not selected:
    forms.alert("No worksets selected. Cancelled", exitscript=True)

# --- EXISTING WORKSETS ---
collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
existing_worksets = {ws.Name: ws for ws in collector}
predef_dict = dict(PREDEFINED_WORKSETS)
created, skipped, renamed, errors = [], [], [], []

def rename_workset_safely(workset_table, workset_id, new_name):
    """Handles rename for compatibility across Revit versions."""
    try:
        workset_table.RenameWorkset(workset_id, new_name)
        return True
    except Exception:
        try:
            workset_table.RenameWorkset(doc, workset_id, new_name)
            return True
        except Exception as e2:
            errors.append("Failed to rename to '{}': {}".format(new_name, e2))
            return False

# --- MAIN TRANSACTION ---
t = Transaction(doc, "Manage Worksets")
try:
    t.Start()
    ws_table = doc.GetWorksetTable()
    vis_settings = WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc)

    # --- RENAME DEFAULT WORKSETS ---
    for ws in collector:
        if ws.Name == "Workset1" and workset1_newname:
            if rename_workset_safely(ws_table, ws.Id, workset1_newname):
                vis_settings.SetWorksetVisibility(ws.Id, True)
                renamed.append("Workset1 → {}".format(workset1_newname))
        elif ws.Name == "Shared Levels and Grids":
            if rename_workset_safely(ws_table, ws.Id, "99_Shared Levels and Grids"):
                vis_settings.SetWorksetVisibility(ws.Id, True)
                renamed.append("Shared Levels and Grids → 99_Shared Levels and Grids")

    # --- REFRESH EXISTING ---
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    existing_worksets = {ws.Name: ws for ws in collector}

    # --- CREATE OR UPDATE WORKSETS ---
    for name in selected:
        vis_flag = predef_dict.get(name, True)
        if name in existing_worksets:
            skipped.append(name)
            wid = existing_worksets[name].Id
            vis_settings.SetWorksetVisibility(wid, vis_flag)
            continue
        try:
            new_ws = Workset.Create(doc, name)
            vis_settings.SetWorksetVisibility(new_ws.Id, vis_flag)
            created.append(name)
        except Exception as e:
            errors.append("Failed to create '{}': {}".format(name, e))

    t.Commit()
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    forms.alert("Error: {}".format(e), exitscript=True)

# --- RELINQUISH ---
try:
    relinq_opts = RelinquishOptions(True)
    relinq_opts.StandardWorksets = relinq_opts.ViewWorksets = True
    relinq_opts.FamilyWorksets = relinq_opts.UserWorksets = True
    relinq_opts.CheckedOutElements = True

    sync_opts = SynchronizeWithCentralOptions()
    sync_opts.SetRelinquishOptions(relinq_opts)
    sync_opts.Comment = "Relinquished by Huntcore Add Standard Worksets Tool"

    twc_opts = TransactWithCentralOptions()
    doc.SynchronizeWithCentral(twc_opts, sync_opts)
except Exception as e:
    errors.append("Relinquish failed: {}".format(e))

# --- REPORT ---
msg = []
if created:
    msg.append("✅ Created:\n" + "\n".join(created))
if skipped:
    msg.append("⚠️ Skipped (already existed):\n" + "\n".join(skipped))
if renamed:
    msg.append("✏️ Renamed:\n" + "\n".join(renamed))
if errors:
    msg.append("❌ Errors:\n" + "\n".join(errors))

forms.alert("\n\n".join(msg) if msg else "No changes made.")
