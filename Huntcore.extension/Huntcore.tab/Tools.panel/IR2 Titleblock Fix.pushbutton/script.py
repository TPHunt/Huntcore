# -*- coding: utf-8 -*-
# pyRevit script: Align "Arup Revision List_GRS" titleblocks
# with "Arup_Titleblock_A0 Vertical_ISO19650_SG_GRS_IR2"
# and apply custom user-defined offset (X,Y in mm).
#
# Compatible with Revit 2023+ (tested under pyRevit IronPython)

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    XYZ,
    Transaction,
    ViewSheet,
    ElementTransformUtils
)

doc = revit.doc

# Family names (use only family names, not types)
FAMILY_A0 = "Arup_Titleblock_A0 Vertical_ISO19650_SG_GRS_IR2"
FAMILY_REV = "Arup Revision List_GRS"


def get_titleblocks_on_sheet(sheet, fam_name):
    """Return all titleblock instances of a given family on a sheet."""
    return [
        tb for tb in FilteredElementCollector(doc)
        .OwnedByView(sheet.Id)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
        if tb.Symbol.Family.Name == fam_name
    ]


def main():
    # --- Prompt user to pick sheets ---
    sheets = forms.select_sheets(
        title="Select Sheets to Process",
        use_selection=True,
        multiple=True
    )
    if not sheets:
        forms.alert("No sheets selected.", exitscript=True)

    # --- Ask for X and Y offsets (in mm) ---
    x_str = forms.ask_for_string(
        prompt="Enter X distance (mm):",
        title="Custom Offset",
        default="0"
    )
    y_str = forms.ask_for_string(
        prompt="Enter Y distance (mm):",
        title="Custom Offset",
        default="0"
    )

    try:
        x_offset = float(x_str) / 304.8  # convert mm → feet
    except:
        x_offset = 0.0
    try:
        y_offset = float(y_str) / 304.8
    except:
        y_offset = 0.0

    # --- Process each sheet ---
    results = []
    with Transaction(doc, "Align Revision List Titleblocks") as t:
        t.Start()

        for sheet in sheets:
            a0_blocks = get_titleblocks_on_sheet(sheet, FAMILY_A0)
            rev_blocks = get_titleblocks_on_sheet(sheet, FAMILY_REV)

            if not a0_blocks:
                results.append(
                    "Sheet {0} ('{1}'): No {2} found.".format(
                        sheet.SheetNumber, sheet.Name, FAMILY_A0
                    )
                )
                continue
            if not rev_blocks:
                results.append(
                    "Sheet {0} ('{1}'): No {2} found.".format(
                        sheet.SheetNumber, sheet.Name, FAMILY_REV
                    )
                )
                continue

            # Use the first A0 instance as reference
            ref_pt = a0_blocks[0].GetTransform().Origin

            moved_count = 0
            for rev_tb in rev_blocks:
                # Move to reference point
                rev_pt = rev_tb.GetTransform().Origin
                delta = ref_pt - rev_pt
                ElementTransformUtils.MoveElement(doc, rev_tb.Id, delta)

                # Apply user offset
                offset_vec = XYZ(x_offset, y_offset, 0)
                ElementTransformUtils.MoveElement(doc, rev_tb.Id, offset_vec)

                moved_count += 1

            results.append(
                "Sheet {0} ('{1}'): Moved {2} '{3}' instance(s).".format(
                    sheet.SheetNumber, sheet.Name, moved_count, FAMILY_REV
                )
            )

        t.Commit()

    # --- Report ---
    forms.alert("\n".join(results), title="Alignment Results")


# --- Run ---
if __name__ == "__main__":
    main()
