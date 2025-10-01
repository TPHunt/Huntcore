# -*- coding: utf-8 -*-
"""
Huntcore: Duplicate placed view on sheet (Revit 2023+)
Run this from an active Sheet. 
Select the placed view (viewport) on the sheet before running.
Options: Duplicate, Duplicate with Detailing
"""

from pyrevit import revit, forms, script
import Autodesk.Revit.DB as DB

uidoc = revit.uidoc
doc = revit.doc

def get_selected_viewport_or_find_on_sheet():
    sel_ids = uidoc.Selection.GetElementIds()
    if not sel_ids:
        return None, "No element selected. Please select the placed view (viewport) on the active sheet."
    first_id = list(sel_ids)[0]
    el = doc.GetElement(first_id)
    if isinstance(el, DB.Viewport):
        return el, None
    if isinstance(el, DB.View):
        sheet = doc.ActiveView
        if not isinstance(sheet, DB.ViewSheet):
            return None, "Tool must be run from an active sheet view."
        collector = DB.FilteredElementCollector(doc, sheet.Id).OfClass(DB.Viewport)
        for vp in collector:
            if vp.ViewId == el.Id:
                return vp, None
        return None, "Selected view is not placed on the active sheet."
    return None, "Selected element is not a viewport or view. Please select a placed view."

def can_duplicate_view(view, dup_option):
    try:
        return view.CanViewBeDuplicated(dup_option)
    except Exception:
        try:
            return view.CanBeDuplicated(dup_option)
        except Exception:
            return True

def main():
    sheet = doc.ActiveView
    if not isinstance(sheet, DB.ViewSheet):
        forms.alert("Please run this tool from a Sheet view (open the sheet first).")
        return

    viewport, err = get_selected_viewport_or_find_on_sheet()
    if not viewport:
        forms.alert(err)
        return

    orig_view = doc.GetElement(viewport.ViewId)
    if orig_view is None:
        forms.alert("Could not resolve the view from the selected viewport.")
        return

    choice = forms.CommandSwitchWindow.show(
        ["Duplicate", "Duplicate with Detailing"],
        message="Choose duplication option"
    )
    if not choice:
        return

    dup_option = DB.ViewDuplicateOption.Duplicate if choice == "Duplicate" else DB.ViewDuplicateOption.WithDetailing

    if not can_duplicate_view(orig_view, dup_option):
        forms.alert("This view cannot be duplicated with the chosen option.")
        return

    try:
        with revit.Transaction("Huntcore: Duplicate view on sheet"):
            new_vid = orig_view.Duplicate(dup_option)
            new_view = doc.GetElement(new_vid)
            if new_view is None:
                forms.alert("Duplicate created but could not retrieve the new view element.")
                return

            # placement location
            center = viewport.GetBoxCenter()
            offset = DB.XYZ(1.0, -0.5, 0)  # offset to avoid exact overlap
            target = DB.XYZ(center.X + offset.X, center.Y + offset.Y, center.Z + offset.Z)

            # create viewport
            new_vp = DB.Viewport.Create(doc, sheet.Id, new_view.Id, target)

            # 🔑 ensure viewport type matches original
            new_vp.ChangeTypeId(viewport.GetTypeId())

        forms.alert('Duplicated view "{}" as "{}" and placed on sheet "{}" with matching viewport type.'.format(orig_view.Name, new_view.Name, sheet.Name))
    except Exception as e:
        forms.alert("Error: {}".format(e))

if __name__ == "__main__":
    main()
