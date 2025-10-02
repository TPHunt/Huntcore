# -*- coding: utf-8 -*-
# Huntcore Script | Viewport Renumber
# Author: Troy Hunt
# Revit 2023+ | pyRevit Extension

from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

def get_active_sheet():
    view = doc.ActiveView
    if isinstance(view, ViewSheet):
        return view
    return None

def is_legend(view):
    return isinstance(view, View) and view.ViewType == ViewType.Legend

def get_viewport_center(viewport):
    outline = viewport.GetBoxOutline()
    min_pt = outline.MinimumPoint
    max_pt = outline.MaximumPoint
    return XYZ((min_pt.X + max_pt.X) / 2.0,
               (min_pt.Y + max_pt.Y) / 2.0,
               0)

def renumber_viewports(sheet, start_num):
    viewports = [doc.GetElement(vpid) for vpid in sheet.GetAllViewports()]

    # Filter out legend views
    valid_vps = []
    for vp in viewports:
        view = doc.GetElement(vp.ViewId)
        if not is_legend(view):
            center = get_viewport_center(vp)
            valid_vps.append((vp, center))

    # Sort top-to-bottom, then left-to-right
    # Y descending (top), then X ascending (left)
    valid_vps.sort(key=lambda x: (-x[1].Y, x[1].X))

    # First prefix all detail numbers to avoid duplicates
    t = Transaction(doc, "Prefix Detail Numbers")
    t.Start()
    for vp, _ in valid_vps:
        param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        if param and not param.IsReadOnly:
            current = param.AsString()
            if current and not current.startswith("x"):
                param.Set("x" + current)
    t.Commit()

    # Then assign new sequential numbers
    t2 = Transaction(doc, "Renumber Viewports")
    t2.Start()
    num = start_num
    for vp, _ in valid_vps:
        param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        if param and not param.IsReadOnly:
            param.Set(str(num))
            num += 1
    t2.Commit()

def main():
    sheet = get_active_sheet()
    if not sheet:
        forms.alert("Please run this tool on an active Sheet view.", exitscript=True)

    start_num = forms.ask_for_string(
        prompt="Enter starting detail number:",
        default="1",
        title="Viewport Renumber"
    )

    if not start_num:
        forms.alert("Cancelled.", exitscript=True)

    try:
        start_num = int(start_num)
    except:
        forms.alert("Starting number must be an integer.", exitscript=True)

    renumber_viewports(sheet, start_num)
    forms.alert("Viewports renumbered successfully.", exitscript=False)

main()
