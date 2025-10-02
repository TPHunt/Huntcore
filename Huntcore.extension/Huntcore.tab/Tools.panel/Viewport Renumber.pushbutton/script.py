# -*- coding: utf-8 -*-
"""Huntcore: Renumber viewports on active sheet and optionally rename views
 - Orders viewports top-left -> right, then next row top-left -> right
 - Supports numeric and alphanumeric increment systems:
    1, 2, 3, ...
    A, B, C, ...
    AA, AB, AC, ...
    A1, A2, A3, ...
 - User inputs starting detail number (can be text or number, e.g. "1", "A1", "D-01")
 - Optionally, user can provide a text string to rename views to "<text> <detail number>"
 - Conflicts handled with user choice: append 'x' or skip
Author: Troy Hunt | Huntcore
"""

from Autodesk.Revit.DB import (
    ViewSheet,
    View,
    Viewport,
    BuiltInParameter,
    ViewType,
    Transaction,
    TransactionGroup,
    FilteredElementCollector
)
from pyrevit import revit, forms, script
import re

doc = revit.doc


# ---------------- HELPERS ---------------- #
def get_active_sheet():
    v = doc.ActiveView
    if isinstance(v, ViewSheet):
        return v
    return None


def viewport_center_and_height(vp):
    try:
        outline = vp.GetBoxOutline()
        minp = outline.MinimumPoint
        maxp = outline.MaximumPoint
        cx = (minp.X + maxp.X) / 2.0
        cy = (minp.Y + maxp.Y) / 2.0
        h = abs(maxp.Y - minp.Y)
        return (cx, cy, h)
    except:
        try:
            c = vp.GetBoxCenter()
            return (c.X, c.Y, 0.0)
        except:
            return (0.0, 0.0, 0.0)


def collect_non_legend_viewports_on_sheet(sheet):
    vp_ids = list(sheet.GetAllViewports())
    vps = [doc.GetElement(eid) for eid in vp_ids]
    filtered = []
    for vp in vps:
        if vp is None:
            continue
        try:
            linked_view = doc.GetElement(vp.ViewId)
            if linked_view is None:
                continue
            if linked_view.ViewType == ViewType.Legend:
                continue
        except:
            continue
        filtered.append(vp)
    return filtered


def all_project_views():
    return [v for v in FilteredElementCollector(doc).OfClass(View).ToElements() if not v.IsTemplate]


def make_unique_name(existing_names, base_name):
    new_name = base_name
    while new_name in existing_names:
        new_name = new_name + 'x'
    return new_name


# ---------------- DETAIL NUMBERING ---------------- #
def int_to_letters(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def generate_alphanumeric(start_str, index):
    # Numeric
    if re.match(r"^\d+$", start_str):
        num = int(start_str)
        return str(num + index)
    # Letters only
    elif re.match(r"^[A-Z]+$", start_str, re.I):
        base_num = sum((ord(c.upper()) - 64) * (26 ** i) for i, c in enumerate(reversed(start_str)))
        return int_to_letters(base_num + index)
    # Letter + Number (A1, B2, etc.)
    elif re.match(r"^([A-Z]+)(\d+)$", start_str, re.I):
        m = re.match(r"^([A-Z]+)(\d+)$", start_str, re.I)
        prefix = m.group(1).upper()
        num = int(m.group(2))
        return "{0}{1}".format(prefix, num + index)
    # Fallback: append index
    else:
        if index == 0:
            return start_str
        return "{0}{1}".format(start_str, index)


# ---------------- MAIN ---------------- #
def main():
    sheet = get_active_sheet()
    if not sheet:
        forms.alert("Please run this tool on an active Sheet view.", title="Huntcore: Viewport Renumber")
        script.exit()

    start_str = forms.ask_for_string(prompt="Start Detail Numbers from:", default="1")
    if start_str is None:
        script.exit()

    filtered_vps = collect_non_legend_viewports_on_sheet(sheet)
    if not filtered_vps:
        forms.alert("No viewports found on this sheet (excluding legends).", title="Huntcore: Renumber Viewports")
        script.exit()

    vp_info = []
    for vp in filtered_vps:
        cx, cy, h = viewport_center_and_height(vp)
        vp_info.append({'vp': vp, 'cx': cx, 'cy': cy, 'h': abs(h)})

    heights = sorted([it['h'] for it in vp_info if it['h'] > 1e-12])
    row_tol = max(heights[len(heights)//2]*0.5, 1e-5) if heights else 1e-5

    vp_info_sorted = sorted(vp_info, key=lambda x: (-x['cy'], x['cx']))

    rows = []
    current_row = []
    last_y = None
    for info in vp_info_sorted:
        if last_y is None:
            current_row = [info]
            rows.append(current_row)
            last_y = info['cy']
        else:
            if abs(last_y - info['cy']) <= row_tol:
                current_row.append(info)
            else:
                current_row = [info]
                rows.append(current_row)
                last_y = info['cy']

    ordered_vps = []
    for row in rows:
        row_sorted = sorted(row, key=lambda r: r['cx'])
        ordered_vps.extend([r['vp'] for r in row_sorted])

    tg = TransactionGroup(doc, "Huntcore: Renumber & Rename Viewports")
    tg.Start()

    # Prefix old detail numbers with 'x'
    t1 = Transaction(doc, "Huntcore: Prefix detail numbers")
    t1.Start()
    for vp in ordered_vps:
        try:
            param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
            if param and not param.IsReadOnly:
                old = param.AsString() or ''
                if not old.startswith('x'):
                    param.Set('x' + old)
        except:
            pass
    t1.Commit()

    # Set new detail numbers
    t2 = Transaction(doc, "Huntcore: Set detail numbers")
    t2.Start()
    for idx, vp in enumerate(ordered_vps):
        newnum = generate_alphanumeric(start_str, idx)
        try:
            param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
            if param and not param.IsReadOnly:
                param.Set(newnum)
        except:
            pass
    t2.Commit()

    try:
        doc.Regenerate()
    except:
        pass

    # Rename views option
    rename = forms.alert(
        "Do you want to update the view names\nwith a text prefix and new detail number?",
        options=["Yes", "No"],
        title="Huntcore: Rename Views"
    )
    if rename == "Yes":
        user_text = forms.ask_for_string(
            prompt="Update view names with prefix below and new Detail Number:",
            default="Sh - "
        )
        if user_text is None:
            tg.Assimilate()
            forms.alert("Renumbering complete. Rename cancelled by user.", title="Huntcore: Renumber Viewports")
            script.exit()

        all_views = all_project_views()
        existing_names = set([v.Name for v in all_views if v.Name is not None])

        t3 = Transaction(doc, "Huntcore: Rename views to '<text> <detail>'")
        t3.Start()
        for idx, vp in enumerate(ordered_vps):
            newnum = generate_alphanumeric(start_str, idx)
            desired_name = (user_text + " " + newnum).strip()
            try:
                view_el = doc.GetElement(vp.ViewId)
                if view_el is None:
                    continue

                # Handle conflicts
                if desired_name in existing_names:
                    choice = forms.alert(
                        "Conflict: View name '{}' already exists. How to handle?".format(desired_name),
                        options=["Append x", "Skip"],
                        title="Huntcore: Rename Conflict"
                    )
                    if choice == "Append x":
                        desired_name = make_unique_name(existing_names, desired_name)
                    elif choice == "Skip":
                        continue

                view_el.Name = desired_name
                existing_names.add(desired_name)
            except:
                pass
        t3.Commit()

    try:
        tg.Assimilate()
    except:
        try:
            tg.Commit()
        except:
            pass

    forms.alert(
        "Operation finished.\nTotal viewports updated: {}".format(len(ordered_vps)),
        title="Huntcore: Renumber & Rename"
    )


if __name__ == "__main__":
    main()