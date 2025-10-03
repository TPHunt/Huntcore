# -*- coding: utf-8 -*-

"""
Select elements last updated by the specified user. 

!!! NOTE: This will run on all elements in the current view - Full model in view will be sloooow !!!


Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from collections import defaultdict
from pyrevit import revit, DB, forms
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# Collect only elements visible in the current view (exclude element types)
view = doc.ActiveView
elem_ids = list(DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElementIds())

user_to_ids = defaultdict(list)
no_info = 0

# Progress bar while scanning
with forms.ProgressBar(title="Scanning elements in current view ({value} of {max_value})") as pb:
    for i, eid in enumerate(elem_ids, 1):
        try:
            # Get worksharing tooltip info (contains LastChangedBy, Creator, Owner)
            wti = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, eid)
            last_changed_by = (wti.LastChangedBy or "").strip()
            if last_changed_by:
                user_to_ids[last_changed_by].append(eid)
            else:
                no_info += 1
        except Exception:
            no_info += 1
        pb.update_progress(i, len(elem_ids))

if not user_to_ids:
    forms.alert(
        "No 'Last Changed By' data found for elements visible in this view.\n\n"
        "Possible causes:\n"
        "• These elements have never been synced to central since creation\n"
        "• Detached/recreated central (history reset)\n"
        "• Mostly non-workshared/system elements\n\n"
        "Tip: Try modifying & syncing a few visible elements, then rerun.",
        exitscript=True
    )

# Build a pick-list with usernames and counts
items = []
display_to_user = {}
for user in sorted(user_to_ids.keys(), key=lambda s: s.lower()):
    label = u"{} ({})".format(user, len(user_to_ids[user]))
    items.append(label)
    display_to_user[label] = user

picked = forms.SelectFromList.show(
    items,
    title="Select user (Last Changed By in current view)",
    multiselect=False,
    button_name="Select"
)

if not picked:
    forms.alert("No user selected.", exitscript=True)

chosen_user = display_to_user[picked]
ids_to_select = user_to_ids[chosen_user]

# Wrap the list in System.Collections.Generic.List[ElementId] for IronPython
ids_collection = List[DB.ElementId](ids_to_select)

# Apply selection in Revit
uidoc.Selection.SetElementIds(ids_collection)

forms.alert(
    "Selected {} element(s) in the current view last updated in central by:\n\n{}".format(len(ids_to_select), chosen_user),
    exitscript=False
)
