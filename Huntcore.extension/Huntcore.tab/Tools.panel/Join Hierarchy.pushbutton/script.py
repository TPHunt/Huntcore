# -*- coding: utf-8 -*-
"""
Enforce Join Hierarchy on Already Joined Elements (pyRevit)
- UI: checkbox list + up/down reorder
- Only enforces join order on elements that are already joined
- Options: Only elements in current view or entire model

Huntcore Script | Author: Troy Hunt | Revit 2023+ | pyRevit Extension
"""

import clr
from pyrevit import revit, forms
from Autodesk.Revit.DB import *
from System import Array

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (
    Form, CheckedListBox, Button, Label, CheckBox, MessageBox, MessageBoxButtons,
    DialogResult, SelectionMode, FormStartPosition, RadioButton, GroupBox
)
from System.Drawing import Point, Size

# --- Category mapping ---
CATEGORY_UI_ORDER = [
    "Floors",
    "Structural Beams",
    "Structural Columns",
    "Structural Foundations",
    "Walls"
]

CAT_MAP = {
    "Floors": BuiltInCategory.OST_Floors,
    "Structural Beams": BuiltInCategory.OST_StructuralFraming,
    "Structural Columns": BuiltInCategory.OST_StructuralColumns,
    "Structural Foundations": BuiltInCategory.OST_StructuralFoundation,
    "Walls": BuiltInCategory.OST_Walls
}

# --- UI ---
class HierarchyDialog(Form):
    def __init__(self):
        self.Text = "Enforce Join Order on Already Joined Elements"
        self.ClientSize = Size(520, 400)
        self.StartPosition = FormStartPosition.CenterParent

        # View scope group
        scope_group = GroupBox()
        scope_group.Text = "Scope"
        scope_group.Location = Point(10, 10)
        scope_group.Size = Size(500, 60)
        self.Controls.Add(scope_group)

        self.rb_current_view = RadioButton()
        self.rb_current_view.Text = "Current view"
        self.rb_current_view.Location = Point(10, 20)
        self.rb_current_view.Checked = True
        scope_group.Controls.Add(self.rb_current_view)

        self.rb_entire_model = RadioButton()
        self.rb_entire_model.Text = "Entire model"
        self.rb_entire_model.Location = Point(220, 20)
        scope_group.Controls.Add(self.rb_entire_model)

        lbl = Label()
        lbl.Text = "Tick categories to include, select one item and use ▲/▼ to set order (top = highest priority)."
        lbl.Location = Point(10, 70)
        lbl.Size = Size(500, 28)
        self.Controls.Add(lbl)

        self.chklist = CheckedListBox()
        self.chklist.Location = Point(10, 100)
        self.chklist.Size = Size(300, 240)
        self.chklist.SelectionMode = SelectionMode.One
        for name in CATEGORY_UI_ORDER:
            self.chklist.Items.Add(name, True)
        self.Controls.Add(self.chklist)

        self.btn_up = Button(Text="▲ Move Up", Location=Point(330, 120), Size=Size(160, 30))
        self.btn_up.Click += self.on_move_up
        self.Controls.Add(self.btn_up)

        self.btn_down = Button(Text="▼ Move Down", Location=Point(330, 160), Size=Size(160, 30))
        self.btn_down.Click += self.on_move_down
        self.Controls.Add(self.btn_down)

        self.btn_select_all = Button(Text="Select All", Location=Point(330, 210), Size=Size(160, 28))
        self.btn_select_all.Click += self.on_select_all
        self.Controls.Add(self.btn_select_all)

        self.btn_clear_all = Button(Text="Clear All", Location=Point(330, 245), Size=Size(160, 28))
        self.btn_clear_all.Click += self.on_clear_all
        self.Controls.Add(self.btn_clear_all)

        self.chk_preview = CheckBox()
        self.chk_preview.Text = "Preview only (do not modify model)"
        self.chk_preview.Location = Point(10, 350)
        self.chk_preview.Size = Size(300, 22)
        self.chk_preview.Checked = False
        self.Controls.Add(self.chk_preview)

        self.btn_ok = Button(Text="Run", Location=Point(330, 300), Size=Size(80, 36))
        self.btn_ok.DialogResult = DialogResult.OK
        self.Controls.Add(self.btn_ok)

        self.btn_cancel = Button(Text="Cancel", Location=Point(410, 300), Size=Size(80, 36))
        self.btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btn_cancel)

        self.AcceptButton = self.btn_ok
        self.CancelButton = self.btn_cancel

    def on_move_up(self, sender, args):
        idx = self.chklist.SelectedIndex
        if idx > 0:
            name = self.chklist.Items[idx]
            checked = self.chklist.GetItemChecked(idx)
            self.chklist.Items.RemoveAt(idx)
            self.chklist.Items.Insert(idx - 1, name)
            self.chklist.SetItemChecked(idx - 1, checked)
            self.chklist.SelectedIndex = idx - 1

    def on_move_down(self, sender, args):
        idx = self.chklist.SelectedIndex
        if idx >= 0 and idx < self.chklist.Items.Count - 1:
            name = self.chklist.Items[idx]
            checked = self.chklist.GetItemChecked(idx)
            self.chklist.Items.RemoveAt(idx)
            self.chklist.Items.Insert(idx + 1, name)
            self.chklist.SetItemChecked(idx + 1, checked)
            self.chklist.SelectedIndex = idx + 1

    def on_select_all(self, sender, args):
        for i in range(self.chklist.Items.Count):
            self.chklist.SetItemChecked(i, True)

    def on_clear_all(self, sender, args):
        for i in range(self.chklist.Items.Count):
            self.chklist.SetItemChecked(i, False)

    def get_ordered_checked(self):
        ordered = []
        for i in range(self.chklist.Items.Count):
            if self.chklist.GetItemChecked(i):
                ordered.append(str(self.chklist.Items[i]))
        return ordered

    def scope_current_view(self):
        return self.rb_current_view.Checked

# --- Main ---
def main():
    uidoc = revit.uidoc
    doc = revit.doc
    view = uidoc.ActiveView

    dlg = HierarchyDialog()
    res = dlg.ShowDialog()
    if res != DialogResult.OK:
        return

    ordered_checked_names = dlg.get_ordered_checked()
    if not ordered_checked_names:
        MessageBox.Show("No categories selected. Operation cancelled.", "Enforce Join Order")
        return

    preview_only = dlg.chk_preview.Checked
    current_view_only = dlg.scope_current_view()

    # Collect elements in selected categories
    elems_by_cat = {}
    for name in ordered_checked_names:
        bic = CAT_MAP[name]
        if current_view_only:
            collector = FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType()
            elems = list(collector)
        else:
            collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
            elems = list(collector)
        elems_by_cat[name] = elems

    summary = []
    total_elems = 0
    for name in ordered_checked_names:
        n = len(elems_by_cat.get(name, []))
        summary.append("{}: {} elements".format(name, n))
        total_elems += n

    # Find already joined pairs
    joined_pairs = []
    for i, higher_name in enumerate(ordered_checked_names):
        higher_elems = elems_by_cat.get(higher_name, [])
        for lower_name in ordered_checked_names[i + 1:]:
            lower_elems = elems_by_cat.get(lower_name, [])
            for he in higher_elems:
                for le in lower_elems:
                    try:
                        if JoinGeometryUtils.AreElementsJoined(doc, he, le):
                            joined_pairs.append((he, le))
                    except Exception:
                        continue

    # Preview
    preview_text_lines = [
        "Enforce Join Order Preview:",
        "View: {}".format(view.Name),
        "Scope: {}".format("Current view only" if current_view_only else "Entire model"),
        "Selected categories (top -> bottom): {}".format(" → ".join(ordered_checked_names)),
        ""
    ]
    preview_text_lines += summary
    preview_text_lines.append("")
    preview_text_lines.append("Already joined pairs found: {}".format(len(joined_pairs)))
    preview_text = "\n".join(preview_text_lines)

    if preview_only:
        MessageBox.Show(preview_text, "Preview - Enforce Join Order")
        return

    if len(joined_pairs) == 0:
        MessageBox.Show(preview_text + "\n\nNo joined pairs found. Nothing to do.", "Enforce Join Order")
        return

    # Apply hierarchy
    switched_count = 0
    errors_count = 0
    trans = Transaction(doc, "Enforce Join Order on Already Joined Elements")
    trans.Start()
    try:
        for he, le in joined_pairs:
            try:
                if not JoinGeometryUtils.IsCuttingElementInJoin(doc, he, le):
                    JoinGeometryUtils.SwitchJoinOrder(doc, he, le)
                    switched_count += 1
            except Exception:
                errors_count += 1
                continue
        trans.Commit()
    except Exception:
        try:
            trans.RollBack()
        except Exception:
            pass
        MessageBox.Show("Transaction failed and was rolled back. See Revit for details.", "Enforce Join Order")
        return

    # Summary
    final_lines = [
        "Enforce Join Order finished.",
        "",
        "Selected categories (top->bottom): {}".format(" → ".join(ordered_checked_names)),
        "Scope: {}".format("Current view only" if current_view_only else "Entire model"),
        "Elements total: {}".format(total_elems),
        "Already joined pairs found: {}".format(len(joined_pairs)),
        "Join orders switched: {}".format(switched_count),
        "Errors encountered: {}".format(errors_count)
    ]
    MessageBox.Show("\n".join(final_lines), "Enforce Join Order")

if __name__ == "__main__":
    main()
