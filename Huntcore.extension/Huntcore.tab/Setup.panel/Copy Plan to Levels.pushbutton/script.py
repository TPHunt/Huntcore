# -*- coding: utf-8 -*-
"""
Copy active plan view to selected levels
- User supplies a naming template using <level> placeholder
- Copies scope box if present, otherwise copies crop region (shape or CropBox)
- Copies annotation-crop active state & offsets when possible
- Applies the active view's view template to each new view

Compatible with Revit 2023+ and pyRevit (IronPython)
"""

import clr
clr.AddReference("Microsoft.VisualBasic")
from Microsoft.VisualBasic import Interaction

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

# -------------------------
# Helpers / safety
# -------------------------
def input_box(prompt, title, default=""):
    """Simple InputBox wrapper (VB Interaction) that returns None on cancel/empty."""
    try:
        # Interaction.InputBox returns "" when cancelled or empty; treat empty as cancel
        res = Interaction.InputBox(prompt, title, default)
        if res is None:
            return None
        # Trim whitespace; if empty string treat as cancelled
        res = res.strip()
        if res == "":
            return None
        return res
    except Exception:
        return None

# -------------------------
# Doc / Active View
# -------------------------
doc = revit.doc
uidoc = revit.uidoc
active_view = revit.active_view

# Ensure active view is a plan view
if not isinstance(active_view, ViewPlan):
    forms.alert("Active view must be a plan (floor/structural) view.", title="Wrong View")
    script.exit()

# -------------------------
# Get naming template from user
# -------------------------
name_template = input_box(
    "Enter naming template for new views.\nUse <level> to insert the level name.\nExample: <level> GA Plan",
    "New View Name Template",
    "<level> GA Plan"
)

if not name_template:
    forms.alert("No naming template entered. Script cancelled.", title="Cancelled")
    script.exit()

# -------------------------
# Collect levels and let user pick
# -------------------------
all_levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
if not all_levels:
    forms.alert("No levels found in the project.", title="No Levels")
    script.exit()

# SelectFromList requires context, title, width, height as positional args
# We'll pass name_attr so items display by Name
selected_levels = forms.SelectFromList.show(
    all_levels,                       # context (list of Level objects)
    "Select Levels",                  # title
    500,                              # width
    600,                              # height
    multiselect=True,
    name_attr="Name"
)

if not selected_levels:
    # user cancelled or selected none
    script.exit()

# -------------------------
# Capture source view properties: view template, scope box, crop shape, annotation crop
# -------------------------
# Active view's template
active_view_template_id = active_view.ViewTemplateId

# Scope box (if any)
scope_box_param = active_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
scope_box_id = None
if scope_box_param and scope_box_param.AsElementId() != ElementId.InvalidElementId:
    scope_box_id = scope_box_param.AsElementId()

# If no scope box: capture crop shape + annotation crop state and offsets
source_crop_shape = None
source_crm = None
source_annotation_active = False
annotation_offsets = {}

if not scope_box_id:
    try:
        source_crm = active_view.GetCropRegionShapeManager()
        if source_crm:
            crop_shape_raw = source_crm.GetCropShape()
            # Try to coerce to list of loops (common return types vary)
            try:
                source_crop_shape = list(crop_shape_raw)
            except Exception:
                # single loop or non-iterable
                if crop_shape_raw is not None:
                    source_crop_shape = [crop_shape_raw]
                else:
                    source_crop_shape = None

            # annotation crop active?
            ann_param = active_view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
            if ann_param and ann_param.StorageType == StorageType.Integer:
                source_annotation_active = bool(ann_param.AsInteger())

            # copy annotation offsets if supported
            try:
                if source_crm and source_crm.CanHaveAnnotationCrop:
                    annotation_offsets = {
                        "Left": source_crm.LeftAnnotationCropOffset,
                        "Right": source_crm.RightAnnotationCropOffset,
                        "Top": source_crm.TopAnnotationCropOffset,
                        "Bottom": source_crm.BottomAnnotationCropOffset
                    }
            except Exception:
                annotation_offsets = {}
    except Exception:
        source_crm = None
        source_crop_shape = None

# -------------------------
# Find a Plan ViewFamilyType if needed (we will reuse active view's type id)
# -------------------------
# We will reuse the active view's type id instead of searching types:
plan_type_id = active_view.GetTypeId()

created_views = []

# -------------------------
# Create views in transaction
# -------------------------
t = Transaction(doc, "Copy Plan to Levels")
t.Start()

for lvl in selected_levels:
    try:
        # Create the new view using the same view type as active view
        new_view = ViewPlan.Create(doc, plan_type_id, lvl.Id)

        # Name using template
        try:
            new_name = name_template.replace("<level>", lvl.Name)
            new_view.Name = new_name
        except Exception:
            # fallback name
            try:
                new_view.Name = "Plan - {}".format(lvl.Name)
            except Exception:
                pass

        # Apply the same view template as the active view (if valid)
        try:
            if active_view_template_id and active_view_template_id != ElementId.InvalidElementId:
                new_view.ViewTemplateId = active_view_template_id
        except Exception:
            # ignore if not applicable
            pass

        # Apply scope box if source has one
        if scope_box_id:
            try:
                sb_param = new_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                if sb_param and sb_param.StorageType == StorageType.ElementId:
                    sb_param.Set(scope_box_id)
            except Exception:
                pass

        # Otherwise attempt to copy crop shape
        elif source_crop_shape:
            try:
                # Make crop visible & active on the new view first
                try:
                    new_view.CropBoxVisible = True
                except Exception:
                    pass
                try:
                    new_view.CropBoxActive = True
                except Exception:
                    pass

                # Regenerate to ensure shape manager is ready
                try:
                    doc.Regenerate()
                except Exception:
                    pass

                new_crm = None
                try:
                    new_crm = new_view.GetCropRegionShapeManager()
                except Exception:
                    new_crm = None

                # If the new view supports custom shape, try to set the first loop
                if new_crm and new_crm.CanHaveShape:
                    loop_to_set = None
                    try:
                        if isinstance(source_crop_shape, list) and len(source_crop_shape) > 0:
                            loop_to_set = source_crop_shape[0]
                        else:
                            loop_to_set = source_crop_shape
                    except Exception:
                        loop_to_set = source_crop_shape

                    if loop_to_set is not None:
                        try:
                            new_crm.SetCropShape(loop_to_set)
                        except Exception:
                            # fallback to copying rectangular CropBox
                            try:
                                new_view.CropBox = active_view.CropBox
                                new_view.CropBoxActive = True
                                new_view.CropBoxVisible = True
                            except Exception:
                                pass
                else:
                    # fallback: copy rectangular CropBox
                    try:
                        new_view.CropBox = active_view.CropBox
                        new_view.CropBoxActive = True
                        new_view.CropBoxVisible = True
                    except Exception:
                        pass

                # Copy annotation crop active parameter if applicable
                try:
                    ann_param_new = new_view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
                    if ann_param_new and ann_param_new.StorageType == StorageType.Integer and source_annotation_active:
                        ann_param_new.Set(1)
                except Exception:
                    pass

                # Copy annotation offsets if supported on new view
                try:
                    if new_crm and new_crm.CanHaveAnnotationCrop and annotation_offsets:
                        try:
                            new_crm.LeftAnnotationCropOffset = annotation_offsets.get("Left", new_crm.LeftAnnotationCropOffset)
                            new_crm.RightAnnotationCropOffset = annotation_offsets.get("Right", new_crm.RightAnnotationCropOffset)
                            new_crm.TopAnnotationCropOffset = annotation_offsets.get("Top", new_crm.TopAnnotationCropOffset)
                            new_crm.BottomAnnotationCropOffset = annotation_offsets.get("Bottom", new_crm.BottomAnnotationCropOffset)
                        except Exception:
                            pass
                except Exception:
                    pass

                # Final regenerate
                try:
                    doc.Regenerate()
                except Exception:
                    pass

            except Exception:
                # any crop-copy problems should not stop creation of other views
                pass

        created_views.append(new_view.Name)

    except Exception as ex:
        print("Failed creating view for level '{}': {}".format(getattr(lvl, "Name", "Unknown"), ex))

# Commit transaction
try:
    t.Commit()
except Exception:
    try:
        t.RollBack()
    except Exception:
        pass

# -------------------------
# Report results
# -------------------------
if created_views:
    forms.alert("Created {} views:\n\n{}".format(len(created_views), "\n".join(created_views)), title="Copy Plan to Levels")
else:
    forms.alert("No views were created.", title="Copy Plan to Levels")
