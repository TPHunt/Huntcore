# -*- coding: utf-8 -*-
"""
Create Structural Plan Views from Selected View Templates and Levels.
Select scope box for primary views, then multiple scope boxes for dependents if required.

View names created using the names of the view templates, levels and scope boxes.

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Integer ViewType enum values for compatibility
# FloorPlan = 1, CeilingPlan = 2, StructuralPlan = 11
plan_view_type_ids = [1, 2, 11]

# Collect only plan-type view templates
all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
view_templates = [v for v in all_views if v.IsTemplate and int(v.ViewType) in plan_view_type_ids]

if not view_templates:
    forms.alert('No plan-type view templates found in this project.', exitscript=True)

# Get levels
levels = list(FilteredElementCollector(doc).OfClass(Level))
if not levels:
    forms.alert('No levels found.', exitscript=True)

# Get scope boxes
scope_boxes = list(FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                   .WhereElementIsNotElementType()
                   .ToElements())
if not scope_boxes:
    forms.alert('No scope boxes found in the project.', exitscript=True)

# User selects templates
selected_templates = forms.SelectFromList.show(
    [v.Name for v in view_templates],
    title='Select Plan View Templates',
    button_name='Select',
    multiselect=True
)
if not selected_templates:
    forms.alert('No templates selected.', exitscript=True)

# Sort levels so lowest elevation is first (top of list), highest last (bottom of list)
levels_sorted = sorted(levels, key=lambda l: l.Elevation, reverse=False)

# User selects levels
selected_levels = forms.SelectFromList.show(
    [l.Name for l in levels_sorted],
    title='Select Levels',
    button_name='Select',
    multiselect=True
)
if not selected_levels:
    forms.alert('No levels selected.', exitscript=True)

# User selects scope box for primary views
selected_scope_primary = forms.SelectFromList.show(
    [sb.Name for sb in scope_boxes],
    title='Select Scope Box for Primary Views',
    button_name='Select'
)
if not selected_scope_primary:
    forms.alert('No primary scope box selected.', exitscript=True)

# User selects additional scope boxes for dependent views
selected_scope_dependents = forms.SelectFromList.show(
    [sb.Name for sb in scope_boxes if sb.Name != selected_scope_primary],
    title='Select Additional Scope Boxes for Dependent Views',
    button_name='Select',
    multiselect=True
)

# Helper to clean view template name
def clean_template_name(vt_name):
    for i, ch in enumerate(vt_name):
        if ch.isalpha():
            return vt_name[i:].strip()
    return vt_name.strip()

# Helper: Create Dependent View with scope box replacing primary
def create_dependent_view(source_view, dep_scope_name):
    try:
        dup_view_id = source_view.Duplicate(ViewDuplicateOption.AsDependent)
        dup_view = doc.GetElement(dup_view_id)

        # Apply dependent scope box
        scope = next((sb for sb in scope_boxes if sb.Name == dep_scope_name), None)
        if scope:
            param = dup_view.LookupParameter('Scope Box')
            if param and not param.IsReadOnly:
                param.Set(scope.Id)

        # Replace primary scope box suffix with dependent scope box
        name_parts = source_view.Name.rsplit(' - ', 1)
        base_name = name_parts[0] if len(name_parts) > 1 else source_view.Name
        new_name = base_name + ' - ' + dep_scope_name

        # Ensure unique name
        existing_names = [v.Name for v in FilteredElementCollector(doc).OfClass(View)]
        count = 1
        temp_name = new_name
        while new_name in existing_names:
            new_name = temp_name + ' ({})'.format(count)
            count += 1

        dup_view.Name = new_name
        return dup_view

    except Exception as ex:
        print('Failed to create dependent view for {}: {}'.format(dep_scope_name, ex))
        return None

# Start transaction
created_views = []
with revit.Transaction('Create Structural Plans with Dependents'):
    for vt_name in selected_templates:
        vt = next(v for v in view_templates if v.Name == vt_name)
        vt_clean = clean_template_name(vt_name)

        for lvl_name in selected_levels:
            lvl = next(l for l in levels if l.Name == lvl_name)

            # Choose ViewFamilyType (prefer StructuralPlan, else FloorPlan)
            vft_candidates = [vft for vft in FilteredElementCollector(doc)
                              .OfClass(ViewFamilyType)
                              if vft.ViewFamily in [ViewFamily.StructuralPlan, ViewFamily.FloorPlan]]
            if not vft_candidates:
                forms.alert('No suitable ViewFamilyType found for plan views.', exitscript=True)
            view_family_type = vft_candidates[0]

            new_view = ViewPlan.Create(doc, view_family_type.Id, lvl.Id)

            # Apply template
            new_view.ViewTemplateId = vt.Id

            # Apply scope box for primary view
            primary_scope = next((sb for sb in scope_boxes if sb.Name == selected_scope_primary), None)
            if primary_scope:
                param = new_view.LookupParameter('Scope Box')
                if param and not param.IsReadOnly:
                    param.Set(primary_scope.Id)

            # Construct view name: "<Template> - Level <LevelName> <Remainder> - <ScopeBox>"
            parts = vt_clean.split('-', 1)
            before = parts[0].strip() if parts else vt_clean
            after = parts[1].strip() if len(parts) > 1 else ''
            new_name = ("{0} - Level {1} {2} - {3}".format(before, lvl_name, after, selected_scope_primary)).strip()

            # Ensure unique name
            existing_names = [v.Name for v in FilteredElementCollector(doc).OfClass(View)]
            count = 1
            base_name = new_name
            while new_name in existing_names:
                new_name = base_name + ' ({})'.format(count)
                count += 1

            # Rename view
            try:
                new_view.Name = new_name
                created_views.append(new_name)
            except Exception as e:
                print('Failed to rename view for {}: {}'.format(lvl_name, e))

            # Create dependent views for each selected scope box
            if selected_scope_dependents:
                for dep_scope in selected_scope_dependents:
                    dep_view = create_dependent_view(new_view, dep_scope)
                    if dep_view:
                        created_views.append(dep_view.Name)

# Summary
TaskDialog.Show('Completed', '{} plan and dependent views created.'.format(len(created_views)))
