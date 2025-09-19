# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
import os, json

from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document


# --- Load JSON config ---
script_dir = os.path.dirname(__file__)
json_path = os.path.join(script_dir, "filters.json")

if os.path.exists(json_path):
    with open(json_path, "r") as f:
        filter_data = json.load(f)
else:
    forms.alert("No filters.json found.\n\nExpected at:\n{}".format(json_path), exitscript=True)


def create_filter(name, categories, rules):
    """Create a filter element with rules and return its Id"""
    cat_ids = List[ElementId]()
    for cat_name in categories:
        bic = getattr(BuiltInCategory, cat_name)
        cat_ids.Add(ElementId(bic))

    # Build rules
    filter_rules = List[FilterRule]()
    for r in rules:
        param_id = getattr(BuiltInParameter, r["param"])
        provider = ParameterValueProvider(ElementId(param_id))

        if r["condition"] == "Equals":
            evaluator = FilterStringEquals()
        elif r["condition"] == "Contains":
            evaluator = FilterStringContains()
        else:
            raise Exception("Unsupported condition: " + r["condition"])

        rule = FilterStringRule(provider, evaluator, r["value"], True)
        filter_rules.Add(rule)

    element_filter = ElementParameterFilter(filter_rules)

    # Create or reuse filter element
    pf = ParameterFilterElement.Create(doc, name, cat_ids, element_filter)
    return pf.Id


def apply_overrides(view, filter_id, overrides):
    """Apply graphic overrides to a view"""
    ogs = OverrideGraphicSettings()

    if "line_color" in overrides:
        r, g, b = overrides["line_color"]
        ogs.SetProjectionLineColor(Color(r, g, b))

    if "surface_transparency" in overrides:
        ogs.SetSurfaceTransparency(overrides["surface_transparency"])

    if "line_pattern" in overrides:
        collector = FilteredElementCollector(doc).OfClass(LinePatternElement)
        lpe = next((x for x in collector if x.Name == overrides["line_pattern"]), None)
        if lpe:
            ogs.SetProjectionLinePatternId(lpe.Id)

    if "halftone" in overrides and overrides["halftone"]:
        ogs.SetHalftone(True)

    view.SetFilterOverrides(filter_id, ogs)


# --- Collect all view templates ---
collector = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()
templates = [v for v in collector if v.IsTemplate]

if not templates:
    forms.alert("No view templates found in this project.", exitscript=True)

# Ask user to select which templates to update
selected_templates = forms.SelectFromList.show(
    sorted(templates, key=lambda v: v.Name),
    multiselect=True,
    name_attr="Name",
    title="Select View Templates to Update"
)

if not selected_templates:
    forms.alert("No templates selected. Exiting.", exitscript=True)

# Ask user to select which filters to add
filter_names = [f["name"] for f in filter_data]
selected_filter_names = forms.SelectFromList.show(
    sorted(filter_names),
    multiselect=True,
    title="Select Filters to Apply"
)

if not selected_filter_names:
    forms.alert("No filters selected. Exiting.", exitscript=True)


# --- Main Transaction ---
t = Transaction(doc, "Batch Create and Apply Filters")
t.Start()

created_filters = {}

# Create only the chosen filters
for f in filter_data:
    if f["name"] in selected_filter_names:
        try:
            fid = create_filter(f["name"], f["categories"], f["rules"])
            created_filters[fid] = f["overrides"]
        except Exception as e:
            print("Failed to create filter {}: {}".format(f["name"], e))

# Apply to the chosen templates
for v in selected_templates:
    for filter_id, overrides in created_filters.items():
        if not v.IsFilterApplied(filter_id):
            v.AddFilter(filter_id)
        apply_overrides(v, filter_id, overrides)
    print("Applied {} filters to template: {}".format(len(created_filters), v.Name))

t.Commit()

print("Done. Applied {} filters to {} templates.".format(len(created_filters), len(selected_templates)))
