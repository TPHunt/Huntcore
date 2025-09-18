# -*- coding: utf-8 -*-
import os, json
from Autodesk.Revit.DB import *
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document


# --- Helpers -----------------------------------------------------

def parse_filter_rules(element_filter):
    """Extract parameter rules in a simplified format"""
    results = []
    try:
        if isinstance(element_filter, ElementParameterFilter):
            rules = element_filter.GetRules()
            for rule in rules:
                if isinstance(rule, FilterStringRule):
                    provider = rule.GetProvider()
                    pid = provider.ParameterId
                    try:
                        bip = System.Enum.ToObject(BuiltInParameter, pid.IntegerValue)
                        param_name = str(bip)
                    except:
                        param_name = str(pid.IntegerValue)

                    evaluator = rule.GetEvaluator()
                    if isinstance(evaluator, FilterStringEquals):
                        cond = "Equals"
                    elif isinstance(evaluator, FilterStringContains):
                        cond = "Contains"
                    else:
                        cond = "Unknown"

                    results.append({
                        "param": param_name,
                        "condition": cond,
                        "value": rule.RuleString
                    })
    except Exception as e:
        print("Could not parse rules: {}".format(e))
    return results


def parse_overrides(view, filter_id):
    """Extract overrides applied in a given view or template"""
    overrides = {}
    ogs = view.GetFilterOverrides(filter_id)

    if ogs.Halftone:
        overrides["halftone"] = True

    if ogs.ProjectionLineColor.IsValid:
        c = ogs.ProjectionLineColor
        overrides["line_color"] = [int(c.Red), int(c.Green), int(c.Blue)]

    if ogs.ProjectionLinePatternId != ElementId.InvalidElementId:
        lpe = doc.GetElement(ogs.ProjectionLinePatternId)
        if lpe:
            overrides["line_pattern"] = lpe.Name

    if ogs.Transparency > 0:
        overrides["surface_transparency"] = int(ogs.Transparency)

    return overrides


def resolve_categories(filter_elem):
    """Resolve filter categories into BuiltInCategory names"""
    cat_names = []
    try:
        for cid in filter_elem.GetCategories():
            # Try BuiltInCategory conversion
            try:
                bic = System.Enum.ToObject(BuiltInCategory, cid.IntegerValue)
                cat_names.append(str(bic))
            except:
                # Fall back to category name
                cat = doc.Settings.Categories.get_Item(cid)
                if cat:
                    cat_names.append(cat.Name)
    except Exception as e:
        print("Could not resolve categories for filter {}: {}".format(filter_elem.Name, e))
    return cat_names


# --- Main --------------------------------------------------------

# Collect all view templates
templates = [v for v in FilteredElementCollector(doc).OfClass(View) if v.IsTemplate]

if not templates:
    forms.alert("No view templates found in this project.", exitscript=True)

# Ask user to select one template to export
selected_template = forms.SelectFromList.show(
    sorted(templates, key=lambda v: v.Name),
    multiselect=False,
    name_attr="Name",
    title="Select View Template to Export"
)

if not selected_template:
    forms.alert("No template selected. Exiting.", exitscript=True)

vt = selected_template

export_filters = []

for fid in vt.GetFilters():
    f_elem = doc.GetElement(fid)

    cats = resolve_categories(f_elem)
    rules = parse_filter_rules(f_elem.GetElementFilter())
    overrides = parse_overrides(vt, fid)

    export_filters.append({
        "name": f_elem.Name,
        "categories": cats,
        "rules": rules,
        "overrides": overrides
    })

# --- Save JSON ---------------------------------------------------

script_dir = os.path.dirname(__file__)
json_path = os.path.join(script_dir, "filters.json")

with open(json_path, "w") as f:
    json.dump(export_filters, f, indent=2)

forms.alert("Exported {} filters to:\n{}".format(len(export_filters), json_path))
