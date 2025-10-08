# -*- coding: utf-8 -*-
"""
Add predefined filters to selected view templates to control Detail Component visibility.
Multiple values can be added to the Comments parameter.

Detail Component - Halftone - Comments contains Halftone
Detail Component - Dashed - Comments contains Dashed
Detail Component - Hidden - Comments contains Hidden
Detail Component - Lineweight 1 - Comments contains LW1
Detail Component - Lineweight 2 - Comments contains LW2
Detail Component - Lineweight 3 - Comments contains LW3
Detail Component - Lineweight 4 - Comments contains LW4
Detail Component - Lineweight 5 - Comments contains LW5
Detail Component - Red - Comments contains Red
Detail Component - Blue - Comments contains Blue
Detail Component - Demo - Comments contains Demo
Detail Component - Existing - Comments contains Existing

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from System import Guid as SysGuid
import os, json, re
from pyrevit import forms, revit

doc = revit.doc
uidoc = revit.uidoc

# ---------------------- Helpers ----------------------
COLOR_NAME_MAP = {
    "red": (255, 0, 0),
    "blue": (0, 0, 255),
    "green": (0, 128, 0),
    "black": (0, 0, 0),
    "white": (255, 255, 255),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "yellow": (255, 255, 0),
    "cyan": (0, 255, 255),
    "magenta": (255, 0, 255),
}

def parse_color(value):
    if value is None:
        return None
    if isinstance(value, (list, tuple)):
        try:
            r, g, b = int(value[0]), int(value[1]), int(value[2])
            return Color(r, g, b)
        except Exception:
            return None
    text = str(value).strip()
    m = re.match(r"^#?([0-9A-Fa-f]{6})$", text)
    if m:
        hexs = m.group(1)
        r = int(hexs[0:2], 16)
        g = int(hexs[2:4], 16)
        b = int(hexs[4:6], 16)
        return Color(r, g, b)
    nums = re.findall(r"\d{1,3}", text)
    if len(nums) >= 3:
        r, g, b = int(nums[0]), int(nums[1]), int(nums[2])
        return Color(r, g, b)
    key = text.lower()
    if key in COLOR_NAME_MAP:
        r, g, b = COLOR_NAME_MAP[key]
        return Color(r, g, b)
    return None

def find_line_pattern(name):
    if not name:
        return None
    for lp in FilteredElementCollector(doc).OfClass(LinePatternElement):
        try:
            if lp.Name == name:
                return lp
        except Exception:
            continue
    return None

def find_fill_pattern(name):
    if not name:
        return None
    try:
        fpe = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, name)
        if fpe:
            return fpe
    except Exception:
        pass
    try:
        fpe = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Model, name)
        if fpe:
            return fpe
    except Exception:
        pass
    for p in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            if p.Name == name:
                return p
            fp = p.GetFillPattern()
            if fp and fp.Name == name:
                return p
        except Exception:
            continue
    return None

def resolve_category_id(cat_token):
    try:
        bic = getattr(BuiltInCategory, cat_token)
        cat = Category.GetCategory(doc, bic)
        if cat is not None:
            return cat.Id
        return ElementId(bic)
    except Exception:
        raise Exception("Unknown BuiltInCategory: {0}".format(cat_token))

def resolve_parameter_elementid(param_token):
    try:
        bip = getattr(BuiltInParameter, param_token)
        return ElementId(bip)
    except Exception:
        pass
    try:
        g = SysGuid(param_token)
        spe = SharedParameterElement.Lookup(doc, g)
        if spe:
            return spe.Id
    except Exception:
        pass
    try:
        for pe in FilteredElementCollector(doc).OfClass(ParameterElement):
            try:
                if pe.GetDefinition().Name == param_token:
                    return pe.Id
            except Exception:
                continue
    except Exception:
        pass
    raise Exception("Parameter token not found: {0}".format(param_token))

def find_existing_filter_by_name(name):
    for pfe in FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements():
        try:
            if pfe.Name == name:
                return pfe
        except Exception:
            continue
    return None

# ---------------------- Create / Update filter ---------------------------
def create_filter(name, categories, rules):
    existing = find_existing_filter_by_name(name)
    if existing:
        return existing.Id

    cat_ids = List[ElementId]()
    for cat_name in categories:
        cat_ids.Add(resolve_category_id(cat_name))

    filter_rules = List[FilterRule]()
    for r in rules:
        pid = resolve_parameter_elementid(r["param"])
        provider = ParameterValueProvider(pid)
        cond = r.get("condition", "Contains").lower()
        if cond == "equals":
            evaluator = FilterStringEquals()
        elif cond == "contains":
            evaluator = FilterStringContains()
        else:
            raise Exception("Unsupported string condition: {0}".format(r.get("condition")))
        rule = FilterStringRule(provider, evaluator, r["value"])
        filter_rules.Add(rule)

    element_filter = ElementParameterFilter(filter_rules)

    try:
        pf = ParameterFilterElement.Create(doc, name, cat_ids, element_filter)
    except Exception:
        pf = ParameterFilterElement.Create(doc, name, cat_ids)
        pf.SetElementFilter(element_filter)
    return pf.Id

# ---------------------- Apply overrides -----------------------------------
def apply_overrides(view, filter_id, overrides):
    ogs = OverrideGraphicSettings()

    col = parse_color(overrides.get("line_color"))
    if col:
        ogs.SetProjectionLineColor(col)

    lw = overrides.get("line_weight")
    if lw is not None:
        try:
            ogs.SetProjectionLineWeight(int(lw))
        except Exception:
            pass

    lp = find_line_pattern(overrides.get("line_pattern"))
    if lp:
        ogs.SetProjectionLinePatternId(lp.Id)

    sp = find_fill_pattern(overrides.get("surface_pattern"))
    if sp:
        try:
            ogs.SetSurfaceForegroundPatternId(sp.Id)
        except Exception:
            pass

    spc = parse_color(overrides.get("surface_pattern_color"))
    if spc:
        try:
            ogs.SetSurfaceForegroundPatternColor(spc)
        except Exception:
            pass

    cp = find_fill_pattern(overrides.get("cut_pattern"))
    if cp:
        try:
            ogs.SetCutForegroundPatternId(cp.Id)
        except Exception:
            pass

    cpc = parse_color(overrides.get("cut_pattern_color"))
    if cpc:
        try:
            ogs.SetCutForegroundPatternColor(cpc)
        except Exception:
            pass

    tr = overrides.get("surface_transparency")
    if tr is not None:
        try:
            ogs.SetSurfaceTransparency(int(tr))
        except Exception:
            pass

    ht = bool(overrides.get("halftone", False))
    try:
        ogs.SetHalftone(ht)
    except Exception:
        pass

    try:
        if not view.IsFilterApplied(filter_id):
            view.AddFilter(filter_id)
        view.SetFilterOverrides(filter_id, ogs)
    except Exception as e:
        print("Failed to apply overrides for filter id {0} on view {1}: {2}".format(filter_id, view.Name, e))

# ---------------------- Main -----------------------------------------------
script_dir = os.path.dirname(__file__)
json_path = os.path.join(script_dir, "filters.json")
if not os.path.exists(json_path):
    forms.alert("No filters.json found.\n\nExpected at:\n{0}".format(json_path), exitscript=True)

try:
    with open(json_path, "r") as f:
        filter_data = json.load(f)
except Exception as e:
    forms.alert("Failed to read JSON: {0}\nFix JSON (no trailing commas).".format(e), exitscript=True)

templates = [v for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType() if getattr(v, "IsTemplate", False)]
if not templates:
    forms.alert("No view templates found in this project.", exitscript=True)

selected_templates = forms.SelectFromList.show(
    sorted(templates, key=lambda v: v.Name),
    multiselect=True,
    name_attr="Name",
    title="Select View Templates to Update"
)
if not selected_templates:
    forms.alert("No templates selected. Exiting.", exitscript=True)

filter_names = [ff.get("name") for ff in filter_data if ff.get("name")]
selected_filter_names = forms.SelectFromList.show(
    sorted(filter_names),
    multiselect=True,
    title="Select Filters to Apply"
)
if not selected_filter_names:
    forms.alert("No filters selected. Exiting.", exitscript=True)

tx = Transaction(doc, "Batch Create/Update and Apply View Filters")
tx.Start()
created_filters = {}

for ff in filter_data:
    fname = ff.get("name")
    if fname not in selected_filter_names:
        continue
    try:
        fid = create_filter(fname, ff.get("categories", []), ff.get("rules", []))
        created_filters[fid] = ff.get("overrides", {})
    except Exception as e:
        print("Failed to create/check filter '{0}': {1}".format(fname, e))

for v in selected_templates:
    for filter_id, overrides in created_filters.items():
        apply_overrides(v, filter_id, overrides)

tx.Commit()
forms.alert(
    "Done. Applied {0} filter(s) to {1} template(s).".format(len(created_filters), len(selected_templates)),
    title="Complete"
)
