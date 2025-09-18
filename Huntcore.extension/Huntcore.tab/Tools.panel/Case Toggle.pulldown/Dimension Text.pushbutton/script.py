# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms

def convert_case(text, case_type):
    if not text:
        return text
    if case_type == "UPPERCASE":
        return text.upper()
    elif case_type == "lowercase":
        return text.lower()
    elif case_type == "Sentence case":
        return text[:1].upper() + text[1:].lower()
    elif case_type == "Title Case":
        return text.title()
    return text

def update_dimension_text_fields(dimension, case_type):
    updated = False

    # ValueOverride (Replace With Text)
    if dimension.ValueOverride:
        new_val = convert_case(dimension.ValueOverride, case_type)
        if new_val != dimension.ValueOverride:
            dimension.ValueOverride = new_val
            updated = True

    # TextPositioning properties
    text_props = ['Prefix', 'Suffix', 'Above', 'Below']
    for prop_name in text_props:
        prop_info = getattr(dimension, prop_name, None)
        if prop_info:
            new_text = convert_case(prop_info, case_type)
            if new_text != prop_info:
                setattr(dimension, prop_name, new_text)
                updated = True

    return updated

def main():
    views = forms.select_views(title="Select Views to Process", multiple=True)
    if not views:
        forms.alert("No views selected.", title="Dimension Text Case Toggle")
        return

    case_options = ["UPPERCASE", "lowercase", "Sentence case", "Title Case"]
    selected_case = forms.SelectFromList.show(
        case_options,
        title="Select case format",
        multiselect=False
    )

    if not selected_case:
        forms.alert("No case type selected.", title="Dimension Text Case Toggle")
        return

    updated_count = 0
    with revit.Transaction("Convert Dimension Text Case"):
        for view in views:
            collector = DB.FilteredElementCollector(revit.doc, view.Id).OfClass(DB.Dimension)
            for dim in collector:
                try:
                    if update_dimension_text_fields(dim, selected_case):
                        updated_count += 1
                except Exception:
                    # Skip dimension if any issue
                    pass

    forms.alert("Updated {} dimension(s) with {} text formatting.".format(updated_count, selected_case), title="Complete")

if __name__ == "__main__":
    main()
