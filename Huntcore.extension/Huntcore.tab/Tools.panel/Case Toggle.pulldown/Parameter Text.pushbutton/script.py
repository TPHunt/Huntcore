# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms

def convert_case(text, case_type):
    if case_type == "UPPERCASE":
        return text.upper()
    elif case_type == "lowercase":
        return text.lower()
    elif case_type == "Sentence case":
        return text[:1].upper() + text[1:].lower() if text else text
    elif case_type == "Title Case":
        return text.title()
    return text

def main():
    selection = revit.get_selection()
    if not selection:
        forms.alert("No elements selected.", title="Parameter Case Toggle")
        return

    param_name = forms.ask_for_string(default="Title on Sheet", prompt="Enter the name of the parameter to convert case:")
    if not param_name:
        forms.alert("No parameter name entered.", title="Parameter Case Toggle")
        return

    case_options = ["UPPERCASE", "lowercase", "Sentence case", "Title Case"]
    selected_case = forms.SelectFromList.show(
    case_options,
    title="Select case format",
    multiselect=False
)

    if not selected_case:
        forms.alert("No case type selected.", title="Parameter Case Toggle")
        return

    updated_count = 0
    with revit.Transaction("Convert Parameter Case"):
        for el in selection:
            param = el.LookupParameter(param_name)
            if param and param.StorageType == DB.StorageType.String and not param.IsReadOnly:
                current_val = param.AsString()
                if current_val:
                    param.Set(convert_case(current_val, selected_case))
                    updated_count += 1

    forms.alert("Updated {} elements' '{}' values to {}.".format(updated_count, param_name, selected_case), title="Complete")

if __name__ == '__main__':
    main()
