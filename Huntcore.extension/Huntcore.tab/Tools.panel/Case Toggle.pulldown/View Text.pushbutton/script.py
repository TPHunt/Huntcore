# -*- coding: utf-8 -*-

"""
Update text case of specified views. 

!!! NOTE: Formatting lost on bulletpoint text notes!!!

Choice of:
UPPERCASE
lowercase
Sentence case
Title case

Huntcore Script | Author: Troy Hunt | Revit 2023+
"""

from pyrevit import revit, DB, forms
import re

def convert_case(text, case_type):
    # Pattern for typical engineering units and AS references (preserve case)
    preserve_pattern = re.compile(
        r'(\b\d+(\.\d+)?\s*(mm|kPa|m/s|MPa|kN|g/m|m)\b'
        r'|\b(UNO|RHS|SHS|CHS|UB|UC|EA|UA|TF|/S|/TB)\b'
        r'|\bAS\s*/?NZL?\s*\d+\b'
        r'|\bAS\s*\d+\b)',
        re.IGNORECASE
    )

    # Find and store unit matches
    preserved_units = {}
    def preserve(match):
        key = "__PRESERVE_{}__".format(len(preserved_units))
        preserved_units[key] = match.group(0)
        return key

    temp_text = preserve_pattern.sub(preserve, text)

    # Apply case conversion
    if case_type == "UPPERCASE":
        temp_text = temp_text.upper()
    elif case_type == "lowercase":
        temp_text = temp_text.lower()
    elif case_type == "Sentence case":
        temp_text = temp_text[:1].upper() + temp_text[1:].lower() if temp_text else temp_text
    elif case_type == "Title Case":
        temp_text = temp_text.title()

    # Restore preserved units
    for key, val in preserved_units.items():
        temp_text = temp_text.replace(key, val)

    return temp_text

def main():
    views = forms.select_views(
        title="Select Views to Process",
        multiple=True
    )
    if not views:
        forms.alert("No views selected.", title="Text Case Converter")
        return

    case_options = ["UPPERCASE", "lowercase", "Sentence case", "Title Case"]
    selected_case = forms.SelectFromList.show(
        case_options,
        title="Select case format",
        multiselect=False
    )

    if not selected_case:
        return

    updated_count = 0
    with revit.Transaction("Convert Text Note Case"):
        for view in views:
            collector = DB.FilteredElementCollector(revit.doc, view.Id).OfClass(DB.TextNote)
            for text_note in collector:
                try:
                    current_text = text_note.Text
                    if current_text:
                        new_text = convert_case(current_text, selected_case)
                        if current_text != new_text:
                            text_note.Text = new_text
                            updated_count += 1
                except Exception as e:
                    print("Skipping a note due to error: {}".format(e))

    forms.alert("Updated {} text note(s) to {}.".format(updated_count, selected_case), title="Complete")

if __name__ == "__main__":
    main()
