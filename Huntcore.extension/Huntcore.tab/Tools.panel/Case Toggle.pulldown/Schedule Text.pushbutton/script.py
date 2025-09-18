def main():
    # Get schedule views only
    all_schedules = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSchedule).ToElements()
    schedule_names = [s.Name for s in all_schedules]
    schedule_map = dict(zip(schedule_names, all_schedules))

    selected_names = forms.SelectFromList.show(
        schedule_names,
        title="Select Schedules to Process",
        multiselect=True
    )
    if not selected_names:
        forms.alert("No schedules selected.", title="Schedule Text Case Toggle")
        return

    selected_schedules = [schedule_map[name] for name in selected_names]

    case_options = ["UPPERCASE", "lowercase", "Sentence case", "Title Case"]
    selected_case = forms.SelectFromList.show(
        case_options,
        title="Select case format",
        multiselect=False
    )
    if not selected_case:
        forms.alert("No case type selected.", title="Schedule Text Case Toggle")
        return

    total_updated = 0
    with revit.Transaction("Update Schedule Text Case"):
        for schedule in selected_schedules:
            try:
                total_updated += update_schedule_elements(schedule, selected_case)
            except Exception:
                # Some schedules may not allow editing
                pass

    forms.alert("Updated {} cell(s) to {}.".format(total_updated, selected_case), title="Complete")
