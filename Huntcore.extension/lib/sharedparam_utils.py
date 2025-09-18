# sharedparam_utils.py (IronPython-compatible with null checks)
import clr
import os
import System

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager


def get_shared_parameters(doc):
    if doc is None:
        return []

    shared_params = []
    guids_seen = set()

    bindings_map = doc.ParameterBindings
    it = bindings_map.ForwardIterator()
    it.Reset()

    while it.MoveNext():
        definition = it.Key
        binding = it.Current
        # Check if it's a shared parameter
        if isinstance(definition, ExternalDefinition):
            guid = definition.GUID
            if guid not in guids_seen:
                guids_seen.add(guid)
                shared_params.append({
                    "Name": definition.Name,
                    "GUID": str(guid),
                    "Category": "Multiple",  # Could refine by inspecting binding categories
                    "ParameterGroup": definition.ParameterGroup.ToString()
                })

    return shared_params


def export_to_shared_file(param_list):
    uiapp = DocumentManager.Instance.CurrentUIApplication
    if uiapp is None:
        return "UI application context not available. Please run this inside an active Revit session."

    app = uiapp.Application
    sp_file = app.SharedParametersFilename

    if not sp_file or not os.path.exists(sp_file):
        return "Shared parameter file not set or not found."

    shared_param_file = app.OpenSharedParameterFile()
    if not shared_param_file:
        return "Unable to open shared parameter file."

    # Create or find the group
    groups = shared_param_file.Groups
    group = groups.get_Item("Exported Parameters")
    if group is None:
        group = groups.Create("Exported Parameters")

    added = 0
    for param in param_list:
        try:
            def_options = ExternalDefinitionCreationOptions(param["Name"], ParameterType.Text)
            def_options.GUID = System.Guid(param["GUID"])
            definition = group.Definitions.Create(def_options)
            added += 1
        except Exception as e:
            continue

    return "Exported {0} shared parameters to group 'Exported Parameters'.".format(added)
