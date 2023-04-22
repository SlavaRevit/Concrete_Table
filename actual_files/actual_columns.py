from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document

columns_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_StructuralColumns). \
    WhereElementIsNotElementType(). \
    ToElements()

columns = {}


def getiing_parameters(columns_collector):
    for el in columns_collector:
        col_type_id = el.GetTypeId()
        col_type_elem = doc.GetElement(col_type_id)
        if col_type_elem:
            parameter_Duplication = col_type_elem.LookupParameter("Duplication Type Mark").AsString()
            if parameter_Duplication in ["Rec", "Round", "Eliptic"]:
                parameter_vol = el.LookupParameter("Volume")
                key = "Columns"
                if key not in columns:
                    parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                    columns[key] = {"Volume": parameter_value_vol}
                elif key in columns:
                    parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                    columns[key]["Volume"] += parameter_value_vol

            elif parameter_Duplication == "Precast":
                if parameter_Duplication not in columns:
                    parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                    columns[parameter_Duplication] = {"Volume": parameter_value_vol, "Count": 1}
                elif parameter_Duplication in columns:
                    parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                    columns[parameter_Duplication]["Volume"] += parameter_value_vol
                    columns[parameter_Duplication]["Count"] += 1
            else:
                pass

getiing_parameters(columns_collector)