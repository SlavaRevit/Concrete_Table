from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document

floors_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_Floors). \
    WhereElementIsNotElementType(). \
    ToElementIds()

floors_up = {}
floors_down = {}


def getting_floors_parameters(floor_list):
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Up":
            if floor_duplicationTypeMark in ["Total Floor Area", "Total Floor Area Commercial","Total Floor Area LSP","Total Floor Area Pergola",
                "Air Double Level","Air Elevator","Air Pergola Aluminium","Air Pergola Steel","Air Pergola Wood","Air Regular","Air Stairs","Landing-H","Landing-S","Landing Steel","Polivid","Backfilling"]:
                continue
            elif floor_duplicationTypeMark in ["Regular", "Balcon", "Regular-T"]:
                combined_key = "Regular_new"
                if combined_key not in floors_up:
                    floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                    floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                    floors_up[combined_key] = {"Area": floor_area, "Volume": floor_volume}
                else:
                    floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                    floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                    floors_up[combined_key]['Area'] += floor_area
                    floors_up[combined_key]['Volume'] += floor_volume
            elif floor_duplicationTypeMark not in floors_up and floor_duplicationTypeMark not in combined_key:
                floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                floors_up[floor_duplicationTypeMark] = {"Area": floor_area, "Volume": floor_volume}

            else:
                floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                floors_up[floor_duplicationTypeMark]["Area"] += floor_area
                floors_up[floor_duplicationTypeMark]["Volume"] += floor_volume

        elif floor_type_comments == "Down":
            if floor_duplicationTypeMark in ["Total Floor Area", "Total Floor Area Commercial","Total Floor Area LSP","Total Floor Area Pergola",
                "Air Double Level","Air Elevator","Air Pergola Aluminium","Air Pergola Steel","Air Pergola Wood","Air Regular","Air Stairs","Landing-H","Landing-S","Landing Steel","Backfilling"]:
                continue
            elif floor_duplicationTypeMark in ["Regular", "Balcon", "Regular-T"]:
                combined_key = "Regular_new_down"
                if combined_key not in floors_down:
                    floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                    floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                    floors_down[combined_key] = {"Area": floor_area, "Volume": floor_volume}
                else:
                    floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                    floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                    floors_down[combined_key]['Area'] += floor_area
                    floors_down[combined_key]['Volume'] += floor_volume
            elif floor_duplicationTypeMark not in floors_down and floor_duplicationTypeMark not in combined_key:
                floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                floors_down[floor_duplicationTypeMark] = {"Area": floor_area, "Volume": floor_volume}

            else:
                floor_area = floor_element.LookupParameter("Area").AsDouble() * 0.092903
                floor_volume = floor_element.LookupParameter("Volume").AsDouble() * 0.0283168466
                floors_down[floor_duplicationTypeMark]["Area"] += floor_area
                floors_down[floor_duplicationTypeMark]["Volume"] += floor_volume

    return floors_up, floors_down


getting_floors_parameters(floors_collector)
"""________FLOORS________"""
