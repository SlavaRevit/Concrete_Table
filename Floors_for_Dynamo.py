
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


def getting_Volume(floor_list, floor_type_check):
    total_Volume = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Up":
            for type_el in floor_type_check:
                if floor_duplicationTypeMark == type_el:
                    floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                    total_Volume = total_Volume + floor_volume * 0.0283168466
    return total_Volume


def getting_Area(floor_list, floor_type_check):
    total_Area = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Up":
            for type_el in floor_type_check:
                if floor_duplicationTypeMark == type_el:
                    floor_area = floor_element.LookupParameter("Area").AsDouble()
                    total_Area = total_Area + floor_area * 0.092903
    return total_Area


def check_for_zero(result, floor_type_check):
    if result == 0:
        pass
    else:
        for type in floor_type_check:
            result_of_beams_vol = 0
            str_check = ", ".join(floor_type_check)
            # res_of_beams = res_of_beams + result
        if len(floor_type_check) > 1:
            volume = 0
            for p in floor_type_check:
                total_v = getting_Volume(floors_collector, [p])
                volume = volume + total_v
            return result, volume
        elif len(floor_type_check) == 1:
            volume = getting_Volume(floors_collector, [type])
            return result, volume

# Floors = {
# 'regular' : check_for_zero(getting_Area(floors_collector, ["Regular", "Regular-T", "Balcon"]),
#                    ["Regular", "Regular-T", "Balcon"]),
# "Regular-Rrestressed" : check_for_zero(getting_Area(floors_collector, ["Regular-P"]), ["Regular-P"]),
# "Rampa" : check_for_zero(getting_Area(floors_collector, ["Rampa"]), ["Rampa"]),
# "Regular-W-Special" : check_for_zero(getting_Area(floors_collector, ["Regular-W Special"]), ["Regular-W Special"]),
# "Koteret" : check_for_zero(getting_Area(floors_collector, ["Koteret"]), ["Koteret"]),
# "Lite-Beton" : check_for_zero(getting_Area(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
# "Geoplast-D" : check_for_zero(getting_Area(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
# "Slab-Rib" : check_for_zero(getting_Area(floors_collector, ["Slab-Rib Special"]), ["Slab-Rib Special"]),
# "CLSM" : check_for_zero(getting_Area(floors_collector, ["CLSM"]), ["CLSM"]),
# }

