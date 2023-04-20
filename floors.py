from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")
from System.Collections.Generic import List
doc = __revit__.ActiveUIDocument.Document


floors_collector = FilteredElementCollector(doc).\
    OfCategory(BuiltInCategory.OST_Floors).\
    WhereElementIsNotElementType().\
    ToElementIds()


def getting_Area(floor_list,floor_type_check):
    total_Area = 0.0
    total_Volume = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Up":
            if floor_duplicationTypeMark == floor_type_check:
                floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                floor_area = floor_element.LookupParameter("Area").AsDouble()
                total_Volume = total_Volume + floor_volume * 0.0283168466
                total_Area = total_Area + floor_area * 0.092903
        if floor_type_comments == "Down":
            if floor_duplicationTypeMark == floor_type_check:
                floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                floor_area = floor_element.LookupParameter("Area").AsDouble()
                total_Volume = total_Volume + floor_volume * 0.0283168466
                total_Area = total_Area + floor_area * 0.092903

    return total_Area

def getting_Volume(floor_list,floor_type_check):
    total_Area = 0.0
    total_Volume = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Up":
            if floor_duplicationTypeMark == floor_type_check:
                floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                floor_area = floor_element.LookupParameter("Area").AsDouble()
                total_Volume = total_Volume + floor_volume * 0.0283168466
                total_Area = total_Area + floor_area * 0.092903
        if floor_type_comments == "Down":
            if floor_duplicationTypeMark == floor_type_check:
                floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                floor_area = floor_element.LookupParameter("Area").AsDouble()
                total_Volume = total_Volume + floor_volume * 0.0283168466
                total_Area = total_Area + floor_area * 0.092903

    return total_Volume


def check_for_zero(area,volume, floor_type_check):
    if result == 0:
        pass
    else:
        for type in floor_type_check:
            result_of_beams_vol = 0
            str_check = ", ".join(floor_type_check)
            # res_of_beams = res_of_beams + result
        if len(floor_type_check) > 1:
            total_a_res = 0
            total_v_res = 0
            for p in floor_type_check:
                total_a = getting_Area(floors_collector, p)
                total_v = getting_Volume(floors_collector,p)
                total_a_res =   total_a_res + total_a * 0.092903
                total_v_res = total_v_res + total_v * 0.0283168466
            return total_a_res, total_v_res
        elif len(floor_type_check) == 1:
            volume = getting_Volume(floors_collector, [type])
            area = getting_Area(floors_collector,[type])
        return area


floors_up_area = {
    "regular": check_for_zero(getting_Area_Volume(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                              ["Regular", "Regular-T", "Balcon"]),
    "reg_prestressed": check_for_zero(getting_Area_Volume(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "ramp": check_for_zero(getting_Area_Volume(floors_collector, ["Rampa"]), ["Rampa"]),
    "special": check_for_zero(getting_Area_Volume(floors_collector, ["Regular-W Special"]), ["Regular-W Special"]),
    "koteret": check_for_zero(getting_Area_Volume(floors_collector, ["Koteret"]), ["Koteret"]),
    "lite_beton": check_for_zero(getting_Area_Volume(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "geoplast": check_for_zero(getting_Area_Volume(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "rib_special": check_for_zero(getting_Area_Volume(floors_collector, ["Slab-Rib Special"]), ["Slab-Rib Special"]),
    "CLSM": check_for_zero(getting_Area_Volume(floors_collector, ["CLSM"]), ["CLSM"])
}