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


def getting_Area_up(floor_list, floor_type_check):
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


def getting_Area_down(floor_list, floor_type_check):
    total_Area = 0.0

    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Down":
            for type_el in floor_type_check:
                if floor_duplicationTypeMark == type_el:
                    floor_area = floor_element.LookupParameter("Area").AsDouble()
                    total_Area = total_Area + floor_area * 0.092903
    return round(total_Area, 2)


def getting_Volume_up(floor_list, floor_type_check):
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


def getting_Volume_down(floor_list, floor_type_check):
    total_Volume = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Down":
            for type_el in floor_type_check:
                if floor_duplicationTypeMark == type_el:
                    floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                    total_Volume = total_Volume + floor_volume * 0.0283168466
    return total_Volume


def check_Area_for_zero(result, floor_type_check):
    if result == 0:
        pass
    else:
        for type in floor_type_check:
            result_of_beams_vol = 0
            str_check = ", ".join(floor_type_check)
            # res_of_beams = res_of_beams + result
        if len(floor_type_check) > 1:
            area = 0
            for p in floor_type_check:
                total_a = getting_Area_up(floors_collector, [p])
                area = area + total_a
            return area
        elif len(floor_type_check) == 1:
            area = getting_Area_up(floors_collector, [type])
            return area


def check_Area_for_zero_down(result, floor_type_check):
    if result == 0:
        pass
    else:
        for type in floor_type_check:
            result_of_beams_vol = 0
            str_check = ", ".join(floor_type_check)
            # res_of_beams = res_of_beams + result
        if len(floor_type_check) > 1:
            area = 0
            for p in floor_type_check:
                total_a = getting_Area_down(floors_collector, [p])
                area = area + total_a
            return area
        elif len(floor_type_check) == 1:
            area = getting_Area_down(floors_collector, [type])
            return area


def check_Volume_for_zero(result, floor_type_check):
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
                total_a = getting_Volume_up(floors_collector, [p])
                volume = volume + total_a
            return volume
        elif len(floor_type_check) == 1:
            volume = getting_Volume_up(floors_collector, [type])
            return volume


def check_Volume_for_zero_down(result, floor_type_check):
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
                total_a = getting_Volume_down(floors_collector, [p])
                volume = volume + total_a
            return volume
        elif len(floor_type_check) == 1:
            volume = getting_Volume_down(floors_collector, [type])
            return volume


"""________FLOORS________"""

floors_up_Area = {
    "Up_Regular": check_Area_for_zero(getting_Area_up(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                      ["Regular", "Regular-T", "Balcon"]),
    "Up_Prestressed": check_Area_for_zero(getting_Area_up(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Up_Rampa": check_Area_for_zero(getting_Area_up(floors_collector, ["Rampa"]), ["Rampa"]),
    "Regular-W Special": check_Area_for_zero(getting_Area_up(floors_collector, ["Regular-W Special"]),
                                             ["Regular-W Special"]),
    "Koteret": check_Area_for_zero(getting_Area_up(floors_collector, ["Koteret"]), ["Koteret"]),
    "Lite Beton": check_Area_for_zero(getting_Area_up(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Geoplast-D": check_Area_for_zero(getting_Area_up(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Slab-Rib Special": check_Area_for_zero(getting_Area_up(floors_collector, ["Slab-Rib Special"]),
                                            ["Slab-Rib Special"]),
    "CLSM": check_Area_for_zero(getting_Area_up(floors_collector, ["CLSM"]), ["CLSM"]),
    "Topping": check_Area_for_zero(getting_Area_up(floors_collector, ["Topping"]), ["Topping"]),
    "Slab": check_Area_for_zero(getting_Area_up(floors_collector, ["Slab"]), ["Slab"]),
    "Up_complition": check_Area_for_zero(getting_Area_up(floors_collector, ["Completion"]), ["Completion"])
}
floors_down_Area = {
    "Dn_Regular": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                           ["Regular", "Regular-T", "Balcon"]),
    "Dn_Prestressed": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Dn_Rampa": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Rampa"]), ["Rampa"]),
    "Dn_Regular-W Special": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Regular-W Special"]),
                                                     ["Regular-W Special"]),
    "Dn_Koteret": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Koteret"]), ["Koteret"]),
    "Dn_Lite Beton": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Dn_Geoplast-D": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Dn_Slab-Rib Special": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Slab-Rib Special"]),
                                                    ["Slab-Rib Special"]),
    "CLSM": check_Area_for_zero_down(getting_Area_down(floors_collector, ["CLSM"]), ["CLSM"]),
    "Topping": check_Area_for_zero(getting_Area_up(floors_collector, ["Topping"]), ["Topping"]),
    "Slab": check_Area_for_zero(getting_Area_up(floors_collector, ["Slab"]), ["Slab"]),
    "Dn_complition": check_Area_for_zero(getting_Area_up(floors_collector, ["Completion"]), ["Completion"])
}
floors_up_Volume = {
    "Up_Regular": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                        ["Regular", "Regular-T", "Balcon"]),
    "Up_Prestressed": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Up_Rampa": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Rampa"]), ["Rampa"]),
    "Regular-W Special": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Regular-W Special"]),
                                               ["Regular-W Special"]),
    "Koteret": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Koteret"]), ["Koteret"]),
    "Lite Beton": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Geoplast-D": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Slab-Rib Special": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Slab-Rib Special"]),
                                              ["Slab-Rib Special"]),
    "CLSM": check_Volume_for_zero(getting_Volume_up(floors_collector, ["CLSM"]), ["CLSM"]),
    "Topping": check_Area_for_zero(getting_Area_up(floors_collector, ["Topping"]), ["Topping"]),
    "Slab": check_Area_for_zero(getting_Area_up(floors_collector, ["Slab"]), ["Slab"]),
    "Up_complition": check_Area_for_zero(getting_Area_up(floors_collector, ["Completion"]), ["Completion"])
}
floors_down_Volume = {
    "Dn_Regular": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                             ["Regular", "Regular-T", "Balcon"]),
    "Dn_Prestressed": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Dn_Rampa": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Rampa"]), ["Rampa"]),
    "Dn_Regular-W Special": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Regular-W Special"]),
                                                       ["Regular-W Special"]),
    "Dn_Koteret": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Koteret"]), ["Koteret"]),
    "Dn_Lite Beton": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Dn_Geoplast-D": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Dn_Slab-Rib Special": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Slab-Rib Special"]),
                                                      ["Slab-Rib Special"]),
    "CLSM": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["CLSM"]), ["CLSM"]),
    "Topping": check_Area_for_zero(getting_Area_up(floors_collector, ["Topping"]), ["Topping"]),
    "Slab": check_Area_for_zero(getting_Area_up(floors_collector, ["Slab"]), ["Slab"]),
    "Dn_complition": check_Area_for_zero(getting_Area_up(floors_collector, ["Completion"]), ["Completion"])
}
