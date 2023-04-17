from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")
from System.Collections.Generic import List
doc = __revit__.ActiveUIDocument.Document


foundation_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_StructuralFoundation). \
    WhereElementIsNotElementType(). \
    ToElementIds()


def getting_Area(found_list, found_type_check):
    total_Area = 0.0
    for el in found_list:
        foundation_element = doc.GetElement(el)
        foundation_type = foundation_element.FloorType
        # floor_type_comments = foundation_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = foundation_type.LookupParameter("Duplication Type Mark").AsString()
        for type_el in found_type_check:
            if floor_duplicationTypeMark == type_el:
                foundation_area = foundation_element.LookupParameter("Area").AsDouble()
                total_Area = total_Area + foundation_area * 0.092903
    return total_Area


def getting_Volume(found_list, found_type_check):
    total_Volume = 0.0
    for el in found_list:
        foundation_element = doc.GetElement(el)
        foundation_type = foundation_element.FloorType
        # floor_type_comments = foundation_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = foundation_type.LookupParameter("Duplication Type Mark").AsString()
        for type_el in found_type_check:
            if floor_duplicationTypeMark == type_el:
                foundation_area = foundation_element.LookupParameter("Volume").AsDouble()
                total_Volume = total_Volume + foundation_area * 0.0283168466
    return total_Volume


def check_Area_for_zero(result, found_type_check):
    if result == 0:
        pass
    else:
        for type in found_type_check:
            result_of_beams_vol = 0
            str_check = ", ".join(found_type_check)
            # res_of_beams = res_of_beams + result
        if len(found_type_check) > 1:
            area = 0
            for p in found_type_check:
                total_a = getting_Area(foundation_collector, [p])
                area = area + total_a
            return area
        elif len(found_type_check) == 1:
            area = getting_Area(foundation_collector, [type])
            return area


def check_Volume_for_zero(result, found_type_check):
    if result == 0:
        pass
    else:
        for type in found_type_check:
            result_of_beams_vol = 0
            str_check = ", ".join(found_type_check)
            # res_of_beams = res_of_beams + result
        if len(found_type_check) > 1:
            volume = 0
            for p in found_type_check:
                total_v = getting_Volume(foundation_collector, [p])
                volume = volume + total_v
            return volume
        elif len(found_type_check) == 1:
            volume = getting_Volume(foundation_collector, [type])
            return volume


"""________Foundation________"""

foundation_Area = {
    "Rafsody" : check_Area_for_zero(getting_Area(foundation_collector, ["Rafsody"]),["Rafsody"]),
    "Basic Plate" : check_Area_for_zero(getting_Area(foundation_collector, ["Basic Plate"]),["Basic Plate"]),
    "Foundation Head" : check_Area_for_zero(getting_Area(foundation_collector, ["Head"]),["Head"])
}

foundation_Volume = {
    "Rafsody" : check_Volume_for_zero(getting_Volume(foundation_collector, ["Rafsody"]),["Rafsody"]),
    "Basic Plate" : check_Volume_for_zero(getting_Volume(foundation_collector, ["Basic Plate"]),["Basic Plate"]),
    "Foundation Head" : check_Volume_for_zero(getting_Volume(foundation_collector, ["Head"]),["Head"])
}