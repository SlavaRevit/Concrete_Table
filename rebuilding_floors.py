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


def getting_Volume(floor_list,floor_type_check):
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


def getting_Area(floor_list,floor_type_check):
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


def getting_Volume_down(floor_list,floor_type_check):
    total_Volume_down = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Down":
            for type_el in floor_type_check:
                if floor_duplicationTypeMark == type_el:
                    floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                    total_Volume_down = total_Volume_down + floor_volume * 0.0283168466
    return total_Volume_down


def getting_Area_down(floor_list,floor_type_check):
    total_Area_down = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Down":
            for type_el in floor_type_check:
                if floor_duplicationTypeMark == type_el:
                    floor_area = floor_element.LookupParameter("Area").AsDouble()
                    total_Area_down = total_Area_down + floor_area * 0.092903
    return total_Area_down



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
                total_v = getting_Volume(floors_collector,[p])
                volume = volume + total_v
            print("Area/Volume of Floors_Regular_Up is {}/{}".format(result,volume))
        elif len(floor_type_check) == 1:
            volume = getting_Volume(floors_collector, [type])
            print("Area/Volume of {}  is {}/{}".format(type,result,volume))


def check_for_zero_down(result, floor_type_check):
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
                total_v = getting_Volume_down(floors_collector,[p])
                volume = volume + total_v
            print("Area/Volume of Floors_Regular_Down is {}/{}".format(result,volume))
        elif len(floor_type_check) == 1:
            volume = getting_Volume_down(floors_collector, [type])
            print("Area/Volume of {}  is {}/{}".format(type,result,volume))



# print("****Floors Up****")
# check_for_zero(getting_Area(floors_collector,["Regular","Regular-T","Balcon"]), ["Regular","Regular-T","Balcon"])
# check_for_zero(getting_Area(floors_collector,["Regular-P"]), ["Regular-P"])
check_for_zero(getting_Area(floors_collector,["Rampa"]), ["Rampa"])
# check_for_zero(getting_Area(floors_collector,["Regular-W Special"]), ["Regular-W Special"])
# check_for_zero(getting_Area(floors_collector,["Koteret"]), ["Koteret"])
# check_for_zero(getting_Area(floors_collector,["Lite Beton"]), ["Lite Beton"])
# check_for_zero(getting_Area(floors_collector,["Geoplast-D"]), ["Geoplast-D"])
# check_for_zero(getting_Area(floors_collector,["Slab-Rib Special"]), ["Slab-Rib Special"])
# check_for_zero(getting_Area(floors_collector,["CLSM"]), ["CLSM"])
# # print("\n")
# print("****Floors Down****")
# check_for_zero_down(getting_Area_down(floors_collector,["Regular","Regular-T","Balcon"]), ["Regular","Regular-T","Balcon"])
# check_for_zero_down(getting_Area_down(floors_collector,["Regular-P"]), ["Regular-P"])
# check_for_zero_down(getting_Area_down(floors_collector,["Rampa"]), ["Rampa"])
# check_for_zero_down(getting_Area_down(floors_collector,["Regular-W Special"]), ["Regular-W Special"])
# check_for_zero_down(getting_Area_down(floors_collector,["Koteret"]), ["Koteret"])
# check_for_zero_down(getting_Area_down(floors_collector,["Lite Beton"]), ["Lite Beton"])
# check_for_zero_down(getting_Area_down(floors_collector,["Geoplast-D"]), ["Geoplast-D"])
# check_for_zero_down(getting_Area_down(floors_collector,["Slab-Rib Special"]), ["Slab-Rib Special"])
# check_for_zero_down(getting_Area_down(floors_collector,["CLSM"]), ["CLSM"])

# print("Check Total for Floors")
# a = check_for_zero(getting_Area(floors_collector,["Regular","Regular-T","Balcon","Regular-P","Rampa","Regular-W Special","Koteret","Lite Beton","Geoplast-D","Slab-Rib Special","CLSM"]),
#                ["Regular","Regular-T","Balcon","Regular-P","Rampa","Regular-W Special","Koteret","Lite Beton","Geoplast-D","Slab-Rib Special","CLSM"])
# b = check_for_zero_down(getting_Area_down(floors_collector,["Regular","Regular-T","Balcon","Regular-P","Rampa","Regular-W Special","Koteret","Lite Beton","Geoplast-D","Slab-Rib Special","CLSM"]),
#                ["Regular","Regular-T","Balcon","Regular-P","Rampa","Regular-W Special","Koteret","Lite Beton","Geoplast-D","Slab-Rib Special","CLSM"])