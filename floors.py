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


def getting_Area_Volume(floor_list,floor_type_check):
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
    if total_Area == 0.0 or total_Volume == 0.0:
        pass
    else:
        return round(total_Area,2), round(total_Volume,2)

def getting_Area_Volume_Down(floor_list,floor_type_check):
    total_Area = 0.0
    total_Volume = 0.0
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Down":
            if floor_duplicationTypeMark == floor_type_check:
                floor_volume = floor_element.LookupParameter("Volume").AsDouble()
                floor_area = floor_element.LookupParameter("Area").AsDouble()
                total_Volume = total_Volume + floor_volume * 0.0283168466
                total_Area = total_Area + floor_area * 0.092903
    if total_Area == 0.0 or total_Volume == 0.0:
        pass
    else:
        return round(total_Area,2), round(total_Volume,2)




