"""________Walls________"""
from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document



wall_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_Walls). \
    WhereElementIsNotElementType(). \
    ToElementIds()

slurry_Bisus = {}
slurry_Dipun = {}
walls_in_new = {}
walls_out_new = {}

def getting_Area_Volume_walls(walls_list):
    for wall in wall_collector:
        new_w = doc.GetElement(wall)
        # going to WallType
        wall_type = new_w.WallType
        volume_param = new_w.LookupParameter("Volume")
        area_param = new_w.LookupParameter("Area")
        wall_duplicationTypeMark = wall_type.LookupParameter("Duplication Type Mark").AsString()
        wall_type_comments = wall_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()

        if wall_duplicationTypeMark == "Slurry Bisus":
            if wall_duplicationTypeMark in slurry_Bisus:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                slurry_Bisus[wall_duplicationTypeMark]["Area"] += wall_area
                slurry_Bisus[wall_duplicationTypeMark]["Volume"] += wall_volume
            elif wall_duplicationTypeMark == "Slurry Bisus":
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                slurry_Bisus[wall_duplicationTypeMark] = {"Area": wall_area, "Volume": wall_volume}

        if wall_duplicationTypeMark == "Slurry Dipun":
            if wall_duplicationTypeMark in slurry_Dipun:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                slurry_Dipun[wall_duplicationTypeMark]["Area"] += wall_area
                slurry_Dipun[wall_duplicationTypeMark]["Volume"] += wall_volume
            elif wall_duplicationTypeMark == "Slurry Dipun":
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                slurry_Dipun[wall_duplicationTypeMark] = {"Area": wall_area, "Volume": wall_volume}

        if wall_type_comments == "FrIN":
            wall_key = "קירות פנימיים מבטון"
            if wall_key not in walls_in_new:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                walls_in_new[wall_key] = {"Area": wall_area, "Volume": wall_volume}
            elif wall_key in walls_in_new:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                walls_in_new[wall_key]["Area"] += wall_area
                walls_in_new[wall_key]["Volume"] += wall_volume

        if wall_type_comments == "FrOut":
            wall_key = "קירות חוץ מבטון"
            if wall_key not in walls_out_new:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                walls_out_new[wall_key] = {"Area": wall_area, "Volume": wall_volume}
            elif wall_key in walls_out_new:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                walls_out_new[wall_key]["Area"] += wall_area
                walls_out_new[wall_key]["Volume"] += wall_volume

getting_Area_Volume_walls(wall_collector)
