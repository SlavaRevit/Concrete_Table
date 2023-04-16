# Boilerplate text

import pandas as pd
import sys
import clr
import System
from System import Array
from System.Collections.Generic import *

getting_Area_Volume = getting_Area

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitNodes")
import Revit

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument

"""________FLOORS________"""
floors_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_Floors). \
    WhereElementIsNotElementType(). \
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


"""________BEAMS________"""

beams_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_StructuralFraming). \
    WhereElementIsNotElementType(). \
    ToElements()


def getting_Volume(beam_list, beam_type_check):
    total_Volume = 0.0
    for el in beam_list:
        element_type = doc.GetElement(el.GetTypeId())
        custom_param = element_type.LookupParameter("Duplication Type Mark").AsString()
        """checking for alot of types"""
        for t in beam_type_check:
            if custom_param == t:
                beam_volume = el.LookupParameter("Volume").AsDouble()
                total_Volume = total_Volume + beam_volume * 0.0283168466
    return total_Volume


def getting_Count(beam_list, beam_type_check):
    total_count = 0
    for el in beam_list:
        element_type = doc.GetElement(el.GetTypeId())
        custom_param = element_type.LookupParameter("Duplication Type Mark").AsString()
        for t in beam_type_check:
            if custom_param == t:
                total_count = total_count + 1
    return total_count


def check_for_zero(result, beam_type_check):
    if result == 0:
        pass
    else:
        for type in beam_type_check:
            result_of_beams_vol = 0
            str_check = ", ".join(beam_type_check)
        if len(beam_type_check) > 1:
            return result
        elif len(beam_type_check) == 1:
            if type == "Concrete":
                count_anchor = getting_Count(beams_collector, [type])

                return result, count_anchor
            elif type == "Precast":
                res = getting_Volume(beams_collector, [type])

                return res, result
            else:

                return result


"""________FLOORS________"""
floors_up = {
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
"""____FLORS_Down_____"""
floors_down = {
    "regular": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                              ["Regular", "Regular-T", "Balcon"]),
    "reg_prestressed": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "ramp": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Rampa"]), ["Rampa"]),
    "special": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Regular-W Special"]), ["Regular-W Special"]),
    "koteret": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Koteret"]), ["Koteret"]),
    "lite_beton": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "geoplast": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "rib_special": check_for_zero(getting_Area_Volume_Down(floors_collector, ["Slab-Rib Special"]), ["Slab-Rib Special"]),
    "CLSM": check_for_zero(getting_Area_Volume_Down(floors_collector, ["CLSM"]), ["CLSM"])
}
"""________Beams________"""
beams = {
    "regular_beams": check_for_zero(
        getting_Volume(beams_collector, ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"]),
        ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"]),

    "beam_tooth": check_for_zero(getting_Volume(beams_collector, ["Regular-T"]), ["Regular-T"]),
    "beam_prestressed": check_for_zero(getting_Volume(beams_collector, ["Prestressed"]), ["Prestressed"]),
    "beam_foundation": check_for_zero(getting_Volume(beams_collector, ["Foundation"]), ["Foundation"]),
    "beam_head": check_for_zero(getting_Volume(beams_collector, ["Head"]), ["Head"]),
    "beam_anchor": check_for_zero(getting_Volume(beams_collector, ["Concrete"]), ["Concrete"]),
    "beam_precast": check_for_zero(getting_Count(beams_collector, ["Precast"]), ["Precast"]),
}


"""________Walls________"""

wall_collector = FilteredElementCollector(doc).\
    OfCategory(BuiltInCategory.OST_Walls).\
    WhereElementIsNotElementType().\
    ToElementIds()
walls_Out = []
walls_In = []
for wall_t in wall_collector:
    #geting element from wall.id
    new_w = doc.GetElement(wall_t)
    # going to WallType
    wall_type = new_w.WallType
    #getting type_comments of this wall
    wall_type_comments = wall_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
    if wall_type_comments == "FrOut":
        walls_Out.append(new_w)
    elif wall_type_comments == "FrIN":
        walls_In.append(new_w)

def getting_Volume_Area_Out(walls_list):
    total_Area_Out = 0.0
    total_Volume_Out = 0.0
    for w_out in walls_Out:
        volume_param_Out = w_out.LookupParameter("Volume")
        area_param_Out = w_out.LookupParameter("Area")
        total_Volume_Out = total_Volume_Out + volume_param_Out.AsDouble() * 0.0283168466
        total_Area_Out = total_Area_Out + area_param_Out.AsDouble() * 0.092903
    return total_Area_Out, total_Volume_Out


def getting_Volume_Area_IN(walls_list):
    total_Area_In = 0.0
    total_Volume_In = 0.0
    for w_in in walls_In:
        volume_param_In = w_in.LookupParameter("Volume")
        area_param_In = w_in.LookupParameter("Area")
        total_Volume_In = total_Volume_In + volume_param_In.AsDouble() * 0.0283168466
        total_Area_In = total_Area_In + area_param_In.AsDouble() * 0.092903
    return total_Area_In, total_Volume_In



walls = {
    "walls_In" : getting_Volume_Area_IN(wall_collector),
    "walls_Out" : getting_Volume_Area_Out(wall_collector)
}


# """________Creating_Excel_File________""" # attempt 1
# # creating data_frama of beams from dict
# df_beams = pd.DataFrame.from_dict(beams, orient="index", columns=["Value"])
#
# # from tuples spliting result to 2 different columns, column1 = Volume, column2 = Count
# df_beams[['Column1', 'Column2']] = df_beams['Value'].apply(
#     lambda x: pd.Series([x[0], x[1]] if isinstance(x, tuple) else [x, None]))
# df_beams = df_beams.drop(columns=['Value']).rename(columns={'Column1': 'Volume', 'Column2': 'Count'})
#
# # creating data_frama of floors from dict
# my_dict_floors = {k: v for k, v in floors.items() if v is not None}
# df_floors = pd.DataFrame.from_dict(my_dict_floors, orient='index', columns=['Area', 'Volume'])
# # Inserting new Name of Type:
#
# name = "Floors"
# df_name_floors = pd.DataFrame(index=["Floors"])
# name_beams = "Beams"
# df_name_beams = pd.DataFrame(index=['Beams'])
# name_walls = "Walls"
# df_name_walls = pd.DataFrame(index=['Walls'])
#
# # concatenating this 2 tables
# # df = pd.concat([df_name_beams, df_beams, df_name_floors, df_floors], axis=0, sort=False)
# df = pd.concat([df_beams,df_floors], axis=0, sort=False)
# df = df.fillna('')
# # changing position of columns
# df = df.reindex(columns=["Area", "Volume", "Count"])
#
# writer = pd.ExcelWriter(IN[1], engine='xlsxwriter')
# df.to_excel(writer, sheet_name="Test")
# writer._save()



#attempt 2
my_dict_floors = {k: v for k, v in floors.items() if v is not None}
my_dict_beams = {k: [v] if not isinstance(v, list) else v for k, v in beams.items() if v is not None}

df_floors = pd.DataFrame.from_dict(my_dict_floors, orient="index", columns=["Area", "Volume"])
df_beams = pd.DataFrame.from_dict(my_dict_beams, orient="index", columns=["Volume", "Count"])

df_rounded_floors = df_floors.round(2)
# beams_name = "Beams"
beams_name_df = pd.DataFrame(index=["Beams"])
# floors_name = "Floors"
floors_name_df = pd.DataFrame(index=["Floors"])
# combaning all DataFrames
df = pd.concat([floors_name_df, df_rounded_floors, beams_name_df, df_beams])
# print(df)
writer = pd.ExcelWriter("file1.xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name="Test")
writer._save()