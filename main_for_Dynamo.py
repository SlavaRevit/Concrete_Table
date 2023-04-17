import sys
import os
import clr
import sys
import pandas as pd
import xlsxwriter

localapp = os.getenv(r'LOCALAPPDATA')
sys.path.append(os.path.join(localapp, r'python-3.8.3-embed-amd64\Lib\site-packages'))

import System
from System import Array
from System.Collections.Generic import *

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

floors_up_Area = {
    "Up_Regular": check_Area_for_zero(getting_Area_up(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                      ["Regular", "Regular-T", "Balcon"]),
    "Up_Regular-P": check_Area_for_zero(getting_Area_up(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Up_Rampa": check_Area_for_zero(getting_Area_up(floors_collector, ["Rampa"]), ["Rampa"]),
    "Regular-W Special": check_Area_for_zero(getting_Area_up(floors_collector, ["Regular-W Special"]),
                                             ["Regular-W Special"]),
    "Koteret": check_Area_for_zero(getting_Area_up(floors_collector, ["Koteret"]), ["Koteret"]),
    "Lite Beton": check_Area_for_zero(getting_Area_up(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Geoplast-D": check_Area_for_zero(getting_Area_up(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Slab-Rib Special": check_Area_for_zero(getting_Area_up(floors_collector, ["Slab-Rib Special"]),
                                            ["Slab-Rib Special"]),
    "CLSM": check_Area_for_zero(getting_Area_up(floors_collector, ["CLSM"]), ["CLSM"])
}

floors_down_Area = {
    "Dn_Regular": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                           ["Regular", "Regular-T", "Balcon"]),
    "Dn_Regular-P": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Dn_Rampa": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Rampa"]), ["Rampa"]),
    "Dn_Regular-W Special": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Regular-W Special"]),
                                                     ["Regular-W Special"]),
    "Dn_Koteret": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Koteret"]), ["Koteret"]),
    "Dn_Lite Beton": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Dn_Geoplast-D": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Dn_Slab-Rib Special": check_Area_for_zero_down(getting_Area_down(floors_collector, ["Slab-Rib Special"]),
                                                    ["Slab-Rib Special"]),
    "CLSM": check_Area_for_zero_down(getting_Area_down(floors_collector, ["CLSM"]), ["CLSM"])
}
floors_up_Volume = {
    "Up_Regular": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                        ["Regular", "Regular-T", "Balcon"]),
    "Up_Regular-P": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Up_Rampa": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Rampa"]), ["Rampa"]),
    "Regular-W Special": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Regular-W Special"]),
                                               ["Regular-W Special"]),
    "Koteret": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Koteret"]), ["Koteret"]),
    "Lite Beton": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Geoplast-D": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Slab-Rib Special": check_Volume_for_zero(getting_Volume_up(floors_collector, ["Slab-Rib Special"]),
                                              ["Slab-Rib Special"]),
    "CLSM": check_Volume_for_zero(getting_Volume_up(floors_collector, ["CLSM"]), ["CLSM"])
}

floors_down_Volume = {
    "Dn_Regular": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Regular", "Regular-T", "Balcon"]),
                                             ["Regular", "Regular-T", "Balcon"]),
    "Dn_Regular-P": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Regular-P"]), ["Regular-P"]),
    "Dn_Rampa": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Rampa"]), ["Rampa"]),
    "Dn_Regular-W Special": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Regular-W Special"]),
                                                       ["Regular-W Special"]),
    "Dn_Koteret": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Koteret"]), ["Koteret"]),
    "Dn_Lite Beton": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Lite Beton"]), ["Lite Beton"]),
    "Dn_Geoplast-D": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Geoplast-D"]), ["Geoplast-D"]),
    "Dn_Slab-Rib Special": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["Slab-Rib Special"]),
                                                      ["Slab-Rib Special"]),
    "CLSM": check_Volume_for_zero_down(getting_Volume_down(floors_collector, ["CLSM"]), ["CLSM"])
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

wall_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_Walls). \
    WhereElementIsNotElementType(). \
    ToElementIds()
walls_Out = []
walls_In = []
for wall_t in wall_collector:
    # geting element from wall.id
    new_w = doc.GetElement(wall_t)
    # going to WallType
    wall_type = new_w.WallType
    # getting type_comments of this wall
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
    "walls_In": getting_Volume_Area_IN(wall_collector),
    "walls_Out": getting_Volume_Area_Out(wall_collector)
}

"""________Creating_Excel_File________"""  # attempt 1

# creating data_frama of beams from dict
df_beams = pd.DataFrame.from_dict(beams, orient="index", columns=["Value"])
df_beams = df_beams.dropna()
df_walls = pd.DataFrame.from_dict(walls, orient="index", columns=["Area", "Volume"])

# from tuples spliting result to 2 different columns, column1 = Volume, column2 = Count
df_beams[['Column1', 'Column2']] = df_beams['Value'].apply(
    lambda x: pd.Series([x[0], x[1]] if isinstance(x, tuple) and len(x) >= 1 else [x, None]))

df_beams_result = df_beams.drop(columns=['Value']).rename(columns={'Column1': 'Volume', 'Column2': 'Count'})

"""Series for floor Area/Volume UP"""
floors_up_Area_filtered = {k: v for k, v in floors_up_Area.items() if v}
area_up_series = pd.Series(floors_up_Area_filtered, name="Area")
floors_up_Volume_filtered = {k: v for k, v in floors_up_Volume.items() if v}
volume_up_series = pd.Series(floors_up_Volume_filtered, name="Volume")
df_floors_up = pd.concat([area_up_series, volume_up_series], axis=1)

"""Series for floor Area/Volume DOWN"""
floors_up_Area_filtered = {k: v for k, v in floors_down_Area.items() if v}
area_down_series = pd.Series(floors_up_Area_filtered, name="Area")
floors_up_Volume_filtered = {k: v for k, v in floors_down_Volume.items() if v}
volume_down_series = pd.Series(floors_up_Volume_filtered, name="Volume")
df_floors_down = pd.concat([area_down_series, volume_down_series], axis=1)

# Inserting new Name of Type:

name_up = "Floors"
df_name_floors_up = pd.DataFrame(index=["Floors"])
name_down = "Floors_Down"
df_name_floors_down = pd.DataFrame(index=["Floors_Down"])
name_beams = "Beams"
df_name_beams = pd.DataFrame(index=['Beams'])
name_walls = "Walls"
df_name_walls = pd.DataFrame(index=['Walls'])

# concatenating this 2 tables
# df = pd.concat([df_name_beams, df_beams, df_name_floors, df_floors], axis=0, sort=False)
# df = pd.concat([df_name_beams,df_beams,df_name_floors,df_floors_up_GENERAL,df_name_floors,df_floors_area_down,df_floors_volume_down,df_name_walls,df_walls ], axis=0, sort=False)

df = pd.concat(
    [df_name_beams, df_beams_result, df_name_floors_up, df_floors_up, df_floors_down, df_name_walls, df_walls], axis=0,
    sort=False).round(2)
col_1_sum = df['Area'].sum()
col_2_sum = df['Volume'].sum()
total_row = pd.Series({"Area":col_1_sum, "Volume":col_2_sum,"Count" : ""},name="Total")
df = df._append(total_row)



df = df.fillna('')
# changing position of columns
df = df.reindex(columns=["Area", "Volume", "Count"])

writer = pd.ExcelWriter(IN[1], engine='xlsxwriter')
df.to_excel(writer, sheet_name="Test")
writer._save()
