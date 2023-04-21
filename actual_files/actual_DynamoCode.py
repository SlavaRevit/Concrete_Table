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

floors_up = {}
floors_down = {}


def getting_floors_parameters(floor_list):
    for el in floor_list:
        floor_element = doc.GetElement(el)
        floor_type = floor_element.FloorType
        floor_type_comments = floor_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()
        floor_duplicationTypeMark = floor_type.LookupParameter("Duplication Type Mark").AsString()
        if floor_type_comments == "Up":
            if floor_duplicationTypeMark in ["Total Floor Area", "Total Floor Area Commercial", "Total Floor Area LSP",
                                             "Total Floor Area Pergola",
                                             "Air Double Level", "Air Elevator", "Air Pergola Aluminium",
                                             "Air Pergola Steel", "Air Pergola Wood", "Air Regular", "Air Stairs",
                                             "Aggregate", "Backfilling", "Polivid"]:
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
                # combine Regular, Balcon, Regular-T keys

        elif floor_type_comments == "Down":
            if floor_duplicationTypeMark in ["Total Floor Area", "Total Floor Area Commercial", "Total Floor Area LSP",
                                             "Total Floor Area Pergola",
                                             "Air Double Level", "Air Elevator", "Air Pergola Aluminium",
                                             "Air Pergola Steel", "Air Pergola Wood", "Air Regular", "Air Stairs",
                                             "Aggregate", "Backfilling", "Polivid"]:
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

"""________BEAMS________"""

beams_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_StructuralFraming). \
    WhereElementIsNotElementType(). \
    ToElements()

beams = {}

def beams_parameters(beam_list):
    for el in beam_list:
        element_type = doc.GetElement(el.GetTypeId())
        custom_param = element_type.LookupParameter("Duplication Type Mark").AsString()
        if custom_param == "Beam Steel":
            continue

        if custom_param == "Beam Anchor":
            if custom_param not in beams:
                beams[custom_param] = {"Count": 1}
            else:
                beams[custom_param]['Count'] += 1

        elif custom_param == "Anchor Polymer":
            if custom_param not in beams:
                beams[custom_param] = {"Count": 1}
            else:
                beams[custom_param]['Count'] += 1

        elif custom_param == "Anchor Steel":
            if custom_param not in beams:
                beams[custom_param] = {"Count": 1}
            else:
                beams[custom_param]['Count'] += 1

        elif custom_param == "Precast":
            if custom_param not in beams:
                beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                beams[custom_param] = {"Volume": beam_volume, "Count": 1}
            else:
                beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                beams[custom_param]['Volume'] += beam_volume
                beams[custom_param]['Count'] += 1

        elif custom_param == "Concrete":
            if custom_param not in beams:
                beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                beams[custom_param] = {"Volume": beam_volume, "Count": 1}
            else:
                beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                beams[custom_param]['Volume'] += beam_volume
                beams[custom_param]['Count'] += 1

        elif custom_param in ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"]:
            combined_key = "Beams_new"
            if combined_key not in beams:
                beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                beams[combined_key] = {"Volume": beam_volume}
            elif combined_key in beams:
                beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                beams[combined_key]['Volume'] += beam_volume
            # else:
            #     beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
            #     beams[combined_key] = {"Volume": beam_volume}
        elif custom_param not in beams:
            beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
            beams[custom_param] = {"Volume": beam_volume}
        else:
            beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
            beams[custom_param]["Volume"] += beam_volume

        # else:
        #     beam_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
        #     beams[custom_param]['Volume'] += beam_volume

    return beams

beams_parameters(beams_collector)


"""________Walls________"""

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
        """Getting Slurry Bisus walls"""
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
        """Getting Slurry Dipun walls"""
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
        """Getting all walls IN"""
        if wall_type_comments == "FrIN":
            wall_key = "Walls_In"
            if wall_key not in walls_in_new:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                walls_in_new[wall_key] = {"Area": wall_area, "Volume": wall_volume}
            elif wall_key in walls_in_new:
                wall_area = area_param.AsDouble() * 0.092903
                wall_volume = volume_param.AsDouble() * 0.0283168466
                walls_in_new[wall_key]["Area"] += wall_area
                walls_in_new[wall_key]["Volume"] += wall_volume
        """Getting all walls Out"""
        if wall_type_comments == "FrOut":
            wall_key = "Walls_Out"
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

"""________Foundations________"""

foundation_collector = FilteredElementCollector(doc). \
    OfCategory(BuiltInCategory.OST_StructuralFoundation). \
    WhereElementIsNotElementType(). \
    ToElements()

Dipuns = {}
Bisus = {}
Basic_Plate = {}
Found_Head = {}
Rafsody = {}


def getting_Length_Volume_Count(found_list):
    for el in found_list:
        if el.Category.Name == "Structural Foundations":
            if isinstance(el, FamilyInstance):
                el_type_id = el.GetTypeId()
                type_elem = doc.GetElement(el_type_id)
                if type_elem:
                    parameter_Duplication = type_elem.LookupParameter("Duplication Type Mark").AsString()
                    if parameter_Duplication == "Dipun":
                        parameter = el.LookupParameter("Length")
                        parameter_vol = el.LookupParameter("Volume")
                        parameter_Descr = type_elem.LookupParameter("Description").AsValueString()
                        if parameter_Descr not in Dipuns:
                            parameter_value = round(parameter.AsDouble() * 0.3048)
                            parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                            Dipuns[parameter_Descr] = {'Length': parameter_value, 'Volume': parameter_value_vol,
                                                       'Count': 1}
                        else:
                            parameter_value = round(parameter.AsDouble() * 0.3048)
                            parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                            Dipuns[parameter_Descr]['Length'] += parameter_value
                            Dipuns[parameter_Descr]['Volume'] += parameter_value_vol
                            Dipuns[parameter_Descr]['Count'] += 1


                    elif parameter_Duplication == "Bisus":
                        parameter = el.LookupParameter("Length")
                        parameter_vol = el.LookupParameter("Volume")
                        parameter_Descr = type_elem.LookupParameter("Description").AsValueString()
                        if parameter_Descr not in Bisus:
                            parameter_value = round(parameter.AsDouble() * 0.3048)
                            parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                            Bisus[parameter_Descr] = {'Length': parameter_value, 'Volume': parameter_value_vol,
                                                      'Count': 1}
                        else:
                            parameter_value = round(parameter.AsDouble() * 0.3048)
                            parameter_value_vol = parameter_vol.AsDouble() * 0.0283168466
                            Bisus[parameter_Descr]['Length'] += parameter_value
                            Bisus[parameter_Descr]['Volume'] += parameter_value_vol
                            Bisus[parameter_Descr]['Count'] += 1

                else:
                    pass
            # For Floor Types ( Rafsody, Head, BasicPlate )
            elif isinstance(el, Floor):
                el_type_id = el.GetTypeId()
                # foundation_element = doc.GetElement(el)
                foundation_type = el.FloorType
                foundation_duplicationTypeMark = foundation_type.LookupParameter("Duplication Type Mark").AsString()
                if foundation_duplicationTypeMark == "Rafsody":
                    if foundation_duplicationTypeMark not in Rafsody:
                        foundation_area = el.LookupParameter("Area").AsDouble() * 0.092903
                        foundation_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                        Rafsody[foundation_duplicationTypeMark] = {"Area": foundation_area, "Volume": foundation_volume}
                    else:
                        foundation_area = el.LookupParameter("Area").AsDouble() * 0.092903
                        foundation_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                        Rafsody[foundation_duplicationTypeMark]["Area"] += foundation_area
                        Rafsody[foundation_duplicationTypeMark]["Volume"] += foundation_volume
                elif foundation_duplicationTypeMark == "Basic Plate":
                    if foundation_duplicationTypeMark not in Basic_Plate:
                        foundation_area = el.LookupParameter("Area").AsDouble() * 0.092903
                        foundation_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                        Basic_Plate[foundation_duplicationTypeMark] = {"Area": foundation_area,
                                                                       "Volume": foundation_volume}
                    else:
                        foundation_area = el.LookupParameter("Area").AsDouble() * 0.092903
                        foundation_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                        Basic_Plate[foundation_duplicationTypeMark]["Area"] += foundation_area
                        Basic_Plate[foundation_duplicationTypeMark]["Volume"] += foundation_volume
                elif foundation_duplicationTypeMark == "Head":
                    if foundation_duplicationTypeMark not in Found_Head:
                        foundation_area = el.LookupParameter("Area").AsDouble() * 0.092903
                        foundation_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                        Found_Head[foundation_duplicationTypeMark] = {"Area": foundation_area,
                                                                      "Volume": foundation_volume}
                    else:
                        foundation_area = el.LookupParameter("Area").AsDouble() * 0.092903
                        foundation_volume = el.LookupParameter("Volume").AsDouble() * 0.0283168466
                        Found_Head[foundation_duplicationTypeMark]["Area"] += foundation_area
                        Found_Head[foundation_duplicationTypeMark]["Volume"] += foundation_volume

    return Dipuns, Bisus, Rafsody, Basic_Plate, Found_Head

getting_Length_Volume_Count(foundation_collector)

"""________Creating_Excel_File________"""  # attempt 1
df_found_dipuns = pd.DataFrame.from_dict(Dipuns, orient="index", columns=["Length", "Volume", "Count"])
df_found_dipuns_sorted = df_found_dipuns.sort_index()
df_found_dipuns_sorted = df_found_dipuns.copy()
df_found_dipuns_sorted.index = pd.to_numeric(df_found_dipuns_sorted.index, errors = "coerce")
df_found_dipuns_sorted = df_found_dipuns_sorted.sort_index().fillna(0)

df_found_bisus = pd.DataFrame.from_dict(Bisus, orient="index", columns=["Length", "Volume", "Count"])
# df_found_bisus_sorted = df_found_bisus.sort_values(by=df_found_bisus.index, axis=1)
df_found_bisus_sorted = df_found_bisus.copy()
df_found_bisus_sorted.index = pd.to_numeric(df_found_bisus_sorted.index, errors="coerce")
df_found_bisus_sorted = df_found_bisus_sorted.sort_index().fillna(0)

# df_found_bisus_sorted = df_found_bisus.sort_index()


df_found_Slurry_Bisus = pd.DataFrame.from_dict(slurry_Bisus, orient="index", columns=["Area", "Volume"])

df_found_Slurry_Dipuns = pd.DataFrame.from_dict(slurry_Dipun, orient="index", columns=["Area", "Volume"])

df_floors_up = pd.DataFrame.from_dict(floors_up, orient="index", columns=["Area", "Volume"])
df_floors_down = pd.DataFrame.from_dict(floors_down, orient="index", columns=["Area", "Volume"])

# creating data_frama of beams from dict
df_beams = pd.DataFrame.from_dict(beams, orient="index", columns=["Volume", "Count"])
# df_beams = df_beams.dropna()
df_walls_in = pd.DataFrame.from_dict(walls_in_new, orient="index", columns=["Area", "Volume"])
df_walls_out = pd.DataFrame.from_dict(walls_out_new, orient="index", columns=["Area", "Volume"])

# from tuples spliting result to 2 different columns, column1 = Volume, column2 = Count
# df_beams[['Column1', 'Column2']] = df_beams['Value'].apply(
#     lambda x: pd.Series([x[0], x[1]] if isinstance(x, tuple) and len(x) >= 1 else [x, None]))
# df_beams_result = df_beams.drop(columns=['Value']).rename(columns={'Column1': 'Volume', 'Column2': 'Count'})


"""Inserting new nam of type_element"""
name_up = "Floors"
df_name_floors_up = pd.DataFrame(index=["Floors"])
name_down = "Floors_Down"
df_name_floors_down = pd.DataFrame(index=["Floors_Down"])
name_beams = "Beams"
df_name_beams = pd.DataFrame(index=['Beams'])
name_walls = "Walls"
df_name_walls = pd.DataFrame(index=['Walls'])
name_foundation = "Foundation"
df_name_Dipuns = pd.DataFrame(index=['Dipuns'])
df_name_Bisus = pd.DataFrame(index=['Bisus'])
df_name_Slurry_walls = pd.DataFrame(index=['Slurry'])

df = pd.concat(
    [df_name_Slurry_walls, df_found_Slurry_Bisus, df_found_Slurry_Dipuns,
     df_name_Dipuns, df_found_dipuns_sorted,
     df_name_Bisus, df_found_bisus_sorted,
     df_name_beams, df_beams,
     df_name_floors_up,df_floors_up, df_floors_down,
     df_name_walls, df_walls_in, df_walls_out], axis=0,
    sort=False).round(2)
col_1_sum = df['Area'].sum()
col_2_sum = df['Volume'].sum()
total_row = pd.Series({"Area": col_1_sum, "Volume": col_2_sum, "Count": ""}, name="Total")
df = df._append(total_row)

df = df.fillna('')
# changing position of columns
df = df.reindex(columns=["Area", "Volume", "Count", "Length"])

writer = pd.ExcelWriter(IN[1], engine='xlsxwriter')
df.to_excel(writer, sheet_name="Test")
writer._save()
