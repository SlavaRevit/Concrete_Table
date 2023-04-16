from Autodesk.Revit.DB import *
import clr

# from openpyxl import  Workbook, load_workbook
# from openpyxl.utils import  get_column_letter


clr.AddReference("RevitAPI")
clr.AddReference("System")
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document

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
            # res_of_beams = res_of_beams + result
        if len(beam_type_check) > 1:
            # print("Volume of Beams is {}".format(result))
            return round(result,1)
        elif len(beam_type_check) == 1:
            if type == "Concrete":
                count_anchor = getting_Count(beams_collector, [type])
                # print("Volume/Count of Beam Anchor Concrete  is {}/{}".format(result,count_anchor))
                return round(result,1), count_anchor
            elif type == "Precast":
                res = getting_Volume(beams_collector, [type])
                # print("Volume/Count of Beam Pergoa Precast  is {}/{}".format(res, result))
                return round(res, 1), result
            else:
                # print("Volume of {} is {}".format(type, result))
                return round(result, 1)


beams = {
    "regular_beams": check_for_zero(
        getting_Volume(beams_collector, ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"]),
        ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"]),
    "beam_tooth": check_for_zero(getting_Volume(beams_collector, ["Regular-T"]), ["Regular-T"]),
    "beam_prestressed": check_for_zero(getting_Volume(beams_collector, ["Prestressed"]), ["Prestressed"]),
    "beam_foundation": check_for_zero(getting_Volume(beams_collector, ["Foundation"]), ["Foundation"]),
    "beam_head": check_for_zero(getting_Volume(beams_collector, ["Head"]), ["Head"]),
    "beam_anchor": check_for_zero(getting_Volume(beams_collector, ["Concrete"]), ["Concrete"]),
    "beam_precast": check_for_zero(getting_Count(beams_collector, ["Precast"]), ["Precast"])
}
