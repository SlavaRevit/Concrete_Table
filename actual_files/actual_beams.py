from Autodesk.Revit.DB import *
import clr

# from openpyxl import  Workbook, load_workbook
# from openpyxl.utils import  get_column_letter


clr.AddReference("RevitAPI")
clr.AddReference("System")


doc = __revit__.ActiveUIDocument.Document

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

        if custom_param == "Precast":
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


# def getting_Volume(beam_list, beam_type_check):
#     total_Volume = 0.0
#     for el in beam_list:
#         element_type = doc.GetElement(el.GetTypeId())
#         custom_param = element_type.LookupParameter("Duplication Type Mark").AsString()
#         """checking for alot of types"""
#         for t in beam_type_check:
#             if custom_param == t:
#                 beam_volume = el.LookupParameter("Volume").AsDouble()
#                 total_Volume = total_Volume + beam_volume * 0.0283168466
#     return total_Volume
#
#
# def getting_Count(beam_list, beam_type_check):
#     total_count = 0
#     for el in beam_list:
#         element_type = doc.GetElement(el.GetTypeId())
#         custom_param = element_type.LookupParameter("Duplication Type Mark").AsString()
#         for t in beam_type_check:
#             if custom_param == t:
#                 total_count = total_count + 1
#     return total_count
#
#
# def check_for_zero(result, beam_type_check):
#     if result == 0:
#         pass
#     else:
#         for type in beam_type_check:
#             result_of_beams_vol = 0
#             str_check = ", ".join(beam_type_check)
#         if len(beam_type_check) > 1:
#             return result
#         elif len(beam_type_check) == 1:
#             if type == "Concrete":
#                 count_anchor = getting_Count(beams_collector, [type])
#
#                 return result, count_anchor
#             elif type == "Precast":
#                 res = getting_Volume(beams_collector, [type])
#
#                 return res, result
#             else:
#
#                 return result
