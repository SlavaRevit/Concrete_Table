from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document

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
            print("Volume of Beams is {}".format(result))
        elif len(beam_type_check) == 1:
            if type == "Concrete":
                count_anchor = getting_Count(beams_collector,[type])
                print("Volume/Count of Beam Anchor Concrete  is {}/{}".format(result,count_anchor))
            elif type == "Precast":
                res = getting_Volume(beams_collector,[type])
                print("Volume/Count of Beam Pergola Precast  is {}/{}".format(res, result))
            else:
                print("Volume of {} is {}".format(type, result))


# check_for_zero(getting_Volume(beams_collector, ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"]),
#                ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"])
#
# check_for_zero(getting_Volume(beams_collector, ["Regular-T"]), ["Regular-T"])
# check_for_zero(getting_Volume(beams_collector, ["Prestressed"]), ["Prestressed"])
# check_for_zero(getting_Volume(beams_collector, ["Foundation"]), ["Foundation"])
# check_for_zero(getting_Volume(beams_collector, ["Head"]), ["Head"])
#
# check_for_zero(getting_Volume(beams_collector, ["Concrete"]), ["Concrete"])
# check_for_zero(getting_Count(beams_collector, ["Precast"]), ["Precast"])

