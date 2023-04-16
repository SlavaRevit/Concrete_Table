"""Calculates total volume of all walls in the model."""
import rebuilding_floors
import beams
__autor__ = "Slava Filimonenko"
from Autodesk.Revit.DB import *
doc = __revit__.ActiveUIDocument.Document

#Creating collector instance and collecting all the walls from the model
wall_collector = FilteredElementCollector(doc).\
    OfCategory(BuiltInCategory.OST_Walls).\
    WhereElementIsNotElementType().\
    ToElementIds()
walls_Out = []
walls_In = []
total_Volume_Out = 0.0
total_Volume_In = 0.0
total_Area_Out = 0.0
total_Area_In = 0.0

total_Area = 0.0
total_Volume = 0.0
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

for w_out in walls_Out:
    volume_param_Out = w_out.LookupParameter("Volume")
    area_param_Out = w_out.LookupParameter("Area")
    total_Volume_Out = total_Volume_Out + volume_param_Out.AsDouble() * 0.0283168466
    total_Area_Out = total_Area_Out + area_param_Out.AsDouble() * 0.092903


for w_in in walls_In:
    volume_param_In = w_in.LookupParameter("Volume")
    area_param_In = w_in.LookupParameter("Area")
    total_Volume_In = total_Volume_In + volume_param_In.AsDouble() * 0.0283168466
    total_Area_In = total_Area_In + area_param_In.AsDouble() * 0.092903

"""Printing all totals for walls"""
print("****Walls Out****")
print("Area/Volume of Walls-Out: {}/{}".format(total_Area_Out, total_Volume_Out))
print("****Walls In****")
print("Area/Volume of Walls-In: {}/{}".format(total_Area_In, total_Volume_In))
# print("Total for walls is {}".format(total_volume_Out + total_volume_In ))
print("\n")



print("****Floors Up****")
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Regular","Regular-T","Balcon"]), ["Regular","Regular-T","Balcon"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Regular-P"]), ["Regular-P"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Rampa"]), ["Rampa"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Regular-W Special"]), ["Regular-W Special"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Koteret"]), ["Koteret"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Lite Beton"]), ["Lite Beton"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Geoplast-D"]), ["Geoplast-D"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["Slab-Rib Special"]), ["Slab-Rib Special"])
rebuilding_floors.check_for_zero(rebuilding_floors.getting_Area(rebuilding_floors.floors_collector,["CLSM"]), ["CLSM"])
# print("\n")
print("****Floors Down****")
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Regular","Regular-T","Balcon"]), ["Regular","Regular-T","Balcon"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Regular-P"]), ["Regular-P"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Rampa"]), ["Rampa"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Regular-W Special"]), ["Regular-W Special"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Koteret"]), ["Koteret"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Lite Beton"]), ["Lite Beton"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Geoplast-D"]), ["Geoplast-D"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["Slab-Rib Special"]), ["Slab-Rib Special"])
rebuilding_floors.check_for_zero_down(rebuilding_floors.getting_Area_down(rebuilding_floors.floors_collector,["CLSM"]), ["CLSM"])


print("**** Beams ****")
beams.check_for_zero(beams.getting_Volume(beams.beams_collector, ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"]),
               ["Regular", "Balcon", "Transformation", "Pergola", "Belts", "Tapered"])

beams.check_for_zero(beams.getting_Volume(beams.beams_collector, ["Regular-T"]), ["Regular-T"])
beams.check_for_zero(beams.getting_Volume(beams.beams_collector, ["Prestressed"]), ["Prestressed"])
beams.check_for_zero(beams.getting_Volume(beams.beams_collector, ["Foundation"]), ["Foundation"])
beams.check_for_zero(beams.getting_Volume(beams.beams_collector, ["Head"]), ["Head"])
beams.check_for_zero(beams.getting_Volume(beams.beams_collector, ["Concrete"]), ["Concrete"])
beams.check_for_zero(beams.getting_Count(beams.beams_collector, ["Precast"]), ["Precast"])

