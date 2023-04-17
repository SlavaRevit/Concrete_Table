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
