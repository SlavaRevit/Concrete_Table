import pandas as pd
import json

floors_data = open("C:\\Users\\gfhap\\Desktop\\Data_Floors", "r", encoding='utf-8')
beams_data = open("C:\\Users\\gfhap\\Desktop\\Data_Beams", "r", encoding='utf-8')

data_floors = "".join(floors_data.readlines())
data_beams = "".join(beams_data.readlines())

floors_dict = eval(data_floors)
beams_dict = eval(data_beams)

# print(beams_dict)
# print(type(floors_dict))

# skiping all None values
my_dict_floors = {k: v for k, v in floors_dict.items() if v is not None}
my_dict_beams = {k: [v] if not isinstance(v, list) else v for k, v in beams_dict.items() if v is not None}

# Creating DataFrame for Floors, Beams from dict
df_floors = pd.DataFrame.from_dict(my_dict_floors, orient="index", columns=["Area", "Volume"])
df_beams = pd.DataFrame.from_dict(my_dict_beams, orient="index", columns=["Volume", "Count"])

# rounding all numbers for floors
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
