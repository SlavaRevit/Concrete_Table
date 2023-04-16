import os.path as op
import pandas as pd
import numpy as np


walls = {
    'Walls_IN' : (1120, 345),
    "Walls_OUT" : (1231, 220)
}

beams = {
'beam_head': 4.6250000013066677,
  'beam_prestressed': 19.250000005438466,
  'beam_precast': (3.562793560393863, 12),
  'regular_beams': 27.318750007718126,
  'beam_tooth': 2.5000000007063048,
  'beam_anchor': (7.8300000022120821, 18),
  'beam_foundation': 4.6250000013066632
}

Floors = {
    "regular" : (100, 10),
    "Prestressed" : (500, 130),
    "clsm" : (150 , 30),
    "lite-beton" : (500,100)
}
floors_data = open(r"C:\\Users\\gfhap\\Desktop\\testing_csv.csv")
print(floors_data)
df_walls = pd.DataFrame.from_dict(walls,orient="index", columns=['Area', 'Volume'])


#creating data_frama of beams from dict
df_beams = pd.DataFrame.from_dict(beams, orient="index", columns=["Value"])
# from tuples spliting result to 2 different columns, column1 = Volume, column2 = Count
df_beams[['Column1', 'Column2']] = df_beams['Value'].apply(lambda x: pd.Series([x[0], x[1]] if isinstance(x, tuple) else [x, None]))
df_beams = df_beams.drop(columns=['Value']).rename(columns={'Column1': 'Volume', 'Column2': 'Count'})

#creating data_frama of floors from dict
df_floors = pd.DataFrame.from_dict(Floors, orient='index', columns=['Area', 'Volume'])

#Inserting new Name of Type:
name = "Floors"
df_name = pd.DataFrame(index=["Floors"])
name_beams = "Beams"
df_name_beams = pd.DataFrame(index=['Beams'])
name_walls = "Walls"
df_name_walls = pd.DataFrame(index=['Walls'])

#concatenating this 2 tables
df = pd.concat([df_name_walls, df_walls, df_name_beams, df_beams, df_name, df_floors], axis=0, sort=False)
df = df.fillna('')
#changing position of columns
df = df.reindex(columns = ["Area", "Volume", "Count"])
writer = pd.ExcelWriter("file1.xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name="Test")
writer._save()
# print(df)

