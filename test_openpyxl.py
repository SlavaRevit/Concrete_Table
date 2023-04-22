import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Color
from openpyxl.utils import get_column_letter

# create a sample DataFrame
df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6], 'C': [7, 8, 9]})

# create an Excel writer object using openpyxl engine
writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')

# write DataFrame to worksheet
df.to_excel(writer, sheet_name='Test', index=False)

# modify cell A1 style
workbook = writer.book
worksheet = writer.sheets['Test']
cell = worksheet['A1']
cell.font = Font(color='FF0000')

# save the workbook
writer._save()