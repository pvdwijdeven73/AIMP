from routines import format_excel, assign_sensivity
from time import sleep

path = f"C:\\Users\\p.vandewijdeven\\Desktop\\Yoko mig\\"
file = "Yoko_Centum_CS3000_Parser_Output_3_3_2023_11_30_ 9_AM.xlsx"
print("formatting..")
format_excel(path, file, sensititivy=False)
print("done..")
