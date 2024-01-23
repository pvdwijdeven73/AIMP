from routines import format_excel, assign_sensivity
from time import sleep

path = f"C:\\Users\\p.vandewijdeven\\OneDrive - Shell\\PY\\AIMP\\Projects\\CEOD\\"
file = "CEOD_Shapes_2024-01-22_check.xlsx"
print("formatting..")
format_excel(path, file, sensititivy=False)
print("done..")
