# import libraries
import pandas as pd
import numpy as np
from simpledbf import Dbf5
from sqlalchemy import all_
from routines import format_excel, ProjDetails, read_db  # ,ErrorLog
from colorama import Fore
from pandasgui import show as show_df
from tqdm import tqdm

Export_output_path = f"Test\\"
Export_input_file = f"SARU_COMM.xlsx"
Export_output_file = f"SARU_COMM_done___.xlsx"
all_sheets = pd.read_excel(
    f"{Export_output_path}{Export_input_file}", sheet_name=None, header=1
)

df_EB = pd.DataFrame()
for sheet in all_sheets:
    if sheet != "Voorblad":
        all_sheets[sheet]["COMSYS"] = sheet

        df_EB = pd.concat([df_EB, all_sheets[sheet]])

with pd.ExcelWriter(
    Export_output_path + Export_output_file, engine="openpyxl", mode="w"
) as writer:
    print(f"- Writing DCS Export")
    df_EB.to_excel(writer, sheet_name="SARU", index=False)
