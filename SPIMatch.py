# import libraries
import pandas as pd
import numpy as np
from simpledbf import Dbf5
from routines import (
    format_excel,
    show_df,
    ProjDetails,
    read_db,
    file_exists,
)  # ,ErrorLog
from colorama import Fore
import typing


class SPIMatch:
    def __init__(self, project: str, phase: str = "Original"):
        self.start(project, phase)

    def get_PLCs(self, input_path: str, PLCs: list[list]):

        lst_db_SM = []
        lst_db_FSC = []
        db_PLC = {}

        for PLC in PLCs:
            print(f"  - Reading {PLC[0]}")
            curDB = read_db(input_path, PLC[0])

            curDB.insert(0, "PLC", PLC[1])
            if curDB.columns[1] == "TYPE":
                lst_db_FSC.append(curDB)
                print("      FSC found")
            else:
                lst_db_SM.append(curDB)
                print("      SM found")

        if len(lst_db_SM) > 1:
            db_PLC["SM"] = pd.concat(lst_db_SM)
        elif len(lst_db_SM) == 1:
            db_PLC["SM"] = lst_db_SM[0]
        if len(lst_db_FSC) > 1:
            db_PLC["FSC"] = pd.concat(lst_db_FSC)
        elif len(lst_db_FSC) == 1:
            db_PLC["FSC"] = lst_db_FSC[0]
        return db_PLC

    def get_loopnames(self, masks, db_PLC) -> pd.DataFrame:
        def func_loopname(row, param):
            tag = row[param]
            for mask in masks:

                IPS_mask = mask[0]
                SPI_mask = mask[1]

                unit_IPS = IPS_mask[0 : IPS_mask.find("X")]
                unit_SPI = SPI_mask[0 : SPI_mask.find("X")]

                if unit_IPS == tag[0 : len(unit_IPS)]:
                    inst_pos = IPS_mask.find("X")
                    len_loop = SPI_mask.count("y")
                    loopnumber = ""
                    instrument = tag[inst_pos]
                    startpos = len(tag) - len_loop
                    while startpos >= 0:
                        if tag[startpos : startpos + len_loop].isnumeric():
                            loopnumber = tag[startpos : startpos + len_loop]
                            startpos = -1
                        else:
                            startpos -= 1

                    result = unit_SPI + instrument + loopnumber
                    return result

        db_temp = pd.DataFrame()  # preventing unbound error

        for PLC in db_PLC:
            if PLC == "FSC":
                params = {
                    "rack": "RACK",
                    "loc": "LOC",
                    "tag": "TAGNUMBER",
                    "type": "TYPE",
                }
            else:
                params = {
                    "rack": "ChassisID/IOTAName",
                    "loc": "Location",
                    "tag": "TagNumber",
                    "type": "PointType",
                }
            db_temp = db_PLC[PLC].copy()
            db_temp = db_temp[
                (db_temp[params["loc"]] != "SYS")
                & (db_temp[params["loc"]] != "CAB")
                & ((db_temp[params["rack"]] != 0) & (db_temp[params["rack"]] != ""))
            ]

            db_temp["Loop_match"] = db_temp.apply(
                lambda row: func_loopname(row, params["tag"]), axis=1
            )

        return db_temp

    def create_SPI_match(self, proj, phase) -> typing.Any:

        db_PLC = self.get_PLCs(proj.path + f"\\PLCs\\{phase}\\", proj.PLC_list[phase])

        if file_exists(proj.path + f"SPI_Match_{phase}.xlsx"):
            print(
                f"{Fore.RED}File already exists, please check user input for confirmation{Fore.RESET}"
            )
            check = input("File already exists, type 'OK' is you are sure to continue")
            if check.upper() != "OK":
                print("Aborted...")
                return
            else:
                print("Confirmed, continuing")

        db_match = self.get_loopnames(proj.SPIMatch, db_PLC)

        with pd.ExcelWriter(
            proj.path + f"SPI_Match_{phase}.xlsx", engine="openpyxl", mode="w"
        ) as writer:
            print(f"- Writing SPI match")
            db_match.to_excel(writer, sheet_name="SPI match", index=False)

        print("Formatting Excel")
        format_excel(proj.path, f"SPI_Match_{phase}.xlsx")
        # dp_SPI = read_db(proj.path + "exports\\", proj.SPIExport)

        print("done!")

    def start(self, project, phase):
        this_proj = ProjDetails(project)
        self.create_SPI_match(this_proj, phase)


def main():
    project = SPIMatch("RVC_MOD1", "Original")


if __name__ == "__main__":
    main()
