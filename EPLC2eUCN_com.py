# This program converts an EB file to a readable flat table.

# Import libraries
import os
import pandas as pd
import numpy as np
from routines import format_excel, check_folder, show_df, read_db
from os import system
from colorama import Fore


class EPLC2eUCN_com:
    def __init__(self, project: str, phase: str = "Optim", oldnew=[], link=["2", "B"]):
        self.start(project, phase, oldnew, link)

    def get_files(self, path, Export_output_path, output_file, phase, project):
        PLC_old = read_db(
            path + f"Original\\Merged\\", f"{project}_total_PLCs_Original.xlsx"
        )
        PLC_new = read_db(
            path + f"{phase}\\Merged\\", f"{project}_total_PLCs_{phase}.xlsx"
        )

        PLCs = PLC_old["PLC"].unique()

        return PLC_old, PLC_new, PLCs

    def get_old_list(self, PLC_old, PLCs, link):
        PLC_old_COM = {}
        for PLC in PLCs:
            # EPLCG or eUCN?
            curPLC = PLC_old[PLC_old["PLC"] == PLC]
            if len(curPLC["DCS_ADDR"].unique()) == 1:
                print(f"{Fore.YELLOW}{PLC} has link on EPLCG{Fore.RESET}")
                PLC_old_COM[PLC] = curPLC[["PLC", "TYPE", "TAGNUMBER", "ADDRESS"]][
                    (curPLC["COM"] == 2) & (curPLC["CHANNEL"] == "B")
                ]
            else:
                print(f"{Fore.YELLOW}{PLC} has link on eUCN{Fore.RESET}")
                PLC_old_COM[PLC] = curPLC[["PLC", "TYPE", "TAGNUMBER", "DCS_ADDR"]][
                    curPLC["DCS_ADDR"] != -1
                ]
        return PLC_old_COM

    def get_new_list(self, PLC_new, PLCs, oldnew):
        def func_COM_signal(row):
            for x in range(1, 11):
                if f"Master{x}" in row:
                    if row[f"Master{x}"] == "EUCN":
                        return row[f"PLCAddress{x}"]
            return ""

        PLC_new_COM = {}
        for PLC in PLCs:
            if PLC not in oldnew:
                oldnew[PLC] = PLC
            curPLC = PLC_new[PLC_new["PLC"] == oldnew[PLC]]
            cur_columns = curPLC.columns
            curPLC["COM_address"] = curPLC.apply(
                lambda row: func_COM_signal(row), axis=1
            )

            PLC_new_COM[PLC] = curPLC[["PLC", "PointType", "TagNumber", "COM_address"]][
                curPLC["COM_address"] != ""
            ]

        return PLC_new_COM

    # def match_tags(self, PLC_new_COM, PLC_old_COM, PLCs):
    #     def func_COM_signal(row):
    #         for x in range(1, 11):
    #             if f"Master{x}" in row:
    #                 if row[f"Master{x}"] == "EUCN":
    #                     return row[f"PLCAddress{x}"]
    #         return ""

    #     PLC_new_COM = {}
    #     for PLC in PLCs:

    #     return PLC_new_COM

    def start(
        self,
        project,
        phase,
        oldnew,
        link,
    ):

        print(
            f"{Fore.MAGENTA}Creating COM signal file for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )

        if project == "Test":
            path = "Test\\"
            Export_output_path = "Test\\"
            Export_output_file = "Test_COMsignals.xlsx"
        else:
            path = f"Projects\\{project}\\PLCs\\"
            Export_output_path = f"Projects\\{project}\\"
            Export_output_file = f"{project}_export_COMsignals_{phase}.xlsx"

        PLC_old, PLC_new, PLCs = self.get_files(
            path, Export_output_path, Export_output_file, phase, project
        )
        PLC_old_COM = self.get_old_list(PLC_old, PLCs, link)
        PLC_new_COM = self.get_new_list(PLC_new, PLCs, oldnew)
        # result = self.match_tags(PLC_old_COM, PLC_new_COM, PLCs)

        with pd.ExcelWriter(
            Export_output_path + Export_output_file, engine="openpyxl", mode="w"
        ) as writer:
            print(f"- Exporting COM signals")
            for PLC in PLCs:
                PLC_old_COM[PLC].to_excel(writer, sheet_name=f"old_{PLC}", index=False)
                PLC_new_COM[PLC].to_excel(writer, sheet_name=f"new_{PLC}", index=False)
        print("Formatting Excel")
        format_excel(Export_output_path, Export_output_file)

        print(
            f"{Fore.MAGENTA}Finished creating COM signal file for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def main():
    system("cls")
    oldnew = {"RDM10_10": "RVC_RDM10", "FHV_55": "RVC_FHV"}
    project = EPLC2eUCN_com("RVC_MOD1_RDM", "Final", oldnew)


if __name__ == "__main__":
    main()
