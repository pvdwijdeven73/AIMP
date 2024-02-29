# This program converts an EB file to a readable flat table.

# Import libraries
import os
import pandas as pd
import numpy as np
from routines import format_excel, check_folder
from os import system
from colorama import Fore


class EBFile:
    def __init__(self, path, filename):
        self.source = filename
        self.filename = path + filename
        self.lines = ""
        self.pnttypes = []

    def read_file(self):
        with open(self.filename, "r") as EBtext:
            self.lines = EBtext.readlines()
            self.lines = [s.replace("\x00", "") for s in self.lines]

    def get_point_types(self):
        if self.lines == "":
            self.read_file()
        for x in self.lines:
            if "&T" in x:
                self.pnttypes.append(x[2:].strip())
        self.pnttypes = list(set(self.pnttypes))
        return self.pnttypes

    def print_file(self):
        for x in self.lines:
            print(f"'{x[:-1].strip()}'")

    def create_DB(self):
        if self.pnttypes == []:
            self.get_point_types()
        DB = {"TOTAL": []}
        for pntt in self.pnttypes:
            DB[pntt] = []
        cur_item = {}
        cur_type = ""
        for x in self.lines:

            if x[0] == "{":
                # new item detected
                if cur_item != {}:
                    # print(cur_item["&N"])
                    DB["TOTAL"].append(cur_item)
                    DB[cur_type].append(cur_item)
                    cur_item = {}
            else:
                if x[0] == "&":
                    cur_item[x[:2]] = x[2:].strip()
                    if x[:2] == "&T":
                        cur_type = x[2:].strip()
                else:
                    loc = x.find("=") + 1
                    cur_item[x[: loc - 1].strip()] = x[loc:].replace('"', "").strip()
        DB["TOTAL"].append(cur_item)
        # print(DB['TOTAL'])
        DB[cur_type].append(cur_item)
        self.EBDB = {}
        for sheet in DB.keys():
            self.EBDB[sheet] = pd.DataFrame.from_records(DB[sheet]).fillna("")
            self.EBDB[sheet]["Source"] = self.source
            columns = sorted(self.EBDB[sheet].columns)
            columns.pop(columns.index("&T"))
            columns.pop(columns.index("&N"))
            columns.pop(columns.index("Source"))
            columns = ["Source", "&T", "&N"] + columns
            self.EBDB[sheet] = self.EBDB[sheet][columns]


class EBperPLC:
    def __init__(self, project: str, phase: str = "Original"):
        self.start(project, phase)

    def get_EB(self, EB_path, Export_output_path, EB_output_file, Export_output_file):

        EB = {}
        files = os.listdir(path=EB_path)

        pnttypes = []

        for file in files:
            if ".EB" in file.upper():
                print(f"reading:{file}")
                EB[file] = EBFile(EB_path, file)
                pnttypes += EB[file].get_point_types()
                EB[file].create_DB()
                # print(EB[file].EBDB)

        pnttypes = list(set(pnttypes))
        pnttypes.sort()
        pnttypes = ["TOTAL"] + pnttypes
        df_EB = {}

        check_folder(Export_output_path)
        with pd.ExcelWriter(
            EB_path + EB_output_file, engine="openpyxl", mode="w"
        ) as writer:

            for sheet in pnttypes:
                df_EB[sheet] = pd.DataFrame()
                for file in files:
                    if ".EB" in file.upper():
                        if sheet in EB[file].EBDB:
                            df_EB[sheet] = pd.concat(
                                [df_EB[sheet], EB[file].EBDB[sheet]]
                            )
                print(f"- Writing converted EB {sheet}")
                df_EB[sheet].to_excel(writer, sheet_name=sheet, index=False)
        if Export_output_file != "":
            with pd.ExcelWriter(
                Export_output_path + Export_output_file, engine="openpyxl", mode="w"
            ) as writer:
                print(f"- Writing DCS Export")
                df_EB["TOTAL"].to_excel(writer, sheet_name="TOTAL", index=False)

        print("Formatting Excel")
        format_excel(EB_path, EB_output_file)
        if Export_output_file != "":
            format_excel(Export_output_path, Export_output_file)

        print("done!")

    def start(
        self,
        project,
        phase,
    ):

        print(
            f"{Fore.MAGENTA}Creating EB files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )

        if project == "Test":
            EB_path = "Test\\"
            Export_output_path = "Test\\"
            EB_output_file = "Test_EB_total.xlsx"
            Export_output_file = "Test_EB_TPStags.xlsx"
        else:
            EB_path = f"Projects\\{project}\\EB\\{phase}\\"
            Export_output_path = f"Projects\\{project}\\Exports\\"
            EB_output_file = f"{project}_export_EB_total_{phase}.xlsx"
            if phase == "Original":
                Export_output_file = f"{project}_DCS_export_TPStags.xlsx"
            else:
                Export_output_file = ""

        self.get_EB(EB_path, Export_output_path, EB_output_file, Export_output_file)

        print(
            f"{Fore.MAGENTA}Finished creating EB files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def main():
    system("cls")
    project = EBperPLC("PSU10_25", "Optim")


if __name__ == "__main__":
    main()
