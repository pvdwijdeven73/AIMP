# This program converts an EB file to a readable flat table.

# Import libraries
import os
import pandas as pd
import numpy as np
from routines import (
    format_excel,
    check_folder,
    file_exists,
    dprint,
    ProjDetails,
    assign_sensivity,
)
from os import system
from colorama import Fore, Back
from math import isnan

ERROR = Fore.WHITE + Back.RED
RESET = Fore.RESET + Back.RESET


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
            self.EBDB[sheet] = pd.DataFrame.from_dict(DB[sheet]).fillna("").copy()
            self.EBDB[sheet]["Source"] = self.source
            columns = sorted(self.EBDB[sheet].columns)
            columns.pop(columns.index("&T"))
            columns.pop(columns.index("&N"))
            columns.pop(columns.index("Source"))
            columns = ["Source", "&T", "&N"] + columns
            self.EBDB[sheet] = self.EBDB[sheet][columns]


class EBtoDB:
    def __init__(
        self,
        project: str,
        phase: str = "Original",
        open=True,
        pre_filter=False,
        AM=False,
    ):
        self.open = open
        self.start(project, phase, pre_filter, AM)

    def get_EB(
        self,
        EB_path,
        Export_output_path,
        EB_output_file,
        Export_output_file,
        project,
        phase,
        pre_filter,
    ):
        def func_assignPLC(row, boxes, UCN):
            pntbox = 0
            if "HWYNUM" in row.keys():
                if row["HWYNUM"] != "":
                    if "PNTBOXIN" in row:
                        try:
                            pntbox = int(row["PNTBOXIN"])
                        except:
                            pntbox = 0
                    if "PNTBOXOT" in row:
                        try:
                            pntbox = int(row["PNTBOXOT"])
                        except:
                            pntbox = pntbox

                    boxnum = 0
                    if "BOXNUM" in row:
                        try:
                            boxnum = int(row["BOXNUM"])
                        except:
                            boxnum = 0
                    if "OUTBOXNM" in row:
                        try:
                            boxnum = int(row["OUTBOXNM"])
                        except:
                            boxnum = boxnum

                    for box in boxes:
                        if (
                            boxnum == int(box[1])
                            and pntbox == int(box[2])
                            and int(row["HWYNUM"]) == int(box[3])
                        ):
                            return box[0]
            if "NODENUM" in row.keys():
                if row["NODENUM"] != "":
                    for node in UCN:
                        if int(row["NTWKNUM"]) == int(node[1]) and int(
                            row["NODENUM"]
                        ) == int(node[2]):
                            return node[0]

                    return "UCN"
            return ""

        if pre_filter:
            proj = ProjDetails(project)
            if phase == "Original":
                ph = phase
            else:
                ph = "Migrated"
            UCN_filter = []
            box_filter = []
            for PLC in proj.PLCs[ph]:
                if "HWParams" in PLC["details"]:
                    for i in range(len(PLC["details"]["HWParams"]["BOXES"])):
                        box_filter.append(
                            [
                                PLC["PLCName"],
                                PLC["details"]["HWParams"]["BOXES"][i],
                                PLC["details"]["HWParams"]["PNTBOXES"][i],
                                PLC["details"]["HWParams"]["HWNUM"],
                            ]
                        )
                elif "eUCNParams" in PLC["details"]:
                    UCN_filter.append(
                        [
                            PLC["PLCName"],
                            PLC["details"]["eUCNParams"]["UCN"]["UCN01"],
                            PLC["details"]["eUCNParams"]["NODENUM"],
                        ]
                    )

        EB = {}
        try:
            files = os.listdir(path=EB_path)
        except:
            dprint("No EB files found for original!", "RED")
            dprint(f"folder: {EB_path}", "RED")
            return

        pnttypes = []

        for file in files:
            if ".EB" in file.upper():
                dprint(f"- Loading:{file}", "CYAN")
                EB[file] = EBFile(EB_path, file)
                pnttypes += EB[file].get_point_types()
                EB[file].create_DB()

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

                columns = sorted(df_EB[sheet].columns)
                columns.pop(columns.index("&T"))
                columns.pop(columns.index("&N"))
                columns.pop(columns.index("Source"))
                columns = ["Source", "&T", "&N"] + columns
                df_EB[sheet] = df_EB[sheet][columns]

                if pre_filter:
                    dprint(f"- Filtering {sheet}", "GREEN")
                    df_EB[sheet]["PLC"] = df_EB[sheet].apply(
                        lambda row: func_assignPLC(row, box_filter, UCN_filter), axis=1
                    )
                    columns = list(df_EB[sheet].columns)
                    columns.pop(columns.index("PLC"))
                    columns = ["PLC"] + columns
                    df_EB[sheet] = df_EB[sheet][columns]

                    # remove tags without PLC
                    df_EB[sheet]["PLC"].replace("", np.nan, inplace=True)
                    df_EB[sheet] = df_EB[sheet].dropna(subset=["PLC"])

                dprint(f"- Writing converted EB {sheet}", "YELLOW")
                df_EB[sheet].to_excel(writer, sheet_name=sheet, index=False)
        if Export_output_file != "":
            with pd.ExcelWriter(
                Export_output_path + Export_output_file, engine="openpyxl", mode="w"
            ) as writer:
                dprint(f"- Writing DCS Export", "YELLOW")
                df_EB["TOTAL"].to_excel(writer, sheet_name="TOTAL", index=False)
        if self.open:
            dprint("- Formatting Excel", "YELLOW")
            format_excel(EB_path, EB_output_file, self.open)
        if Export_output_file != "":
            format_excel(Export_output_path, Export_output_file)

    def start(self, project, phase, pre_filter, AM):
        print(
            f"{Fore.MAGENTA}Creating EB files {'AM ' if AM else ''}for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )

        if project == "Test":
            if AM:
                EB_path = "Test\\EB\\Original\\AM\\"
                Export_output_path = "Test\\EB\\Original\\AM\\"
                EB_output_file = "Test_export_EB_AM.xlsx"
                Export_output_file = "Test_DCS_export_AM.xlsx"
            else:
                EB_path = "Test\\EB\\Original\\"
                Export_output_path = "Test\\EB\\Original\\"
                EB_output_file = "Test_export_EB_total_Original.xlsx"
                Export_output_file = "Test_DCS_export_TPStags.xlsx"

        else:
            if not AM:
                EB_path = f"Projects\\{project}\\EB\\{phase}\\"
                EB_output_file = f"{project}_export_EB_total_{phase}.xlsx"
            else:
                EB_path = f"Projects\\{project}\\EB\\{phase}\\AM\\"
                EB_output_file = f"{project}_export_EB_AM_{phase}.xlsx"
            Export_output_path = f"Projects\\{project}\\Exports\\"

            if phase == "Original" and self.open != False:
                Export_output_file = f"{project}_DCS_export_TPStags.xlsx"
            else:
                Export_output_file = ""

        self.get_EB(
            EB_path,
            Export_output_path,
            EB_output_file,
            Export_output_file,
            project,
            phase,
            pre_filter,
        )

        print(
            f"{Fore.MAGENTA}Done creating EB files {'AM ' if AM else ''}for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def get_eb_files(proj, phase, pre_filter=False, AM=False) -> pd.DataFrame:
    if proj == "Test":
        projpath = f"Test\\"
    else:
        projpath = f"Projects\\{proj}\\"
    if not AM:
        filename = projpath + f"EB\\{phase}\\{proj}_export_EB_total_{phase}.xlsx"
    else:
        filename = projpath + f"EB\\{phase}\\AM\\{proj}_export_EB_AM_{phase}.xlsx"
    if not file_exists(filename):
        dprint(f"{filename} not found, creating new files from EB", "GREEN")
        EBtoDB(proj, phase, open=False, pre_filter=pre_filter, AM=AM)
    try:
        dprint(f"- Loading {'AM' if AM else 'EB'} files {phase}", "CYAN")
        my_eb = pd.read_excel(
            filename,
            sheet_name=None,
        )
        sheet = "TOTAL"
        return my_eb["TOTAL"]
    except FileNotFoundError:
        print(f"{ERROR}ERROR: file {filename} not found{RESET}")
        exit(f"{ERROR}ABORTED: File not found{RESET}")


def main():
    system("cls")
    project = EBtoDB("RVC_AM", "Original", pre_filter=False, AM=True)


if __name__ == "__main__":
    main()
