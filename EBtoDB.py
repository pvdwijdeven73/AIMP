# This program converts an EB file to a readable flat table.

# Import libraries
import os
from typing import Any
import pandas as pd
import numpy as np
from routines import (
    format_excel,
    check_folder,
    file_exists,
    dprint,
    ProjDetails,
)
from os import system
from colorama import Fore, Back

ERROR = Fore.WHITE + Back.RED
RESET = Fore.RESET + Back.RESET


class ReadEBFile:
    def __init__(self, path, filename) -> None:
        # initialize class and assign internal values
        self.source = filename
        self.filename = path + filename
        self.lines = ""
        self.point_types = []

    def read_file(self) -> None:
        # open EB files and read-lines.
        with open(file=self.filename, mode="r", encoding="utf8") as EB_text:
            self.lines = EB_text.readlines()
            # NOTE: the encoding to utf8 should take care of replacing \x00
            # if not: uncomment the following line
            # * self.lines = [s.replace("\x00", "") for s in self.lines]

    def get_point_types(self) -> list[Any]:
        # creates a list of all point types
        if self.lines == "":
            self.read_file()
        # scan every line for point type
        for x in self.lines:
            if "&T" in x:
                self.point_types.append(x[2:].strip())
        # remove duplicate point types
        self.point_types = list(set(self.point_types))
        return self.point_types

    def create_DB(self) -> None:
        # first retrieve list of point types
        if self.point_types == []:
            self.get_point_types()
        # create a dictionary of all point types
        # and also a key containing all point types combined
        dict_point_types = {"TOTAL": []}
        for point_type in self.point_types:
            dict_point_types[point_type] = []
        cur_item = {}
        cur_type = ""
        # scan the lines for new tags and its parameters
        for x in self.lines:
            # "{" indicates beginning of a new tag
            if x[0] == "{":
                if cur_item != {}:
                    # add previous item to "TOTAL"" and cur_type lists
                    dict_point_types["TOTAL"].append(cur_item)
                    dict_point_types[cur_type].append(cur_item)
                    cur_item = {}
            else:
                # only params starting with "&" are different from other params
                # example:
                # * &T DIGINHG
                # * &N 700TZ015
                if x[0] == "&":
                    cur_item[x[:2]] = x[2:].strip()
                    if x[:2] == "&T":
                        cur_type = x[2:].strip()
                else:
                    # other params are separated by "=" sign and spaces
                    # example:
                    # * PTDESC   ="STOOM RS7001/2B HOOGHOOG"
                    loc = x.find("=") + 1
                    cur_item[x[: loc - 1].strip()] = x[loc:].replace('"', "").strip()
        # add final item to "TOTAL" and cur_type lists
        dict_point_types["TOTAL"].append(cur_item)
        dict_point_types[cur_type].append(cur_item)
        # now create a dictionary with point types (and "TOTAL") as keys and dataframes as values
        self.EBDB = {}
        for sheet in dict_point_types.keys():
            self.EBDB[sheet] = (
                pd.DataFrame.from_records(data=dict_point_types[sheet])
                .fillna(value="")
                .copy()
            )
            # add source (the EB filename) and rename colums where necessary
            # also put source, tagname (&N) and point type (&T) in front, the other parameters follow alphabetically
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
    ) -> None:
        self.open = open
        self.start(project=project, phase=phase, pre_filter=pre_filter, AM=AM)

    def get_EB(
        self,
        EB_path,
        Export_output_path,
        EB_output_file,
        Export_output_file,
        project,
        phase,
        pre_filter,
    ) -> None:
        def func_assignPLC(row, boxes, UCN) -> Any:
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

        UCN_filter = []
        box_filter = []
        if pre_filter:
            proj = ProjDetails(project=project)
            if phase == "Original":
                ph = phase
            else:
                ph = "Migrated"

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
            dprint(text="No EB files found for original!", color="RED")
            dprint(text=f"folder: {EB_path}", color="RED")
            return

        pnttypes = []

        for file in files:
            if ".EB" in file.upper():
                dprint(text=f"- Loading:{file}", color="CYAN")
                EB[file] = ReadEBFile(path=EB_path, filename=file)
                pnttypes += EB[file].get_point_types()
                EB[file].create_DB()

        pnttypes = list(set(pnttypes))
        pnttypes.sort()
        pnttypes = ["TOTAL"] + pnttypes
        df_EB = {}

        check_folder(folder=Export_output_path)
        with pd.ExcelWriter(
            path=EB_path + EB_output_file, engine="openpyxl", mode="w"
        ) as writer:
            for sheet in pnttypes:
                df_EB[sheet] = pd.DataFrame()
                for file in files:
                    if ".EB" in file.upper():
                        if sheet in EB[file].EBDB:
                            df_EB[sheet] = pd.concat(
                                objs=[df_EB[sheet], EB[file].EBDB[sheet]]
                            )

                columns = sorted(df_EB[sheet].columns)
                columns.pop(columns.index("&T"))
                columns.pop(columns.index("&N"))
                columns.pop(columns.index("Source"))
                columns = ["Source", "&T", "&N"] + columns
                df_EB[sheet] = df_EB[sheet][columns]

                if pre_filter:
                    dprint(f"- Filtering {sheet}", color="GREEN")
                    df_EB[sheet]["PLC"] = df_EB[sheet].apply(
                        lambda row: func_assignPLC(
                            row=row, boxes=box_filter, UCN=UCN_filter
                        ),
                        axis=1,
                    )
                    columns = list(df_EB[sheet].columns)
                    columns.pop(columns.index("PLC"))
                    columns = ["PLC"] + columns
                    df_EB[sheet] = df_EB[sheet][columns]

                    # remove tags without PLC
                    df_EB[sheet]["PLC"].replace("", np.nan, inplace=True)
                    df_EB[sheet] = df_EB[sheet].dropna(subset=["PLC"])

                dprint(f"- Writing converted EB {sheet}", color="YELLOW")
                df_EB[sheet].to_excel(writer, sheet_name=sheet, index=False)
        if Export_output_file != "":
            with pd.ExcelWriter(
                Export_output_path + Export_output_file, engine="openpyxl", mode="w"
            ) as writer:
                dprint(text=f"- Writing DCS Export", color="YELLOW")
                df_EB["TOTAL"].to_excel(writer, sheet_name="TOTAL", index=False)
        if self.open:
            dprint(text="- Formatting Excel", color="YELLOW")
            format_excel(path=EB_path, filename=EB_output_file, first_time=self.open)
        if Export_output_file != "":
            format_excel(path=Export_output_path, filename=Export_output_file)

    def start(self, project, phase, pre_filter, AM) -> None:
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
            EB_path=EB_path,
            Export_output_path=Export_output_path,
            EB_output_file=EB_output_file,
            Export_output_file=Export_output_file,
            project=project,
            phase=phase,
            pre_filter=pre_filter,
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
    if not file_exists(filename=filename):
        dprint(text=f"{filename} not found, creating new files from EB", color="GREEN")
        EBtoDB(project=proj, phase=phase, open=False, pre_filter=pre_filter, AM=AM)
    try:
        dprint(text=f"- Loading {'AM' if AM else 'EB'} files {phase}", color="CYAN")
        my_eb = pd.read_excel(
            io=filename,
            sheet_name=None,
        )
        sheet = "TOTAL"
        return my_eb["TOTAL"]
    except FileNotFoundError:
        print(f"{ERROR}ERROR: file {filename} not found{RESET}")
        exit(code=f"{ERROR}ABORTED: File not found{RESET}")


def main() -> None:
    system(command="cls")
    project = EBtoDB(project="PGPMODC", phase="Original", pre_filter=False, AM=False)


if __name__ == "__main__":
    main()
