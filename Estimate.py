"""
Program delivers input for estimate:
1. # IO PLC
2. # COM signals HG
3. # CL files
4. # CDS files
5. # Displays/shapes
6. Essential functions to be retrieved from SGM
7. AIMP scope to be done
"""


import pandas as pd
from routines import format_excel, dprint, read_db, assign_sensivity
from EBtoDB import get_eb_files
from XLSMerge import get_plc
from TagFromAM import get_CDS_files
from IOCount import get_IO_count
from shutil import copyfile
from colorama import Fore
from os import system
from pathlib import Path
import win32com.client as win32
from win32com import __gen_path__


class Estimate:
    def __init__(
        self,
        project: str,
        phase: str = "Original",
    ):
        self.start(project, phase)

    def produce_overview(self, project, path, filename, noHG=False):
        def write_value(cursheet, param, value):
            for column in range(1, cursheet.UsedRange.Columns.Count + 1):
                cell = cursheet.Cells(2, column)
                if cell.Value == param:
                    cursheet.Cells(3, column).Value = value
            return

        def get_value(df, key, param):
            if key in df.index:
                return df.loc[key][param]
            return 0

        path = str(Path(path).resolve())
        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except AssertionError:  # AtributeError
            print(
                f"{Fore.RED}Warning!!! |To format excel, first delete the following folder:{__gen_path__}{Fore.RESET}"
            )
            exit()
        wb = excel.Workbooks.Open(path + "\\" + filename)
        excel.Visible = True
        cursheet = excel.Worksheets(1)

        # start writing values here

        # Project items
        write_value(cursheet, "Shell_ID", "NL??_0????")
        write_value(cursheet, "Description", project)

        # Number of PLCs - FSC
        if "PLC_original_FSC" in self.main:
            PLCsFSC = self.main["PLC_original_FSC"]["PLC"].nunique()
        else:
            PLCsFSC = 0
        write_value(cursheet, "PLCs_FSC2SM", PLCsFSC)

        # Number of PLCs - SM
        if "PLC_original_SM" in self.main:
            PLCsSM = self.main["PLC_original_SM"]["PLC"].nunique()
        else:
            PLCsSM = 0
        write_value(cursheet, "PLCs_SM2SM", PLCsSM)

        # Number of PLCs - Total
        write_value(cursheet, "PLCs_Total", PLCsSM + PLCsFSC)

        if PLCsFSC > 0:
            # Number of IOcards
            temp = self.main["IOCards"].set_index("TYPE")
            write_value(cursheet, "PLC_cards_AI", get_value(temp, "AI", "AMOUNT"))
            write_value(cursheet, "PLC_cards_AO", get_value(temp, "AO", "AMOUNT"))
            write_value(cursheet, "PLC_cards_DI", get_value(temp, "I", "AMOUNT"))
            write_value(cursheet, "PLC_cards_DO", get_value(temp, "O", "AMOUNT"))
            write_value(cursheet, "PLC_cards_Total", get_value(temp, "Total", "AMOUNT"))

            # Number of IO
            temp = self.main["IOCount"].set_index("TYPE")
            write_value(cursheet, "PLC_IO_AI", get_value(temp, "AI", "AMOUNT"))
            write_value(cursheet, "PLC_IO_AO", get_value(temp, "AO", "AMOUNT"))
            write_value(cursheet, "PLC_IO_DI", get_value(temp, "I", "AMOUNT"))
            write_value(cursheet, "PLC_IO_DO", get_value(temp, "O", "AMOUNT"))
            write_value(cursheet, "PLC_IO_Total", get_value(temp, "Total", "AMOUNT"))
        else:
            # Number of IOcards
            write_value(cursheet, "PLC_cards_AI", 0)
            write_value(cursheet, "PLC_cards_AO", 0)
            write_value(cursheet, "PLC_cards_DI", 0)
            write_value(cursheet, "PLC_cards_DO", 0)
            write_value(cursheet, "PLC_cards_Total", 0)

            # Number of IO
            write_value(cursheet, "PLC_IO_AI", 0)
            write_value(cursheet, "PLC_IO_AO", 0)
            write_value(cursheet, "PLC_IO_DI", 0)
            write_value(cursheet, "PLC_IO_DO", 0)
            write_value(cursheet, "PLC_IO_Total", 0)

        # Number of DCS tags
        if noHG:
            write_value(cursheet, "DCS_tags", 0)
        else:
            write_value(cursheet, "DCS_tags", self.main["DCS_original"]["&N"].nunique())

        # Number of DCS displays
        if noHG:
            write_value(cursheet, "DCS_displays", 0)
        else:
            write_value(
                cursheet, "DCS_displays", self.main["HMI_refs"]["Object Name"].nunique()
            )

        # Number of CL & CLrefs
        if noHG:
            write_value(cursheet, "DCS_CL", 0)
            write_value(cursheet, "DCS_CL_tags", 0)
        else:
            write_value(
                cursheet, "DCS_CL", self.main["CL_refs"]["Input Object Name"].nunique()
            )
            write_value(cursheet, "DCS_CL_tags", len(self.main["CL_refs"].index))

        # Number of CDS & CDSrefs
        if noHG:
            write_value(cursheet, "DCS_CDS", 0)
            write_value(cursheet, "DCS_CDS_tags", 0)
        else:
            write_value(cursheet, "DCS_CDS", self.main["CDS_refs"]["CDS_tag"].nunique())
            write_value(cursheet, "DCS_CDS_tags", len(self.main["CDS_refs"].index))

        return

    def start(self, project, phase):
        print(
            f"{Fore.MAGENTA}Creating estimate overview for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )

        self.main = {}
        self.projpath = f"Projects\\{project}\\"

        # retrieve PLC files, both FSC and SM if available
        self.main["PLC_original_FSC"] = get_plc(project, phase, "FSC")
        self.main["PLC_original_SM"] = get_plc(project, phase, "SM")

        # retrieve DCS files, from EB files
        try:
            self.main["DCS_original"] = get_eb_files(project, phase, pre_filter=True)
            temp = self.main["DCS_original"][
                self.main["DCS_original"]["&T"].str.contains("HG")
            ]
            tag_list = list(temp["&N"])
        except:
            tag_list = []

        noHG = len(tag_list) == 0

        # retrieve HMI reference file
        # {project}_export_HMIrefs is a DOC4000 export from a query:
        # Asset:                  ****_EPKS_******
        # Query Type:             References
        # For Objects of type:    HMIWebDisplay
        # Reference Types:        HMIWeb Display - EPKS Entity
        if not noHG:
            self.main["HMI_refs"] = read_db(
                self.projpath + "Exports\\", f"{project}_export_HMIrefs.xlsx"
            )

            self.main["HMI_refs"] = self.main["HMI_refs"][
                self.main["HMI_refs"]["Output Object Name"].isin(tag_list)
            ]

            # retrieve CL reference file
            self.main["CL_refs"] = read_db(
                self.projpath + "Exports\\", f"{project}_export_CLrefs.xlsx"
            )

            self.main["CL_refs"] = self.main["CL_refs"][
                self.main["CL_refs"]["Object Name"].isin(tag_list)
            ]

            # retrieve CDS parameters
            self.main["CDS_refs"] = get_CDS_files(project, phase)
        else:
            dprint("- No HG point found, skipping HMI, CL and CDS counts", "GREEN")

        if self.main["PLC_original_FSC"].empty:
            dprint("- No FSC found, skipping IO counts", "GREEN")
        else:
            df_IO = get_IO_count(project, phase)
            self.main["IOCards"] = df_IO["IO Cards Overview"]
            self.main["IOCount"] = df_IO["IO Counts Overview"]

        # remove empty dataframes
        for key in list(self.main.keys()):
            if self.main[key].empty:
                del self.main[key]

        # copy template file
        template_file = f"Projects\\Templates\\ProjectInfo_template.xlsx"
        overview_file = f"{project}_estimate_overview_{phase}.xlsx"
        copyfile(template_file, self.projpath + overview_file)
        assign_sensivity(self.projpath, overview_file)

        self.produce_overview(project, self.projpath, overview_file, noHG)

        # write details
        detail_file = f"{project}_estimate_{phase}.xlsx"
        with pd.ExcelWriter(self.projpath + detail_file, mode="w") as writer:
            for sheet in self.main:
                dprint(f"- Writing sheet {sheet}", "YELLOW")
                self.main[sheet].to_excel(writer, sheet_name=sheet, index=False)

        format_excel(
            self.projpath,
            detail_file,
            first_time=True,
        )

        print(
            f"{Fore.MAGENTA}Done creating estimate overview for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )

    # TODO check if CL refs are correct (EIP code?)


def main():
    system("cls")
    project = Estimate("PGPMODA_50", "Original")


if __name__ == "__main__":
    main()
