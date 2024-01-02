# import libraries
import pandas as pd
import numpy as np
import os
from simpledbf import Dbf5
import win32com.client as win32
from routines import format_excel, check_folder, read_db, file_exists, dprint
from colorama import Fore, Back

ERROR = Fore.WHITE + Back.RED
RESET = Fore.RESET + Back.RESET


class XLSMerge:
    def __init__(
        self,
        project: str,
        phase: str = "Original",
        isPLC: bool = True,
        open: bool = True,
    ):
        self.start(project, phase, isPLC, open)

    def get_files(self, input_path, output_file):
        all_files = []
        # all_files=[["1FGS6_63.xls","1FGS6_63"],["2FGS5_52.xls","2FGS5_52"],["11CD6.xls","11CD6"],["12CD6.xls","12CD6"],["13CD6.xls","13CD6"],["14CD6.xls","14CD6"]]

        if all_files == []:
            files = os.listdir(path=input_path)
            for file in files:
                if file != output_file and "." in file:
                    all_files.append([file, file.split(".", 1)[0]])

        return all_files

    def read_and_merge(
        self, all_files, input_path, output_path, output_file, isPLC, open
    ):
        lst_db_SM = []
        lst_db_FSC = []
        lst_db = []

        db_FSC = pd.DataFrame()
        db_SM = pd.DataFrame()
        db = pd.DataFrame()

        for curFile in all_files:
            curDB = read_db(input_path, curFile[0])

            if isPLC:
                curDB.insert(0, "PLC", curFile[1])
                if curDB.columns[1] == "TYPE":
                    lst_db_FSC.append(curDB)
                else:
                    lst_db_SM.append(curDB)
            else:
                curDB.insert(0, "File", curFile[1])
                lst_db.append(curDB)

        if isPLC:
            if len(lst_db_SM) > 1:
                db_SM = pd.concat(lst_db_SM)
            elif len(lst_db_SM) == 1:
                db_SM = lst_db_SM[0]
            if len(lst_db_FSC) > 1:
                db_FSC = pd.concat(lst_db_FSC)
            elif len(lst_db_FSC) == 1:
                db_FSC = lst_db_FSC[0]
        else:
            if len(lst_db) > 1:
                db = pd.concat(lst_db)
            elif len(lst_db) == 1:
                db = lst_db[0]

        check_folder(output_path)

        if isPLC:
            with pd.ExcelWriter(f"{output_path}{output_file}") as writer:
                if len(lst_db_FSC) > 0:
                    dprint("- Writing joined FSC", "YELLOW")
                    db_FSC.to_excel(writer, sheet_name="FSC", index=False)
                if len(lst_db_SM) > 0:
                    dprint("- Writing joined SM", "YELLOW")
                    db_SM.to_excel(writer, sheet_name="SM", index=False)
        else:
            with pd.ExcelWriter(f"{output_path}{output_file}") as writer:
                dprint("- Writing joined DB", "YELLOW")
                db.to_excel(writer, sheet_name="Merged Files", index=False)
        if open:
            dprint("- Formatting Excel", "YELLOW")

            format_excel(output_path, output_file, open)

    def start(self, project, phase, isPLC, open):
        print(
            f"{Fore.MAGENTA}Creating PLC files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )
        if project == "Test":
            input_path = "Test\\"
            output_path = "Test\\"
            output_file = "Test_PLCs_total.xlsx"
        else:
            input_path = f"Projects\\{project}\\PLCs\\{phase}\\"
            output_path = f"Projects\\{project}\\PLCs\\{phase}\\Merged\\"
            output_file = f"{project}_total_PLCs_{phase}.xlsx"

        self.read_and_merge(
            self.get_files(input_path, output_file),
            input_path,
            output_path,
            output_file,
            isPLC,
            open,
        )
        print(
            f"{Fore.MAGENTA}Done creating PLC files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def get_plc(proj, phase, plc_type="FSC") -> pd.DataFrame:
    projpath = f"Projects\\{proj}\\"
    if not file_exists(
        projpath + f"PLCs\\{phase}\\Merged\\{proj}_total_PLCs_{phase}.xlsx"
    ):
        result = XLSMerge(proj, phase, True, False)

    try:

        dprint(f"- Loading {phase} PLCs file ({plc_type})", "CYAN")

        return pd.read_excel(
            projpath + f"PLCs\\{phase}\\Merged\\{proj}_total_PLCs_{phase}.xlsx",
            sheet_name=plc_type,
        )
    except FileNotFoundError:
        print(
            f"{ERROR}ERROR: file "
            f"'{proj}_total_PLCs_{phase}.xlsx' not found in "
            f"'Projects\\{proj}\\PLCs\\{phase}\\Merged\\'"
        )
        exit(f"ABORTED: File not found{RESET}")
    except ValueError:
        return pd.DataFrame()


def main():
    project = XLSMerge("PGPMODC", "Original", True)


if __name__ == "__main__":
    main()
