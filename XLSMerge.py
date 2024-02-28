# import libraries
import pandas as pd
import os
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
    ) -> None:
        self.start(project=project, phase=phase, isPLC=isPLC, open=open)

    def get_files(self, input_path, output_file) -> list[str]:
        all_files = []
        files = os.listdir(path=input_path)
        for file in files:
            if file != output_file and "." in file:
                all_files.append([file, file.split(".", 1)[0]])
        return all_files

    def read_and_merge(
        self, all_files, input_path, output_path, output_file, isPLC, open
    ) -> None:
        lst_db_SM = []
        lst_db_FSC = []
        lst_db = []
        db_FSC = pd.DataFrame()
        db_SM = pd.DataFrame()
        db = pd.DataFrame()
        for curFile in all_files:
            curDB = read_db(path=input_path, filename=curFile[0])
            assert isinstance(curDB, pd.DataFrame)
            if isPLC:
                curDB.insert(loc=0, column="PLC", value=curFile[1])
                if curDB.columns[1] == "TYPE":
                    lst_db_FSC.append(curDB)
                else:
                    lst_db_SM.append(curDB)
            else:
                curDB.insert(loc=0, column="File", value=curFile[1])
                lst_db.append(curDB)
        if isPLC:
            if len(lst_db_SM) > 1:
                db_SM = pd.concat(objs=lst_db_SM)
            elif len(lst_db_SM) == 1:
                db_SM = lst_db_SM[0]
            if len(lst_db_FSC) > 1:
                db_FSC = pd.concat(objs=lst_db_FSC)
            elif len(lst_db_FSC) == 1:
                db_FSC = lst_db_FSC[0]
        else:
            if len(lst_db) > 1:
                db = pd.concat(objs=lst_db)
            elif len(lst_db) == 1:
                db = lst_db[0]
        check_folder(folder=output_path)
        if isPLC:
            with pd.ExcelWriter(path=f"{output_path}{output_file}") as writer:
                if len(lst_db_FSC) > 0:
                    dprint(text="- Writing joined FSC", color="YELLOW")
                    db_FSC.to_excel(writer, sheet_name="FSC", index=False)
                if len(lst_db_SM) > 0:
                    dprint(text="- Writing joined SM", color="YELLOW")
                    db_SM.to_excel(writer, sheet_name="SM", index=False)
        else:
            with pd.ExcelWriter(path=f"{output_path}{output_file}") as writer:
                dprint(text="- Writing joined DB", color="YELLOW")
                db.to_excel(writer, sheet_name="Merged Files", index=False)
        if open:
            dprint(text="- Formatting Excel", color="YELLOW")

            format_excel(path=output_path, filename=output_file, first_time=open)

    def start(self, project, phase, isPLC, open) -> None:
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
            all_files=self.get_files(input_path=input_path, output_file=output_file),
            input_path=input_path,
            output_path=output_path,
            output_file=output_file,
            isPLC=isPLC,
            open=open,
        )
        print(
            f"{Fore.MAGENTA}Done creating PLC files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def get_plc(proj, phase, plc_type="FSC") -> pd.DataFrame:
    projpath = f"Projects\\{proj}\\"
    if not file_exists(
        filename=projpath + f"PLCs\\{phase}\\Merged\\{proj}_total_PLCs_{phase}.xlsx"
    ):
        _result = XLSMerge(project=proj, phase=phase, isPLC=True, open=False)
    try:
        dprint(text=f"- Loading {phase} PLCs file ({plc_type})", color="CYAN")
        return pd.read_excel(
            io=projpath + f"PLCs\\{phase}\\Merged\\{proj}_total_PLCs_{phase}.xlsx",
            sheet_name=plc_type,
        )
    except FileNotFoundError:
        print(
            f"{ERROR}ERROR: file "
            f"'{proj}_total_PLCs_{phase}.xlsx' not found in "
            f"'Projects\\{proj}\\PLCs\\{phase}\\Merged\\'"
        )
        exit(code=f"ABORTED: File not found{RESET}")
    except ValueError:
        return pd.DataFrame()


def main() -> None:
    _project = XLSMerge(project="CVP_MOD11", phase="Original", isPLC=True)


if __name__ == "__main__":
    main()
