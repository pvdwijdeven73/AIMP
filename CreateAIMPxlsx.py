# import libraries
import pandas as pd
import numpy as np
from simpledbf import Dbf5
from routines import (
    format_excel,
    assign_sensivity,
    ProjDetails,
    read_db,
    check_folder,
)  # ,ErrorLog
from colorama import Fore
from pathlib import Path
import win32com.client as win32
from win32com import __gen_path__
import typing
from time import sleep


class CreateAIMPxlsx:
    def __init__(self, project: str, phase: str, proj_date: str, author: str):
        self.start(project, phase, proj_date, author)

    def get_db(self, path: str, PLCs: list) -> pd.DataFrame:
        lst_db_FSC: list[pd.DataFrame] = []
        db_FSC: pd.DataFrame = pd.DataFrame()
        for curFile in PLCs:
            print(f"  - Reading {curFile[0]}")
            curDB = read_db(path, curFile[0])
            assert isinstance(curDB, pd.DataFrame)
            curDB.insert(0, "PLC", curFile[1])
            lst_db_FSC.append(curDB)
            if len(lst_db_FSC) > 1:
                db_FSC = pd.concat(lst_db_FSC)
            elif len(lst_db_FSC) == 1:
                db_FSC = lst_db_FSC[0]

        print("- Done reading files")

        return db_FSC

    def create_frontsheet(self, path, db, projectID, issue_date, author, title):
        def rgbToInt(rgb):
            colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
            return colorInt

        print(f"    - 5 second delay before creating front-sheet...")
        sleep(5)

        path = str(Path(path).resolve())
        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except AssertionError:  # AtributeError
            exit(
                f"{Fore.RED}Warning!!! |To format excel, first delete the following folder:{__gen_path__}{Fore.RESET}"
            )
        wb = excel.Workbooks.Open(path)
        excel.Visible = True
        first = True
        firstsheet = ""
        for sheet in excel.Worksheets:
            if first:
                firstsheet = sheet
                first = False
            sheet.Activate()
            sheet.UsedRange.HorizontalAlignment = win32.constants.xlLeft
            sheet.UsedRange.VerticalAlignment = win32.constants.xlTop
            sheet.Range(
                sheet.Cells(1, 1), sheet.Cells(1, sheet.UsedRange.Columns.Count)
            ).Interior.Color = 0
            sheet.Range(
                sheet.Cells(1, 1), sheet.Cells(1, sheet.UsedRange.Columns.Count)
            ).Font.Color = rgbToInt((255, 255, 255))
            sheet.Range(
                sheet.Cells(1, 1), sheet.Cells(1, sheet.UsedRange.Columns.Count)
            ).Font.Bold = True
            excel.ActiveSheet.Columns.AutoFilter(Field=1)
            excel.ActiveSheet.Columns.AutoFit()
            excel.Cells.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
        excel.Worksheets.Add()
        sheet = excel.Worksheets("Sheet1")
        wb.Worksheets("Sheet1").Move(Before=firstsheet)
        sheet.Select
        sheet.Name = "Frontpage"

        sheet.Activate()
        sheet.Cells(13, 5).Value = "Revision"
        sheet.Cells(13, 6).Value = "Date"
        sheet.Cells(13, 7).Value = "Author"
        sheet.Cells(14, 5).Value = "Rev0"
        sheet.Cells(14, 6).Value = issue_date
        sheet.Cells(14, 7).Value = author
        sheet.Cells(18, 5).Value = projectID
        sheet.Cells(18, 7).Value = title

        sheet.UsedRange.HorizontalAlignment = win32.constants.xlLeft
        sheet.UsedRange.VerticalAlignment = win32.constants.xlBottom
        sheet.Range(sheet.Cells(13, 5), sheet.Cells(13, 7)).Interior.Color = 0
        sheet.Range(sheet.Cells(13, 5), sheet.Cells(13, 7)).Font.Color = rgbToInt(
            (255, 255, 255)
        )
        sheet.Range(sheet.Cells(13, 5), sheet.Cells(13, 7)).Font.Bold = True
        sheet.Range(sheet.Cells(18, 7), sheet.Cells(18, 7)).Font.Size = 18
        excel.ActiveSheet.Columns.AutoFit()
        excel.Cells.Range("A2").Select()

        wb.Close(True)

    def create_aimp_excel(
        self,
        path: str,
        db: typing.Any,
        projectID: str,
        issue_date: str,
        author: str,
        title: str,
        params: list,
        new_params: list,
        sort_params: list,
    ) -> None:
        db = db[params].copy()
        for param in new_params:
            db[param] = ""

        db.sort_values(by=sort_params, inplace=True)

        filename = f"{path}{projectID} - {title} - working.xlsx"

        with pd.ExcelWriter(filename) as writer:
            print(f"  - Writing {title}")
            db.to_excel(writer, sheet_name=title, index=False)
        assign_sensivity(f"{path}", f"{projectID} - {title} - working.xlsx", True)

        self.create_frontsheet(filename, db, projectID, issue_date, author, title)

    def create_SOE_excel(
        self,
        path: str,
        db: typing.Any,
        projectID: str,
        issue_date: str,
        author: str,
        title: str,
        params: list,
        new_params: list,
        sort_params: list,
        PLCs: list,
    ) -> None:
        # TODO prepare know values

        def func_soe(row):
            if row["LOC"] == "FLD":
                if row["TYPE"] in ["AI", "AO"]:
                    return "FALSE"
                else:
                    return "TRUE"
            if row["LOC"] == "PNL":
                if row["TYPE"] == "I":
                    return "TRUE"
                else:
                    return "FALSE"
            if row["LOC"] == "COM":
                if row["TYPE"] in ["BI", "BO"]:
                    return "FALSE"
                elif row["TYPE"] == "DI":
                    return "TRUE"
                else:
                    return "TRUE (CHECK)"
            if row["LOC"] == "FSC":
                if row["TYPE"] in ["BI", "BO"]:
                    return "FALSE"
                elif row["TYPE"] == "DI":
                    return "TRUE"
                else:
                    return "FALSE"
            if row["LOC"] == "ANN":
                return "TRUE"
            if row["LOC"] == "":
                return "FALSE"

            return "FALSE (CHECK)"

        db = db[params].copy()
        for param in new_params:
            db[param] = ""

        db["SER_NEW"] = db.apply(lambda row: func_soe(row), axis=1)

        db.rename(columns={"SER": "SER_OLD"}, inplace=True)

        db.sort_values(by=sort_params, inplace=True)

        filename = f"{path}{projectID} - {title} - working.xlsx"
        with pd.ExcelWriter(filename) as writer:
            for PLC in PLCs:
                mask = db["PLC"] == PLC[1]
                print(f"  - Writing {title} - {PLC[1]}")
                db[mask].to_excel(writer, sheet_name=PLC[1], index=False)
        assign_sensivity(f"{path}", f"{projectID} - {title} - working.xlsx", True)

        self.create_frontsheet(filename, db, projectID, issue_date, author, title)

    def create_aimp_docs(
        self, project: str, proj_phase: str, proj_issue_date: str, proj_author: str
    ) -> None:
        my_proj = ProjDetails(project)

        proj_ID = my_proj.projectCode
        proj_path = my_proj.path
        proj_PLCs = my_proj.PLC_list[proj_phase]
        check_folder(f"{proj_path}AIMP\\")
        proj_db: pd.DataFrame = self.get_db(
            f"{proj_path}PLCs\\{proj_phase}\\", proj_PLCs
        )

        params = ["PLC", "TYPE", "TAGNUMBER", "SERVICE", "SHEET"]
        new_params = ["FAULTREACT", "Comment"]
        sort_params = ["PLC", "SHEET", "TAGNUMBER"]
        params_mask: pd.Series = proj_db.TYPE == "AI"
        self.create_aimp_excel(
            f"{proj_path}AIMP\\",
            proj_db[params_mask],
            proj_ID,
            proj_issue_date,
            proj_author,
            "AI Fault Reaction",
            params,
            new_params,
            sort_params,
        )

        params = ["PLC", "TYPE", "TAGNUMBER", "SERVICE", "SHEET"]
        new_params = [
            "Trip",
            "Tag trip value",
            "Comment",
            "Display",
            "Triptag DCS",
            "Status",
            "Comment status",
        ]
        sort_params = ["PLC", "SHEET", "TAGNUMBER"]
        params_mask: pd.Series = proj_db.TYPE == "AI"
        self.create_aimp_excel(
            f"{proj_path}AIMP\\",
            proj_db[params_mask],
            proj_ID,
            proj_issue_date,
            proj_author,
            "AI Trip Values",
            params,
            new_params,
            sort_params,
        )

        params = ["PLC", "TYPE", "TAGNUMBER", "SERVICE", "SHEET", "AENGUNIT"]
        new_params = ["EUNEW"]
        sort_params = ["PLC", "SHEET", "TAGNUMBER"]
        params_mask: pd.Series = proj_db.TYPE == "AI"
        self.create_aimp_excel(
            f"{proj_path}AIMP\\",
            proj_db[params_mask],
            proj_ID,
            proj_issue_date,
            proj_author,
            "EU Descriptors",
            params,
            new_params,
            sort_params,
        )

        params = [
            "PLC",
            "TYPE",
            "TAGNUMBER",
            "SERVICE",
            "QUALIFICAT",
            "LOC",
            "SHEET",
            "RACK",
            "POS",
            "CHAN",
        ]
        new_params = ["Force Enabled", "Comment"]
        sort_params = ["PLC", "SHEET", "TAGNUMBER"]
        self.create_aimp_excel(
            f"{proj_path}AIMP\\",
            proj_db,
            proj_ID,
            proj_issue_date,
            proj_author,
            "Force Enabled",
            params,
            new_params,
            sort_params,
        )

        params = ["PLC", "TYPE", "TAGNUMBER", "SERVICE", "LOC", "SHEET", "SER"]
        new_params = ["SER_NEW"]
        sort_params = ["PLC", "SHEET", "TAGNUMBER"]
        self.create_SOE_excel(
            f"{proj_path}AIMP\\",
            proj_db,
            proj_ID,
            proj_issue_date,
            proj_author,
            "Soe Enabled",
            params,
            new_params,
            sort_params,
            proj_PLCs,
        )

        print(f"{Fore.GREEN}Ready!!!{Fore.RESET}")

    def start(self, project, phase, proj_date, author):
        # project = "SWS6"
        # proj_phase = "Original"     # Original / TUV / SHUFFLE / FINAL
        # proj_issue_date = "2022-01-31"
        # proj_author = "Pascal van de Wijdeven"

        self.create_aimp_docs(project, phase, proj_date, author)


def main():
    project = CreateAIMPxlsx("GIRB50", "Original", "2024-03-13", "Akash Soerdjbalie")


if __name__ == "__main__":
    main()
