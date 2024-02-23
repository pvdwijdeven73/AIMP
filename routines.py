import typing
from xmlrpc.client import Boolean
from isort import file
import win32com.client as win32
from win32com import __gen_path__
import pandas as pd
import numpy as np
from simpledbf import Dbf5
from pathlib import Path
from tqdm import tqdm
import json
from os.path import isdir
from os import makedirs, walk
from colorama import Fore, Back
from pathlib import Path
import xlwings as xw
from time import sleep
import glob


def revision_file(path, filename, current=False):
    orig_filename = filename
    if filename[-4:].upper() == "XLSX":
        filename = filename[:-5]
    files = glob.glob(f"{path}/*{filename}_*.*")

    if files:
        # Extract the revision numbers from the filenames
        revisions = [int(file.split("_")[-1].split(".")[0]) for file in files]

        # Find the highest revision number
        latest_revision = max(revisions)

    else:
        # No files found with the given filename pattern
        latest_revision = -1

    if current:
        if latest_revision == -1:
            return orig_filename
        else:
            return filename + "_{:04d}".format(latest_revision) + ".xlsx"
    else:
        return filename + "_{:04d}".format(latest_revision + 1)


def quick_excel(df, path, filename="Test", format=True, revision=False):
    """
    Write something to xlsx file:
    in case of dict, every dict key is written as sheet
    else only one sheet is written
    sheet(s) are formatted in the end.
    """

    if ".xlsx" in filename:
        filename = filename[:-5]
    sheetname = filename
    if revision:
        filename = revision_file(path, filename)
    with pd.ExcelWriter(
        f"{path}{filename}.xlsx",
        engine="openpyxl",
        mode="w",
    ) as writer:
        if type(df) is dict:
            for sheet in df.keys():
                dprint(f"Writing {sheet} dataframe to excel\n", "Yellow")
                df[sheet].to_excel(writer, sheet_name=sheet, index=False)
        else:
            dprint(f"Writing {filename} dataframe to excel\n", "Yellow")
            df.to_excel(writer, sheet_name=sheetname, index=False)
    if format:
        format_excel(
            path=path,
            filename=f"{filename}.xlsx",
            first_time=True,
            different_red=False,
            different_blue=False,
            check_existing_red=False,
        )


def dprint(text, color):
    """
    print text in GREEN, RED, BLUE, CYAN, YELLOW or IMPORTANT
    """
    if color.upper() == "GREEN":
        print(f"{Fore.GREEN}{text}{Fore.RESET}")
    elif color.upper() == "MAGENTA":
        print(f"{Fore.MAGENTA}{text}{Fore.RESET}")
    elif color.upper() == "BLUE":
        print(f"{Fore.BLUE}{text}{Fore.RESET}")
    elif color.upper() == "CYAN":
        print(f"{Fore.CYAN}{text}{Fore.RESET}")
    elif color.upper() == "RED":
        print(f"{Fore.RED}{text}{Fore.RESET}")
    elif color.upper() == "BLUE":
        print(f"{Fore.BLUE}{text}{Fore.RESET}")
    elif color.upper() == "YELLOW":
        print(f"{Fore.YELLOW}{text}{Fore.RESET}")
    elif color.upper() == "IMPORTANT":
        print(f"{Fore.BLACK}{Back.YELLOW}{text}{Fore.RESET}{Back.RESET}")


class ProjDetails:
    """
    Get project details from json file
    """

    def __init__(self, project: str):
        self.debug = False
        self.get_details(project)

    def get_details(self, project: str):
        with open(f"Projects\\{project}\\{project}.json") as file:
            self.data = json.load(file)
        if self.debug:
            print(self.data)

        def try_detail(key1: str, key2: str) -> typing.Any:
            try:
                result = self.data[key1][key2]
            except KeyError:
                result = ""
                # print(
                #     f"{Fore.YELLOW}Warning: [{key1}][{key2}] not found in JSON{Fore.RESET}"
                # )
            return result

        # general items
        self.path: str = try_detail("general", "path")
        self.projectCode: str = try_detail("general", "projectCode")
        self.projectDesc: str = try_detail("general", "projectDesc")

        # files
        self.outputFile: dict = try_detail("files", "outputFile")
        self.shuffleList: str = try_detail("files", "shuffleList")
        self.tagChangeList: str = try_detail("files", "tagChangeList")
        self.ioList: str = try_detail("files", "ioList")
        self.DCSExport: str = try_detail("files", "DCSExport")
        self.HMIExport: str = try_detail("files", "HMIExport")
        self.CLExport: str = try_detail("files", "CLExport")
        self.SPIExport: str = try_detail("files", "SPIExport")

        # SPI match
        if self.SPIExport != "":
            self.SPIMatch: list[list[str]] = []
            for unit in self.data["SPI match"]:
                self.SPIMatch.append([unit, self.data["SPI match"][unit]])

        self.PLCs: dict[str, list[typing.Any]] = {}
        self.PLC_list: dict[str, typing.Any] = {}
        self.PLCs_prep: list[typing.Any] = []
        params: dict[str, typing.Any]
        for phase in self.data["PLCs"]:
            self.PLCs[phase] = []
            self.PLC_list[phase] = []
            for PLC in self.data["PLCs"][phase]:
                cur_PLC: dict[str, typing.Any] = self.data["PLCs"][phase][PLC]
                self.PLCs[phase].append(cur_PLC)
                self.PLC_list[phase].append(
                    [cur_PLC["details"]["PLCFile"], cur_PLC["PLCName"], PLC]
                )
                if phase == "Original":
                    this_PLC: list[typing.Any] = [
                        cur_PLC["details"]["PLCFile"],
                        cur_PLC["PLCName"],
                    ]
                    if cur_PLC["details"]["CommType"] == "UCN":
                        # example: ["324503.xlsx","324503",0,0,39,["3","A"],[1]]
                        this_PLC.extend([0, 0])
                        params = cur_PLC["details"]["UCNParams"]
                        this_PLC.append(int(params["NODENUM"]))
                        this_PLC.append([params["COM"], params["CHAN"]])
                        UCNs: list[typing.Any] = []
                        for UCN in params["UCN"]:
                            UCNs.append(int(params["UCN"][UCN]))
                        this_PLC.append(UCNs)
                    elif cur_PLC["details"]["CommType"] == "HW":
                        # example: ["570042.xlsx","570042",'2','2',0,["2","B"]]
                        params = cur_PLC["details"]["HWParams"]
                        this_PLC.append(params["HWNUM"])
                        this_PLC.append(params["PNTBOXIN"])
                        this_PLC.append(0)
                        this_PLC.append([params["COM"], params["CHAN"]])
                    self.PLCs_prep.append(this_PLC)


class ErrorLog:
    """
    Create error log from dictionary
    """

    def __init__(self):
        self.log: dict[str, typing.Any] = {}
        self.readable = ""
        self.df_readable: list[typing.Any] = []

    def make_readable(self, req_key: str) -> None:
        self.readable = ""
        self.df_readable = []

        def get_tabs(amount: int) -> str:
            tabs: str = ""
            for _ in range(amount):
                tabs += "\t"
            return tabs

        def compile_log(d_log: dict[str, typing.Any], level: int):
            if d_log:
                for key in d_log:
                    if (level == 0 and key == req_key) or req_key == "" or level > 0:
                        self.readable += get_tabs(level) + f"{key}\n"
                        self.df_readable.append({level: key})
                        if d_log:
                            compile_log(d_log[key], level + 1)

        compile_log(self.log, 0)

    def show(self, req_key: str = ""):
        self.make_readable(req_key)
        print(self.readable)

    def get_df(self, req_key: str = "", fill_all: Boolean = False):
        self.make_readable(req_key)
        df_error_log: pd.DataFrame = pd.DataFrame(self.df_readable).fillna("")
        df_error_log.columns = ["PNTTYPE", "Item", "Description"]
        if fill_all:
            df_error_log = df_error_log.replace(r"^s*$", float("NaN"), regex=True)
            df_error_log[["PNTTYPE", "Item"]] = df_error_log[
                ["PNTTYPE", "Item"]
            ].fillna(method="ffill")
            df_error_log.dropna(inplace=True)
        return df_error_log


def assign_sensivity(path: str, file: str, auto: bool = True):
    """
    Adds sensitivty label "Internal" to excel file
    """
    with xw.App(visible=True) as app:
        wb = xw.Book(path + file)
        if auto:
            label = "d0cb1e24-a0e2-4a4c-9340-733297c9cd7c"
            labelinfo = wb.api.SensitivityLabel.CreateLabelInfo()
            labelinfo.AssignmentMethod = 2
            labelinfo.Justification = "init"
            labelinfo.LabelId = "d0cb1e24-a0e2-4a4c-9340-733297c9cd7c"
            wb.api.SensitivityLabel.SetLabel(labelinfo, labelinfo)

            wb.save()
            wb.close
            sleep(5)
        else:
            name = input(
                "Set sensitivitylabel, save and close workbook and press enter to continue: "
            )


def format_excel(
    path,
    filename,
    first_time=True,
    different_red=False,
    different_blue=False,
    check_existing_red=False,
    how="different",
    sensititivy=True,
):
    """
    Format excel file:
    - highlight headers
    - adjust column width
    - freeze panes
    for all sheets
    """

    def rgbToInt(rgb):
        colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
        return colorInt

    if sensititivy:
        dprint("Adding sensitivty label....", "GREEN")
        assign_sensivity(path, filename, False)
    dprint("Actually formatting....", "GREEN")
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
    for sheet in excel.Worksheets:
        sheet.Activate()
        sheet.UsedRange.HorizontalAlignment = win32.constants.xlLeft
        sheet.UsedRange.VerticalAlignment = win32.constants.xlTop
        sheet.Range(
            sheet.Cells(1, 1), sheet.Cells(1, sheet.UsedRange.Columns.Count)
        ).Interior.Color = 0
        if first_time:
            sheet.Range(
                sheet.Cells(1, 1), sheet.Cells(1, sheet.UsedRange.Columns.Count)
            ).Font.Color = rgbToInt((255, 255, 255))
        sheet.Range(
            sheet.Cells(1, 1), sheet.Cells(1, sheet.UsedRange.Columns.Count)
        ).Font.Bold = True
        excel.ActiveSheet.Columns.AutoFilter(Field=1)
        excel.ActiveSheet.Columns.AutoFit()
        excel.Cells.Range("B2").Select()
        excel.ActiveWindow.FreezePanes = True
        totfound = 0
        wtotfound = 0
        checkfound = False
        if different_red:
            if (how == "check" and "_check" in sheet.Name) or how == "different":
                print(f"Coloring 'Different' cells red for sheet {sheet.Name}")
                for column in tqdm(
                    range(1, sheet.UsedRange.Columns.Count + 1), "coloring..."
                ):
                    # for column in range(1,sheet.UsedRange.Columns.Count+1):
                    #     if column % 10 == 0:
                    #         print(f"column {column} of {sheet.UsedRange.Columns.Count+1}")
                    cell = sheet.Cells(1, column)
                    # print(str.(cell.Value).upper())
                    if "_different" in str(cell.Value) or "_check" in str(cell.Value):
                        # print(f"{cell.Value} contains checks")
                        checkfound = True
                        found = 0
                        wfound = 0
                        for row in range(2, sheet.UsedRange.Rows.Count + 1):
                            cell = sheet.Cells(row, column)
                            if how == "different":
                                if (
                                    "DIFFERENT" in str(cell.Value).upper()
                                    or "INCORRECT" in str(cell.Value).upper()
                                ):
                                    cell.Interior.Color = rgbToInt((255, 0, 0))
                                    found += 1
                            else:
                                if str(cell.Value).upper() != "OK":
                                    if "WARNING" in str(cell.Value).upper():
                                        cell.Interior.Color = rgbToInt((255, 191, 0))
                                        wfound += 1
                                    else:
                                        cell.Interior.Color = rgbToInt((255, 0, 0))
                                        found += 1

                        if found >= 1:
                            cell = sheet.Cells(1, column)
                            cell.Interior.Color = rgbToInt((255, 0, 0))
                            totfound += 1
                        elif wfound >= 1:
                            cell = sheet.Cells(1, column)
                            cell.Interior.Color = rgbToInt((255, 191, 0))
                            wtotfound += 1
        if checkfound:
            if totfound >= 1:
                sheet.Tab.Color = rgbToInt((255, 0, 0))
            elif wtotfound >= 1:
                sheet.Tab.Color = rgbToInt((255, 191, 0))
            else:
                sheet.Tab.Color = rgbToInt((0, 255, 0))
        if different_blue:
            print(f"Coloring 'Different' columns blue for sheet {sheet.Name}")
            for column in tqdm(
                range(1, sheet.UsedRange.Columns.Count + 1), "coloring..."
            ):
                #             for column in range(1,sheet.UsedRange.Columns.Count+1):
                #                 if column % 10 == 0:
                #                     print(f"Column {column} of {sheet.UsedRange.Columns.Count+1}")
                cell = sheet.Cells(1, column)
                if "_different" in str(cell.Value):
                    found = False
                    for row in range(2, sheet.UsedRange.Rows.Count + 1):
                        cell = sheet.Cells(row, column)
                        if cell.Interior.Color == rgbToInt((0, 176, 240)):
                            found = True
                            break
                    if found:
                        cell = sheet.Cells(1, column)
                        cell.Interior.Color = rgbToInt((0, 176, 240))
            print(f"Coloring 'Different' rows blue for sheet {sheet.Name}")
            for row in tqdm(range(1, sheet.UsedRange.Rows.Count + 1), "coloring..."):
                #             for row in range(1,sheet.UsedRange.Rows.Count+1):
                #                 if row % 10 == 0:
                #                     print(f"Row {row} of {sheet.UsedRange.Rows.Count+1}")
                found = False
                for column in range(2, sheet.UsedRange.Columns.Count + 1):
                    cell = sheet.Cells(row, column)
                    if cell.Interior.Color == rgbToInt((0, 176, 240)):
                        found = True
                        break
                if found:
                    cell = sheet.Cells(row, 2)
                    cell.Interior.Color = rgbToInt((0, 176, 240))
        if check_existing_red:
            print(f"Searching for red cells in sheet {sheet.Name}")
            for row in tqdm(range(1, sheet.UsedRange.Rows.Count + 1), "coloring..."):
                #             for row in range(1,sheet.UsedRange.Rows.Count+1):
                #                 if row % 10 == 0:
                #                     print(f"Row {row} of {sheet.UsedRange.Rows.Count+1}")
                for column in range(2, sheet.UsedRange.Columns.Count + 1):
                    cell = sheet.Cells(row, column)
                    if cell.Interior.Color == rgbToInt((255, 0, 0)):
                        cell = sheet.Cells(1, column)
                        rowval = str(cell.Value)
                        cell = sheet.Cells(row, 2)
                        colval = str(cell.Value)
                        print(f"red found in {rowval}, {colval} on sheet {sheet.Name}")
        excel.Worksheets[1].Activate()


def show_df(df: pd.DataFrame):
    """
    Quickly show complete dataframe
    """
    # more options can be specified also
    pd.options.display.max_rows = None  # type: ignore
    pd.options.display.max_columns = None  # type: ignore
    print(df)
    pd.reset_option("all")


def check_folder(folder: str) -> None:
    """
    Check if folder exists
    """
    if not isdir(folder):
        makedirs(folder)
        dprint(f"- Created folder : {folder}", "YELLOW")


def file_exists(filename: str) -> typing.Any:
    """
    Check if file exists
    """
    my_file = Path(filename)
    return my_file.is_file()


# This function imports the database, (.dbf or .xls/xlsx)
def read_db(
    path: str, filename: str, sheet: typing.Union[str, int, None] = 0
) -> dict | pd.DataFrame:
    """
    import database with correct coding
    """

    if sheet == None:
        sheets = " sheet: all"
    else:
        sheets = " sheet: " + str(sheet)

    # Check if revisions are applied, if so, take highest one
    filename = revision_file(path, filename, True)

    dprint(f"- Loading {path}{filename}{sheets}", "CYAN")
    if not file_exists(path + filename) and filename[-4:].upper() == "XLSX":
        filename = filename[:-1]
    try:
        df: typing.Any
        if filename[-4:].upper() == ".DBF":
            dbf: Dbf5 = Dbf5(path + filename, codec="ISO-8859-1")
            df = dbf.to_dataframe()
        elif filename[-4:].upper() == ".CSV":
            df = pd.read_csv(path + filename)
        else:
            df = pd.read_excel(path + filename, na_filter=False, sheet_name=sheet)
        if type(df) != dict:
            df = df.replace(np.nan, "", regex=True)
        else:
            for key, df_temp in df.items():
                df[key] = df_temp.replace(np.nan, "", regex=True)
        return df
    except PermissionError:
        print(f"{Fore.RED}{filename} in folder {path} is locked{Fore.RESET}")
        exit()
    except FileNotFoundError:
        print(f"{Fore.RED}{filename} in folder {path} is not found{Fore.RESET}")
        exit()
