import pandas as pd
import numpy as np


from datetime import datetime
from os import system, path
from colorama import Fore
from glob import glob
from routines import format_excel, ProjDetails
from tqdm import tqdm


class DisplayParams:
    def __init__(self, project: str, phase: str) -> None:
        self.start(project=project, phase=phase)

    def get_param(self, line, param, log, sep=":", end=" ") -> str:
        try:
            pos_start = line.upper().index(param.upper())
            pos_sep = line[pos_start:].index(sep)
            pos_end = line[pos_sep + pos_start :].index(end)
            return line[pos_start + pos_sep + len(sep) : pos_end + pos_start + pos_sep]
        except:
            if param != "Point?tagname":
                log.append(f"exception found for param {param}. Line: {line}")
            return ""

    def get_title(self, line: str, log) -> str:
        return self.get_param(
            line=line, param="<TITLE", log=log, sep=">", end="</TITLE>"
        )

    def create_params(self, line: str, tagname, log) -> dict:
        params = {}
        params["id"] = self.get_param(line=line, param="id", log=log, sep="=")
        params["HEIGHT"] = self.get_param(
            line=line, param="HEIGHT", log=log, sep=": ", end="; "
        )
        params["WIDTH"] = self.get_param(
            line=line, param="WIDTH", log=log, sep=": ", end="; "
        )
        params["LEFT"] = self.get_param(
            line=line, param="LEFT", log=log, sep=": ", end="; "
        )
        params["TOP"] = self.get_param(
            line=line, param="TOP", log=log, sep=": ", end="; "
        )
        params["src"] = self.get_param(
            line=line, param="src", log=log, sep=' = "', end='" '
        )
        test = params["src"]
        params["display"] = self.get_param(
            line=test, param=".", log=log, sep="\\", end="_files"
        )
        params["shape"] = self.get_param(
            line=test, param=".", log=log, sep="_files\\", end=".sha"
        ).upper()
        params["src"] = params["src"].split(sep="\\")[-1]

        temp = self.get_param(
            line=line, param="parameters", log=log, sep=' = "', end='" '
        )
        temp = temp.replace("&amp;", "&")
        temp = temp.replace("&#10;", "")
        temp1 = temp.split(sep=";")
        temp2 = {}
        for param in temp1:
            if param == "":
                break
            try:
                param_name = param[param.index("?") + 1 : param.index(":")]
                param_value = param[param.index(":") + 1 :]
                params[param_name] = param_value
            except:
                print(temp1)

        return params

    def write_Overview(self, folder_displays, file_output, project, phase) -> None:
        log = []
        if project != "Test":
            my_proj = ProjDetails(project=project)
            my_path = my_proj.path
            try:
                df_tags = pd.read_excel(
                    io=my_path
                    + f"EB\\{phase}\\"
                    + f"{project}_export_EB_total_{phase}.xlsx"
                )
                df_tags_created = True
            except FileNotFoundError:
                print(
                    f"{Fore.RED}Warning: no taglist found, continue with empty taglist{Fore.RESET}"
                )
                df_tags_created = False
        else:
            try:
                df_tags = pd.read_excel(
                    io="Test\\" + f"{project}_export_EB_total_{phase}.xlsx"
                )
                df_tags_created = True
            except FileNotFoundError:
                print(
                    f"{Fore.RED}Warning: no taglist found, continue with empty taglist{Fore.RESET}"
                )
                df_tags_created = False

        if not df_tags_created:
            tag_list = []
        else:
            tag_list = df_tags["&N"]

        print(f"{Fore.YELLOW}{len(tag_list)} tags found{Fore.RESET}")

        filenames = glob(pathname=folder_displays + "\\*.htm")
        print(f"{Fore.YELLOW}{len(filenames)} displays found{Fore.RESET}")

        total = {}

        print(f"{Fore.GREEN}Processing displays:{Fore.RESET}")
        tag_disp = []

        dict_disp_dates = {}

        for file in tqdm(filenames):
            with open(file=file, mode="r") as f:
                text = f.readlines()
                found = ""
                title = ""
                display_name = file.split("\\")[-1]
                dict_disp_dates[display_name] = {}
                dict_disp_dates[display_name]["Display"] = display_name
                dict_disp_dates[display_name]["Modified"] = datetime.utcfromtimestamp(
                    path.getmtime(filename=file)
                ).strftime("%Y-%m-%d %H:%M:%S")
                for line in text:
                    if title == "":
                        if "<TITLE>" in line:
                            title = self.get_title(line=line, log=log)
                            dict_disp_dates[display_name]["Title_head"] = title
                    if found != "":
                        found += line
                        if '">' in line:
                            found = found.replace("\r", "").replace("\n", "")
                            found = found.split(sep='">', maxsplit=1)[0]
                            tag_dummy = True
                            for tagname in tag_list:
                                if tagname in found:
                                    result = self.create_params(
                                        line=found, tagname=tagname, log=log
                                    )
                                    if result["shape"] in total:
                                        total[result["shape"]].append(result)
                                    else:
                                        total[result["shape"]] = [result]
                                    tag_disp.append(
                                        [
                                            display_name,
                                            tagname,
                                            title,
                                            result["shape"],
                                            f'{result["LEFT"]},{result["TOP"]}',
                                            f'H:{result["HEIGHT"]},W:{result["WIDTH"]}',
                                            result["src"],
                                            result["id"],
                                        ]
                                    )
                                    tag_dummy = False
                            if tag_dummy:
                                result = self.create_params(
                                    line=found, tagname="", log=log
                                )
                                if (
                                    result["src"].upper()
                                    == "All_DspTitle_eoc_01.sha".upper()
                                ):
                                    dict_disp_dates[display_name]["Title_shape"] = (
                                        result["Title"]
                                    )
                                    dict_disp_dates[display_name]["Title_compare"] = (
                                        ""
                                        if (
                                            dict_disp_dates[display_name]["Title_shape"]
                                            == dict_disp_dates[display_name][
                                                "Title_head"
                                            ]
                                        )
                                        else "Different"
                                    )
                                if result["shape"] in total:
                                    total[result["shape"]].append(result)
                                else:
                                    total[result["shape"]] = [result]
                                tag_disp.append(
                                    [
                                        display_name,
                                        "",
                                        title,
                                        result["shape"],
                                        f'{result["LEFT"]},{result["TOP"]}',
                                        f'H:{result["HEIGHT"]},W:{result["WIDTH"]}',
                                        result["src"],
                                        result["id"],
                                    ]
                                )
                            found = ""
                    if "hsc.shape.1" in line:
                        found = line

        df = {}
        df_shape = {}
        id = 0
        sheets = []
        for shape in sorted(total):
            df[shape] = pd.DataFrame(data=total[shape]).drop_duplicates()
            shapetxt = shape
            if len(shapetxt) > 28:
                shapetxt = f"_{shapetxt[len(shapetxt) - 29 :]}"
                if shapetxt in sheets:
                    shapetxt = "x" + shapetxt[1:]
            if len(shapetxt) >= 1 and "\\" not in shape:
                df_shape[id] = [shape, shapetxt, df[shape].shape[0]]
                id += 1
                sheets.append(shapetxt)
        print(f"{Fore.YELLOW}{id} shapes found{Fore.RESET}")
        df_shape = pd.DataFrame.from_dict(
            data=df_shape, orient="index", columns=["Shape", "Sheet", "#occurances"]
        )

        df_shape["Sheet"] = (
            "=HYPERLINK("
            + chr(34)
            + "#"
            + df_shape["Sheet"].astype(dtype=str)
            + "!A1"
            + chr(34)
            + ","
            + chr(34)
            + df_shape["Sheet"].astype(dtype=str)
            + chr(34)
            + ")"
        )

        tag_disp = pd.DataFrame(
            tag_disp,
            columns=[
                "Display",
                "Tagname",
                "Display Title",
                "Shape",
                "Position",
                "Size",
                "Source",
                "id",
            ],
        )
        tag_disp.drop_duplicates(inplace=True)

        disp_overview = pd.DataFrame(data=dict_disp_dates).transpose()

        print(f"{Fore.YELLOW}Writing to Excel...{Fore.RESET}")
        if project == "Test":
            filename = "Test\\Display_params.xlsx"
        else:
            filename = f"Projects\\{project}\\{file_output}"

        with pd.ExcelWriter(path=filename, mode="w") as writer:
            df_shape.to_excel(excel_writer=writer, sheet_name="Index", index=False)
            disp_overview.to_excel(
                excel_writer=writer, sheet_name="DisplayList", index=False
            )
            tag_disp.to_excel(excel_writer=writer, sheet_name="Overview", index=False)

            sheets = []
            for shape in sorted(df):
                shapetxt = shape
                if len(shapetxt) > 28:
                    shapetxt = f"_{shapetxt[len(shapetxt) - 29 :]}"
                    if shapetxt in sheets:
                        shapetxt = "x" + shapetxt[1:]
                if len(shapetxt) >= 1 and "\\" not in shapetxt:
                    df[shape].rename(
                        columns={
                            "id": "=HYPERLINK("
                            + chr(34)
                            + "#Index!A1"
                            + chr(34)
                            + ","
                            + chr(34)
                            + "id"
                            + chr(34)
                            + ")"
                        },
                        inplace=True,
                    )

                    df[shape].to_excel(writer, sheet_name=shapetxt, index=False)
                    sheets.append(shapetxt)
            if len(log) > 0:
                pd.DataFrame(data=log).to_excel(
                    excel_writer=writer, sheet_name="Exceptions", index=False
                )
        print(f"{Fore.YELLOW}Formatting Excel...{Fore.RESET}")
        if project == "Test":
            format_excel(path=f"Test\\", filename="Display_params.xlsx")
        else:
            format_excel(path=f"Projects\\{project}\\", filename=file_output)

        return

    def start(self, project, phase) -> None:
        print(
            f"{Fore.MAGENTA}Creating display overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )

        if project == "Test":
            Export_display_folder = "Test\\Displays"
            Export_output_file = "Display_Params.xlsx"
        else:
            Export_display_folder = f"Projects\\{project}\\Displays\\{phase}"
            Export_output_file = f"{project}_Display_Params_{phase}_ALL.xlsx"

        self.write_Overview(
            folder_displays=Export_display_folder,
            file_output=Export_output_file,
            project=project,
            phase=phase,
        )
        print(
            f"{Fore.MAGENTA}Finished creating display overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )


def main() -> None:
    system(command="cls")
    # Optional: Excel file of EB of phase (via EBtoDB.py) or dummy file with column &N containing tag names
    #           name: {project}_export_EB_total_{phase}.xlsx
    # Required: htm files in Projects\\{project}\\Displays\\{phase}
    project = DisplayParams(project="CEOD", phase="Test")


if __name__ == "__main__":
    main()
