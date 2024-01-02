import pandas as pd
import numpy as np
from pandasgui import show as show_df
from os import walk, getcwd, system
from colorama import Fore
from glob import glob
from routines import format_excel, ProjDetails
from tqdm import tqdm


class DisplayParams:
    def __init__(self, project: str, phase: str):
        self.start(project, phase)

    def getParam(self, line, param, log, sep=":", end=" "):
        try:
            # print(line)
            pos_start = line.index(param)
            # print(pos_start)
            pos_sep = line[pos_start:].index(sep)
            # print(pos_sep)
            pos_end = line[pos_sep + pos_start :].index(end)
            # print(pos_end)
            return line[pos_start + pos_sep + len(sep) : pos_end + pos_start + pos_sep]
        except:
            if param != "Point?tagname":
                log.append(f"exception found for param {param}. Line: {line}")
            return ""

    def getTitle(self, line: str, log):
        return self.getParam(line, "<TITLE", log, ">", "</TITLE>")

    def createParams(self, line: str, tagname, log):
        params = {}
        params["id"] = self.getParam(line, "id", log, "=")
        params["LEFT"] = self.getParam(line, "LEFT", log, ": ", "; ")
        params["TOP"] = self.getParam(line, "TOP", log, ": ", "; ")
        params["src"] = self.getParam(line, "src", log, ' = "', '" ')
        test = params["src"]
        params["display"] = self.getParam(test, ".", log, "\\", "_files")
        params["shape"] = self.getParam(test, ".", log, "_files\\", ".sha").upper()
        params["src"] = params["src"].split("\\")[-1]

        temp = self.getParam(line, "parameters", log, ' = "', '" ')
        temp = temp.replace("&amp;", "&")
        temp = temp.replace("&#10;", "")
        temp1 = temp.split(";")
        temp2 = {}
        params["Tagname"] = tagname
        params["Point?tagname"] = self.getParam(line, "Point?tagname", log, ":", ";")
        for param in temp1:
            if param == "":
                break
            try:
                param_name = param[param.index("?") + 1 : param.index(":")]
                param_value = param[param.index(":") + 1 :]
                params[param_name] = param_value
                # params['parameters'] = temp2
            except:
                print(temp1)
        # params["Line"] = line

        return params

    def write_Overview(self, folder_displays, file_output, project, phase):
        log = []
        if project != "Test":
            my_proj = ProjDetails(project)
            my_path = my_proj.path
            df_tags = pd.read_excel(
                my_path + f"EB\\{phase}\\" + f"{project}_export_EB_total_{phase}.xlsx"
            )
        else:
            df_tags = pd.read_excel(
                "Test\\" + f"{project}_export_EB_total_{phase}.xlsx"
            )

        tag_list = df_tags["&N"]
        print(f"{Fore.YELLOW}{len(tag_list)} tags found{Fore.RESET}")

        filenames = glob(folder_displays + "\\*.htm")
        print(f"{Fore.YELLOW}{len(filenames)} displays found{Fore.RESET}")

        total = {}

        print(f"{Fore.GREEN}Processing displays:{Fore.RESET}")
        tag_disp = []

        for file in tqdm(filenames):
            # with open(display_dir + "\\" + file, "r") as f:
            with open(file, "r") as f:
                text = f.readlines()
                found = ""
                title = ""
                displayname = file.split("\\")[-1]
                for line in text:
                    if title == "":
                        if "<TITLE>" in line:
                            title = self.getTitle(line, log)
                    if found != "":
                        found += line
                        if '">' in line:
                            found = found.replace("\r", "").replace("\n", "")
                            found = found.split('">', 1)[0]
                            for tagname in tag_list:
                                if tagname in found:
                                    result = self.createParams(found, tagname, log)
                                    if result["shape"] in total:
                                        total[result["shape"]].append(result)
                                    else:
                                        total[result["shape"]] = [result]
                                    tag_disp.append(
                                        [
                                            displayname,
                                            tagname,
                                            title,
                                            result["shape"],
                                            f'{result["LEFT"]},{result["TOP"]}',
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
            df[shape] = pd.DataFrame(total[shape])
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
            df_shape, orient="index", columns=["Shape", "Sheet", "#occurances"]
        )

        df_shape["Sheet"] = (
            "=HYPERLINK("
            + chr(34)
            + "#"
            + df_shape["Sheet"].astype(str)
            + "!A1"
            + chr(34)
            + ","
            + chr(34)
            + df_shape["Sheet"].astype(str)
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
                "Source",
                "id",
            ],
        )
        tag_disp.drop_duplicates(inplace=True)

        print(f"{Fore.YELLOW}Writing to Excel...{Fore.RESET}")
        if project == "Test":
            filename = "Test\\Display_params.xlsx"
        else:
            filename = f"Projects\\{project}\\{file_output}"

        with pd.ExcelWriter(filename, mode="w") as writer:
            df_shape.to_excel(writer, sheet_name="Index", index=False)
            tag_disp.to_excel(writer, sheet_name="Overview", index=False)
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
                pd.DataFrame(log).to_excel(writer, sheet_name="Exceptions", index=False)
        print(f"{Fore.YELLOW}Formatting Excel...{Fore.RESET}")
        if project == "Test":
            format_excel(f"Test\\", "Display_params.xlsx")
        else:
            format_excel(f"Projects\\{project}\\", file_output)

        return

    def start(self, project, phase):
        print(
            f"{Fore.MAGENTA}Creating display overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )

        if project == "Test":
            Export_display_folder = "Test\\Displays"
            Export_output_file = "Display_Params.xlsx"
        else:
            Export_display_folder = f"Projects\\{project}\\Displays\\{phase}"
            Export_output_file = f"{project}_Display_Params_{phase}.xlsx"

        # self.get_params_list(EB_path)
        self.write_Overview(Export_display_folder, Export_output_file, project, phase)
        print(
            f"{Fore.MAGENTA}Finished creating display overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )


def main():
    system("cls")
    # Requires: Excel file of EB of phase (via EBtoDB.py) or dummy file with column &N containing tag names
    #           name: {project}_export_EB_total_{phase}.xlsx
    # Requires: htm files in Projects\\{project}\\Displays\\{phase}
    project = DisplayParams("CEOD", "2023-08-03")


if __name__ == "__main__":
    main()
