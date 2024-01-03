import pandas as pd
import numpy as np

from datetime import datetime
from os import walk, getcwd, system, path
from colorama import Fore
from glob import glob
from routines import format_excel, ProjDetails
from tqdm import tqdm
from hashlib import md5


class CheckShapes:
    def __init__(self, project: str, phase: str):
        self.start(project, phase)

    def md5sum(self, filename):
        hash = md5()
        with open(filename, "rb") as f:
            for chunk in iter(lambda: f.read(128 * hash.block_size), b""):
                hash.update(chunk)
        return hash.hexdigest()

    def get_display(self, file_names, folder_displays):
        display_name = file_names.replace(folder_displays, "").replace("_files", "")
        return display_name[1:-1]

    def get_shape(self, file_name, folder_name):
        shape_name = file_name.replace(folder_name, "")
        return shape_name

    def get_shapes(self, project, phase):
        REV_SEARCH = '<parameter name="Revision" type="Num" description='

        folder_displays = f"Projects\\{project}\\Displays\\{phase}"
        folder_names = glob(folder_displays + "\\*\\", recursive=True)
        dict_shapes = {}
        dict_per_shape = {}
        for folder_name in tqdm(folder_names):
            display_name = self.get_display(folder_name, folder_displays)
            file_names = glob(folder_name + "\\*.sha", recursive=True)
            for file_name in file_names:
                with open(file_name, "r") as f:
                    rev = "Not found"
                    text = f.readlines()
                    for line in text:
                        pos = line.find(REV_SEARCH)
                        if pos > -1:
                            pos += len(REV_SEARCH)
                            rev = line[pos + 1 : line[pos + 1 :].find('"') + pos + 1]
                shape_name = self.get_shape(file_name, folder_name)
                checksum = self.md5sum(file_name)
                if shape_name not in dict_per_shape:
                    dict_per_shape[shape_name] = []
                dict_per_shape[shape_name].append("[" + checksum + "] _ " + rev)
                dict_shapes[display_name + "_" + shape_name] = {}
                dict_shapes[display_name + "_" + shape_name]["revision"] = rev
                dict_shapes[display_name + "_" + shape_name]["display"] = display_name
                dict_shapes[display_name + "_" + shape_name]["shape"] = shape_name
                dict_shapes[display_name + "_" + shape_name]["checksum"] = checksum
                # print(display_name, shape_name, self.md5sum(file_name))
        df_shapes = pd.DataFrame(dict_shapes).transpose()
        for shape in dict_per_shape:
            dict_per_shape[shape] = list(set(dict_per_shape[shape]))
        df_per_shape = pd.DataFrame.from_dict(dict_per_shape, orient="index")
        df_per_shape.reset_index(inplace=True)
        cols = ["shape"]
        for i in range(len(df_per_shape.columns)):
            cols.append("ver" + str(i))
        df_per_shape.columns = cols[:-1]

        return df_shapes, df_per_shape

    def start(self, project, phase):
        print(
            f"{Fore.MAGENTA}Creating shape check for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )

        if project == "Test":
            Export_display_folder = "Test\\Displays"
            Export_output_file = "Display_Params.xlsx"
        else:
            Export_display_folder = f"Projects\\{project}\\Displays\\{phase}"
            Export_output_file = f"{project}_Shapes_{phase}_check.xlsx"

        # self.write_Overview(Export_display_folder, Export_output_file, project, phase)

        df_shapes_CDA, df_per_shape_CDA = self.get_shapes(project, "CDA")

        df_shapes, df_per_shape = self.get_shapes(project, phase)

        df_per_shape["num_versions"] = df_per_shape.apply(
            lambda row: row.count() - 1, axis=1
        )
        cols = list(df_per_shape.columns.values)
        cols.remove("shape")
        cols.remove("num_versions")
        cols = ["shape", "num_versions"] + cols
        df_per_shape = df_per_shape[cols]

        df_match = df_shapes.merge(
            df_shapes_CDA, on="shape", how="left", suffixes=("_disp", "_lib")
        )
        df_shapes_CDA = df_shapes_CDA[["shape", "display", "revision", "checksum"]]
        df_shapes_CDA.columns = ["shape", "folder", "revision", "checksum"]
        df_shapes = df_shapes[["shape", "display", "revision", "checksum"]]
        df_CDA = df_match[df_match["checksum_lib"].notna()]
        rev = df_shapes_CDA["revision"].mode()[0]

        # df_CDA["error"] = np.where(
        #     df_CDA["checksum_disp"] == df_CDA["checksum_lib"],
        #     "OK",
        #     "Incorrect version of template",
        # )

        def func_error(row, rev):
            if row["checksum_disp"] == row["checksum_lib"]:
                return "OK"
            elif len(str(row["checksum_lib"])) > 5:
                return "Incorrect version of template"
            elif row["shape"][:3].upper() == "CDA":
                return "Not found in CDA templates!!!"
            elif int(row["num_versions"]) == 1:
                try:
                    if rev in row["revision_disp"]:
                        return "OK"
                    else:
                        return f"Only 1 version, but not revision {rev}"
                except:
                    return "DONT KNOW WHAT TO DO HERE"
            else:
                return f"multiple revisions ({int(row['num_versions'])}), to be investigated"

        df_result1 = df_shapes.merge(df_per_shape, on="shape", how="left")
        df_result = df_result1.merge(
            df_shapes_CDA, on="shape", how="left", suffixes=("_disp", "_lib")
        )
        df_result["error"] = df_result.apply(lambda row: func_error(row, rev), axis=1)
        df_result = df_result[
            [
                "shape",
                "display",
                "revision_disp",
                "checksum_disp",
                "revision_lib",
                "checksum_lib",
                "num_versions",
                "error",
            ]
        ]
        with pd.ExcelWriter(
            f"Projects\\{project}\\{project}_Shapes_{phase}_check.xlsx", mode="w"
        ) as writer:
            df_result.to_excel(writer, sheet_name="Overview", index=False)
            df_shapes_CDA.to_excel(writer, sheet_name="Template_overview", index=False)
            # df_result1.to_excel(writer, sheet_name="df_result1", index=True)
            # df_shapes.to_excel(writer, sheet_name="df_shapes", index=False)
            df_per_shape.to_excel(writer, sheet_name="Shapes_overview", index=False)
            # df_match.to_excel(writer, sheet_name="shapes_match", index=True)
            # df_CDA.to_excel(writer, sheet_name="df_CDA", index=False)

        format_excel(f"Projects\\{project}\\", f"{project}_Shapes_{phase}_check.xlsx")
        return

        print(
            f"{Fore.MAGENTA}Finished creating shapes overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )


def main():
    system("cls")
    # Requires: display folders ("_files") in Projects\\{project}\\Displays\\{phase}
    project = CheckShapes("CEOD", "2024-01-03")


if __name__ == "__main__":
    main()
