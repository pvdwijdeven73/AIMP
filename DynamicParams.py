import pandas as pd
import numpy as np
from pandasgui import show as show_df
from os import walk, getcwd, system
from colorama import Fore
from glob import glob
from routines import format_excel, ProjDetails
from tqdm import tqdm
import xml.etree.ElementTree as ET


class DisplayParams:
    def __init__(self, project: str, phase: str):
        self.start(project, phase)

    def analyse_line(self, line: str, tag: str):
        tag = 'PointRefPointName">' + tag
        results = []
        found = line.find(tag)
        while found > 0:

            print(found, line[found - 1 : found + len(tag) + 1])
            results.append(line[found - 1 : found + len(tag) + 1])
            found = line.find(tag, found + 1)
        return results

    def write_Overview(self, folder_displays, file_output, project, phase):

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

        tag_df = df_tags[
            (df_tags["&T"] == "DICMPNIM")
            & (df_tags["NODINPTS"] == 2)
            & (df_tags["NODOPTS"] == 0)
        ]["&N"]
        tag_list = tag_df.tolist()
        # print(tag_list)

        print(f"{Fore.YELLOW}{len(tag_list)} tags found{Fore.RESET}")
        filenames = glob(folder_displays + "\\*.htm")
        print(f"{Fore.YELLOW}{len(filenames)} displays found{Fore.RESET}")

        total = {}

        print(f"{Fore.GREEN}Processing displays:{Fore.RESET}")
        tag_disp = []

        for file in tqdm(filenames):
            try:
                with open(f"{file[:-4]}_files\\DS_datasource1.dsd", "r") as f:
                    tree = ET.parse(f"{file[:-4]}_files\\DS_datasource1.dsd")

                    text = f.readlines()
                    root = tree.getroot()
                    for child in root:
                        # Do something with the child element
                        # print(child.tag, child.attrib, child.text)
                        if len(child) > 0:
                            PVfound = False
                            Tagfound = ""
                            PresType = ""
                            for subchild in child:
                                if "PointRefParamName" in str(subchild.attrib):
                                    if str(subchild.text).upper() == "PV":
                                        PVfound = True
                                if "PointRefPointName" in str(subchild.attrib):
                                    Tagfound = str(subchild.text)
                                if "PresentationType" in str(subchild.attrib):
                                    PresType = str(subchild.text)
                            if PVfound and Tagfound != "":
                                if Tagfound in tag_list:
                                    # print(f"{file}: {Tagfound}")
                                    tag_disp.append(
                                        [file.split("\\")[-1], Tagfound, PresType]
                                    )

                    # found = ""
                    # title = ""
                    # displayname = file.split("\\")[-1]
                    # for line in text:
                    #     for tag in tag_list:
                    #         if tag in line:
                    #             tag_info = self.analyse_line(line, tag)
                    #             print(tag_info)
                    #             # print(f"tagname:{tag} - line found")
            except:
                print(f"no DS in {file[:-4]}_files\\")

        # print(f"{Fore.YELLOW}Formatting Excel...{Fore.RESET}")
        # if project == "Test":
        #     format_excel(f"Test\\", "Display_dynamic_params.xlsx")
        # else:
        #     format_excel(f"Projects\\{project}\\", file_output)

        print(tag_disp)

        return

    def start(self, project, phase):

        print(
            f"{Fore.MAGENTA}Creating dynamic parameter overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )

        if project == "Test":
            Export_display_folder = "Test\\Displays"
            Export_output_file = "Display_Dynamic_Params.xlsx"
        else:
            Export_display_folder = f"Projects\\{project}\\Displays\\{phase}"
            Export_output_file = f"{project}_Display_Dynamic_Params_{phase}.xlsx"

        # self.get_params_list(EB_path)
        self.write_Overview(Export_display_folder, Export_output_file, project, phase)
        print(
            f"{Fore.MAGENTA}Finished creating display overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )


def main():
    system("cls")
    # Requires: Excel file of EB of phase (via EBtoDB.py) or dummy file with column &N containing tag names
    #           name: {project}_export_EB_total_{phase}.xlsx
    # Requires: htm files in Projects\\{project}\\Displays\\{phase} including subfolders
    project = DisplayParams("SARU", "Final")


if __name__ == "__main__":
    main()
