# import libraries
from unittest import skip
import pandas as pd
import numpy as np

# from pandasgui import show as show_df
from os import walk, getcwd, system
from colorama import Fore, Back
from glob import glob
from routines import format_excel, ProjDetails, read_db, dprint, file_exists
from tqdm import tqdm
from EBtoDB import get_eb_files

ERROR = Fore.WHITE + Back.RED
RESET = Fore.RESET + Back.RESET


class TagFromAM:
    def __init__(self, project: str, phase: str, format: bool = True):
        self.start(project, phase, format)

    def getTagname(self, line: str):
        return line[15 : line.find("(")]

    def getParam(self, line: str, taglist):
        try:
            param = line.split("=")[0].strip()
            paramtag = line.split("=")[1].strip()
            tagname = ""
            for tag in taglist:
                if tag in paramtag:
                    tagname = tag
            return param, paramtag, tagname
        except:
            print(line)
            exit()

    def write_Overview(
        self,
        project,
        phase,
        EB_folder,
        my_path,
        Export_output_file,
        CLrefs,
        format=True,
    ):
        df_clrefs = read_db(my_path + "Exports\\", CLrefs)

        df_clrefs = df_clrefs[df_clrefs["RefType"] == "CL Block Attached To Tag"]

        # retrieve DCS files, from EB files
        df_EB = get_eb_files(project, phase, pre_filter=True)
        # tag_list = list(df_EB["&N"])
        if "PLC" not in df_EB.columns:
            df_EB["PLC"] = "N.A."
        taglist = df_EB[["&N", "&T", "PTDESC", "PLC"]].set_index("&N").T.to_dict("list")

        # curtype = ""
        # curdesc = ""
        # curname = ""
        # taglist = {}
        # for filename in filenames:
        #     with open(filename, "r") as file:
        #         lines = file.readlines()
        #         lines = [s.replace("\x00", "") for s in lines]
        #     for line in lines:
        #         if "{SYSTEM ENTITY " in line:
        #             curname = self.getTagname(line)
        #             taglist[curname] = []
        #             curtype = ""
        #             curdesc = ""
        #         elif "&T " in line:
        #             curtype = line[2:].strip()
        #             taglist[curname].append(curtype)
        #         elif "&N " in line:
        #             continue
        #         elif "PTDESC" in line:
        #             curdesc = line.split("=")[1].replace('"', "").strip()
        #             taglist[curname].append(curdesc)
        dprint(f"- info: {len(taglist)} HG/SMM tags found", "BLUE")

        filenames = glob(EB_folder + "\\AM\\*.EB")

        paramlist = []
        curtag = ""
        curtype = ""
        curdesc = ""
        for filename in filenames:
            curfile = filename.split("\\")[-1]
            with open(filename, "r") as file:
                lines = file.readlines()
                lines = [s.replace("\x00", "") for s in lines]
            for line in lines:
                if "{SYSTEM ENTITY " in line:
                    curtag = self.getTagname(line)
                    curtype = ""
                    curdesc = ""
                elif "&T " in line:
                    curtype = line[2:].strip()
                elif "&N " in line:
                    continue
                elif "PTDESC" in line:
                    curdesc = line.split("=")[1].replace('"', "").strip()
                else:
                    if any(tag in line for tag in taglist):
                        param, paramtag, tagname = self.getParam(line, taglist)
                        paramlist.append(
                            [
                                curfile,
                                curtag,
                                curtype,
                                curdesc,
                                param,
                                paramtag,
                                tagname,
                                taglist[tagname][0],
                                taglist[tagname][1],
                                taglist[tagname][2],
                            ]
                        )
        df_CDS = pd.DataFrame(
            paramlist,
            columns=[
                "Source",
                "CDS_tag",
                "CDS_tag_type",
                "CDS_tag_desc",
                "Param",
                "ParamValue",
                "HG/SMM tag",
                "HG/SMM tag type",
                "HG/SMM tag desc",
                "PLC",
            ],
        )

        result = df_CDS.merge(
            df_clrefs[["Object Name", "Input Object Name"]],
            how="left",
            left_on="CDS_tag",
            right_on="Object Name",
        )
        result.rename(columns={"Input Object Name": "CL_file"}, inplace=True)
        result.drop(columns=["Object Name"], inplace=True)

        with pd.ExcelWriter(
            my_path + Export_output_file, engine="openpyxl", mode="w"
        ) as writer:
            dprint(f"- Writing CDS Export", "YELLOW")
            result.to_excel(writer, sheet_name="CDS", index=False)
        if format:
            dprint("- Formatting Excel", "YELLOW")
            format_excel(my_path, Export_output_file)

        return

    def start(self, project, phase, format):
        print(
            f"{Fore.MAGENTA}Making AM CDS overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )

        if project == "Test":
            EB_folder = "Test"
            Export_output_file = "Test_CDS_params.xlsx"
            my_path = "Test\\"
            EB_folder = my_path + f"EB\\{phase}\\"
            CLExport = "Test_export_CL_refs.xlsx"
        else:
            my_proj = ProjDetails(project)
            my_path = my_proj.path
            EB_folder = my_path + f"EB\\{phase}\\"
            Export_output_file = f"{project}_CDS_Params_{phase}.xlsx"
            CLExport = f"{project}_export_CLrefs.xlsx"

        # self.get_params_list(EB_path)
        self.write_Overview(
            project, phase, EB_folder, my_path, Export_output_file, CLExport, format
        )
        print(
            f"{Fore.MAGENTA}Done creating AM overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )


def get_CDS_files(proj, phase) -> pd.DataFrame:
    projpath = f"Projects\\{proj}\\"
    filename = projpath + f"{proj}_CDS_params_{phase}.xlsx"
    if not file_exists(filename):
        TagFromAM(proj, phase, False)
    try:
        dprint(f"- Loading CDS parameter file {phase}", "CYAN")
        my_CDS = pd.read_excel(
            filename,
        )
        return my_CDS

    except FileNotFoundError:
        print(f"{ERROR}ERROR: file {filename} not found")
        exit(f"ABORTED: File not found{RESET}")


def main():
    system("cls")
    # Requires: EB files of HG or SMM tags
    # Requires: EB files in subfolder AM of AM tags
    # Requires: CL reference export ({project}_export_CLrefs.xlsx)
    project = TagFromAM("RVC_AM", "Original")


if __name__ == "__main__":
    main()
