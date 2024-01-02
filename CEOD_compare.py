# import libraries
import pandas as pd
import numpy as np
from routines import read_db, dprint, quick_excel
import os
import re
from pandasgui import show as show_df


# constants
path = "projects/CEOD/"
outputpath = "projects/CEOD/Output/"
file_Yoko_export = "Yoko_Centum_CS3000_Parser_Output_3_3_2023_11_30_ 9_AM.xlsx"
file_overview = "IO_summary.xlsx"
file_params = "Hon_params.xlsx"
file_match = "Hon_Yok_match.xlsx"
Hon_path = "projects/CEOD/Hon/"

dict_Typical = {
    "AUTOMAN": {"file": "AUTOMAN_VALIDATION_1.xlsx", "sheet": "AUTOMAN_C300"},
    "DACA": {"file": "DAC_VALIDATION_3.xlsx", "sheet": "DACA_C300"},
    "DIGACQA": {"file": "DIGACQA_VALIDATION_4.xlsx", "sheet": "C300_DIGACQA"},
    "PID": {"file": "PID_VALIDATION_1.xlsx", "sheet": "PID_C300"},
    "DEVCTRL": {"file": "DEVCTRL_VALIDATION_1.xlsx", "sheet": "DEVCTL_C300"},
}

columns_IO_DB = [
    "FCS#",
    "YOKOGAWA POINT NAME",
    "Honeywell CM template Name ",
    "C300 CM NAME",
    "Is Priority Loop?",
    "Unit/Section",
    "Demolition Scope- Warmwater",
    "Demolition scope- ZZRef with no logic",
    "Card retain/ remove/ add",
]

sheets_IO_DB = ["CCC9CPM01", "U3000CPM03", "CEODCPM05"]


def import_Yoko_data():
    dprint("- Loading Yokogawa export...", "BLUE")
    data_Yoko = read_db(path, file_Yoko_export, sheet=None)

    return data_Yoko


def update_from_previous(df_final):
    dprint("- Loading previous match export...", "BLUE")
    df_previous = read_db(outputpath, file_match, sheet=0)
    merged_df = pd.merge(
        df_final,
        df_previous[["CM_NAME", "FBNAME", "Manual_Match"]],
        on=["CM_NAME", "FBNAME"],
        how="left",
    )

    def combine_tags(row):
        if row["Manual_Match"] == "":
            return row["CM_NAME"]
        else:
            return row["Manual_Match"]

    merged_df["match_tag"] = merged_df.apply(combine_tags, axis=1)

    df_final = pd.merge(
        df_final,
        merged_df[["CM_NAME", "FBNAME", "match_tag"]],
        on=["CM_NAME", "FBNAME"],
        how="left",
    )
    return merged_df


def import_overview():
    dprint("- Loading overview...", "BLUE")
    data_overview = read_db(outputpath, file_overview, "IO_summary")

    return data_overview


def import_params():
    dprint("- Loading params...", "BLUE")
    data_overview = read_db(outputpath, file_params)

    return data_overview


def match_source(file, sheet, data_overview):
    df_temp = read_db(path=Hon_path, filename=file, sheet=sheet)
    df_match = df_temp[["CM_NAME", "FBNAME"]]
    return df_match


def get_loopname(tagname):
    if pd.isnull(tagname):
        return ""
    pattern = re.compile(r"^(.*?)(\d)(DR)(\d)(.*?)$")
    matches = pattern.search(tagname)
    if matches:
        return ""
    else:
        pattern = re.compile(r"(\d{2})([A-Z])([A-Z]{0,3})(\d{3})")
        try:
            if tagname[2:4] == "DR":
                return ""
            matches = pattern.search(tagname)

            if matches:
                # Creating the loopname
                loopname = matches.group(1) + matches.group(2) + matches.group(4)
                return loopname
            # elif tagname[2:5] in ["GBN", "GBO", "KBS", "GZN", "GZO"]:
            #     return tagname[0:3] + tagname[5:8]
            else:
                return ""
        except TypeError:
            print(tagname)
            return ""


def match_loops(df_result, df_overview):
    def get_matching_pointnames(row):
        loopname = row["Loopname_CM"]
        if loopname == "":
            return ""
        matching_pointnames = df_overview.loc[
            df_overview["Loopname"] == loopname, "PointName"
        ]
        matching_pointnames = list(set(matching_pointnames))
        pointname_typical_map = df_overview.set_index("PointName")[
            "Yoko_typical"
        ].to_dict()
        matching_pointnames = [
            f"{pointname} [{pointname_typical_map.get(pointname)}]"
            for pointname in matching_pointnames
        ]
        return ", ".join(list(set(matching_pointnames)))

    df_result["Matching_Pointnames"] = df_result.apply(get_matching_pointnames, axis=1)
    df_result.loc[df_result["_merge"] == "both", "Matching_Pointnames"] = ""

    return df_result


def main():
    # Clear the screen
    os.system("cls" if os.name == "nt" else "clear")
    dprint("Importing data", "YELLOW")
    data_Yoko = import_Yoko_data()
    data_overview = import_overview()
    data_params = import_params()
    df_match = pd.DataFrame()
    for item in dict_Typical:
        df_match = pd.concat(
            [
                df_match,
                match_source(
                    dict_Typical[item]["file"],
                    dict_Typical[item]["sheet"],
                    data_overview,
                ),
            ],
            axis=0,
        )
    #! Make sure to combine with previous version to not lose manual tagging!
    df_match = update_from_previous(df_match)
    dprint("Matching data", "YELLOW")
    df_match = pd.merge(
        df_match,
        data_overview,
        left_on="match_tag",
        right_on="PointName",
        how="left",
        indicator=True,
    )
    dprint("Matching loopnames", "YELLOW")
    dprint("- creating CM loopnames", "YELLOW")
    df_match["Loopname_CM"] = df_match["CM_NAME"].apply(get_loopname)
    dprint("- Matching loopnames", "YELLOW")
    df_final = match_loops(df_match, data_overview)

    # quick_excel(df_final, outputpath, file_match, format=True, revision=True)
    # show_df(df_final)


if __name__ == "__main__":
    main()
