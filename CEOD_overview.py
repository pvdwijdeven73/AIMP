# import libraries
import pandas as pd
import numpy as np
from routines import read_db, dprint, quick_excel
import os
from pandasgui import show as show_df
import re

# constants
path = "projects/CEOD/"
outputpath = "projects/CEOD/Output/"
file_connections = "Yoko_connections.xlsx"
file_Typicals = "Typicals.xlsx"
file_IO_db = "SHELL_NLCH03041_ IO LIST_REV 4.xlsx"


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


def import_connections():
    dprint("- Loading Connections...", "BLUE")
    data_Yoko = read_db(outputpath, file_connections, sheet="Yoko_connections")

    return data_Yoko


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


def import_IO_data():
    dprint("- Loading IO database...", "BLUE")
    data_IO = {}
    for sheet in sheets_IO_DB:
        data_IO[sheet] = read_db(path, file_IO_db, sheet=sheet)

        # fix column mismatch in CPM05 sheet
        if "Demolition scope- ZZ Ref with no logic" in data_IO[sheet].columns.to_list():
            data_IO[sheet] = data_IO[sheet].rename(
                columns={
                    "Demolition scope- ZZ Ref with no logic": "Demolition scope- ZZRef with no logic"
                }
            )

        data_IO[sheet] = data_IO[sheet][columns_IO_DB]
    df_total = pd.concat(data_IO.values(), ignore_index=True)
    ignore_values_column1 = ["SPARE", "NA", ""]
    ignore_values_column2 = ["Yes"]
    ignore_values_column3 = ["Demolish"]

    # Add a new column based on the conditions
    df_total["Scope"] = [
        "No"
        if (x in ignore_values_column1)
        or (y in ignore_values_column2)
        or (z in ignore_values_column3)
        else "Yes"
        for x, y, z in zip(
            df_total["C300 CM NAME"],
            df_total["Demolition Scope- Warmwater"],
            df_total["Demolition scope- ZZRef with no logic"],
        )
    ]

    Yoko_Typicals = read_db(outputpath, file_Typicals, "Typicals")

    df_total = df_total.merge(
        Yoko_Typicals,
        how="outer",
        left_on="YOKOGAWA POINT NAME",
        right_on="PointName",
        indicator=True,
    )

    return df_total


def add_group_info(data_IO, connections):
    dprint("Adding group info", "YELLOW")
    result_list = [
        (value, column)
        for column, series in connections.iteritems()
        for value in series
    ]

    # Create a new DataFrame from the list of tuples
    result_df = pd.DataFrame(
        result_list, columns=["PointName", "Group"]
    ).drop_duplicates()
    result_df = result_df[result_df["PointName"] != ""]

    data_IO = pd.merge(
        data_IO,
        result_df,
        left_on="PointName",
        right_on="PointName",
        how="left",
    )

    data_IO["Loopname"] = data_IO["PointName"].apply(get_loopname)

    #    data_IO = data_IO.drop("PointName_y", axis=1)
    #    data_IO = data_IO.rename(columns={"PointName_x": "PointName_Yoko"})
    data_IO.columns = [
        "IODB_FCS#",
        "IODB_YOKOGAWA POINT NAME",
        "IODB_Honeywell CM template Name ",
        "IODB_C300 CM NAME",
        "IODB_Is Priority Loop?",
        "IODB_Unit/Section",
        "IODB_Demolition Scope- Warmwater",
        "IODB_Demolition scope- ZZRef with no logic",
        "IODB_Card retain/ remove/ add",
        "Scope",
        "Controller",
        "PointName",
        "DrawingName",
        "Tag_Comment",
        "Yoko_typical",
        "Hon_typical",
        "_merge",
        "Group",
        "Loopname",
    ]
    data_IO = data_IO[
        [
            "Controller",
            "PointName",
            "Scope",
            "Tag_Comment",
            "Loopname",
            "DrawingName",
            "Yoko_typical",
            "Hon_typical",
            "Group",
            "IODB_FCS#",
            "IODB_YOKOGAWA POINT NAME",
            "IODB_Honeywell CM template Name ",
            "IODB_C300 CM NAME",
            "IODB_Is Priority Loop?",
            "IODB_Unit/Section",
            "IODB_Demolition Scope- Warmwater",
            "IODB_Demolition scope- ZZRef with no logic",
            "IODB_Card retain/ remove/ add",
            "_merge",
        ]
    ]

    # Replace values in the "_merge" column
    data_IO["_merge"] = data_IO["_merge"].replace(
        {"both": "IODB&YokoExport", "left_only": "IODB", "right_only": "YokoExport"}
    )

    # Rename the column to "Source"
    data_IO = data_IO.rename(columns={"_merge": "Source"})

    data_IO = data_IO.sort_values(
        by=["Controller", "PointName", "IODB_FCS#", "IODB_YOKOGAWA POINT NAME"]
    )

    group_scope = data_IO.groupby("Group")["Scope"].apply(
        lambda x: "Yes (from other IO in group)"
        if "Yes" in x.values
        else (
            "No (from other IO in group)"
            if "No" in x.values
            else "Unknown (not in IODB)"
        )
    )
    data_IO["Scope"] = data_IO["Scope"].fillna(data_IO["Group"].map(group_scope))
    data_IO["Scope"] = data_IO["Scope"].fillna("Unknown (no Group, not in IODB)")

    return data_IO


def main():
    # Clear the screen
    os.system("cls" if os.name == "nt" else "clear")
    dprint("Importing data", "YELLOW")
    connections = import_connections()
    data_IO = add_group_info(import_IO_data(), connections)
    quick_excel(data_IO, outputpath, "IO_summary", format=True, revision=True)
    # show_df(data_IO)


if __name__ == "__main__":
    main()
