# import libraries
import pandas as pd
import numpy as np
from routines import read_db, dprint, quick_excel
import os
from pandasgui import show as show_df


# constants
path = "projects/CEOD/"
outputpath = "projects/CEOD/Output/"
file_Yoko_export = "Yoko_Centum_CS3000_Parser_Output_3_3_2023_11_30_ 9_AM.xlsx"
file_Typical_Match = "Typical_match.xlsx"


def import_data():
    dprint("Importing data", "YELLOW")
    dprint("- Loading Yokogawa export...", "BLUE")
    data_Yoko = pd.read_excel(path + file_Yoko_export, sheet_name=None, na_filter=False)
    dprint("- Loading typical match...", "BLUE")
    typical_match = pd.read_excel(path + file_Typical_Match, na_filter=False)
    typical_match.columns = [
        "ID",
        "Yoko_typical",
        "Description",
        "Tool",
        "Experience",
        "Hon_typical",
        "RefAvailable",
    ]
    return data_Yoko, typical_match[["Yoko_typical", "Hon_typical"]]


def show_data(tag_typical_list, typical_match):
    print(tag_typical_list)

    print(typical_match)


def create_tag_typical_list(data_Yoko, typical_match):
    base_columns = ["Controller", "PointName", "DrawingName", "Tag Comment"]
    all_columns = [
        "Controller",
        "PointName",
        "Tag Comment",
        "DrawingName",
        "Yoko_typical",
    ]
    dprint("Creating typical list", "YELLOW")
    dprint("- Creating list of Yoko typicals...", "BLUE")
    tag_typical_list = pd.DataFrame()
    # Access each sheet's data
    for sheet_name, sheet_data in data_Yoko.items():
        if sheet_name in typical_match["Yoko_typical"].tolist():
            df_temp = sheet_data.loc[:, base_columns].copy()
            df_temp.loc[:, "Yoko_typical"] = sheet_name
            tag_typical_list = pd.concat([tag_typical_list, df_temp])

    # add IO
    df_temp = (
        data_Yoko["IO Channel"]
        .loc[:, ["Controller", "Label / Tag Name", "P&ID", "IO Type"]]
        .copy()
    )
    df_temp.loc[:, "PointName"] = df_temp["Label / Tag Name"].str.split("-").str[1]
    df_temp.loc[:, "Tag Comment"] = df_temp["Label / Tag Name"]
    df_temp.loc[:, "DrawingName"] = df_temp["P&ID"]
    df_temp.loc[:, "Yoko_typical"] = "IO_" + df_temp["IO Type"]
    df_temp = df_temp.dropna(subset=["PointName"])
    tag_typical_list = pd.concat([tag_typical_list, df_temp[all_columns]])

    # add COM analog
    df_temp = (
        data_Yoko["Communication IOs"].loc[:, ["Controller", "Name", "Comment"]].copy()
    )
    df_temp.loc[:, "PointName"] = df_temp["Name"].str.split("-").str[1]
    df_temp.loc[:, "Tag Comment"] = df_temp["Comment"]
    df_temp.loc[:, "DrawingName"] = ""
    df_temp.loc[:, "Yoko_typical"] = "COM_SAI"
    df_temp = df_temp.dropna(subset=["PointName"])
    df_temp["PointName"] = df_temp["PointName"].str.replace("%%I_30XA", "30XA")
    df_temp = df_temp[~df_temp["PointName"].str.contains("SPARE")]
    tag_typical_list = pd.concat([tag_typical_list, df_temp[all_columns]])

    # COM digital
    df_temp = (
        data_Yoko["Communication Tags"]
        .loc[:, ["Controller", "Tag Name", "Comment"]]
        .copy()
    )
    df_temp["Tag Name"] = df_temp["Tag Name"].apply(
        lambda x: x[:-2].replace("_", "-") + x[-2:]
    )

    df_temp.loc[:, "PointName"] = df_temp["Tag Name"].str.split("-").str[1]
    df_temp["Tag Name"] = df_temp["Tag Name"].str.replace("%%", "")
    df_temp.loc[:, "Tag Comment"] = df_temp["Comment"]
    df_temp.loc[:, "DrawingName"] = ""

    df_temp.loc[:, "Yoko_typical"] = "COM_" + df_temp["Tag Name"].str.split("-").str[0]
    df_temp = df_temp.dropna(subset=["PointName"])
    df_temp = df_temp[~df_temp["PointName"].str.contains("SPARE")]
    tag_typical_list = pd.concat([tag_typical_list, df_temp[all_columns]])

    # Switches
    df_temp = (
        data_Yoko["Switches"].loc[:, ["Controller", "Tag Name", "Tag Comment"]].copy()
    )

    df_temp.loc[:, "PointName"] = df_temp["Tag Name"]
    df_temp.loc[:, "DrawingName"] = ""

    df_temp.loc[:, "Yoko_typical"] = "SWITCH"
    df_temp = df_temp.dropna(subset=["PointName"])
    df_temp = df_temp[~df_temp["Tag Comment"].str.contains("System Reserved")]
    tag_typical_list = pd.concat([tag_typical_list, df_temp[all_columns]])

    # Annunciators
    df_temp = (
        data_Yoko["Switches"].loc[:, ["Controller", "Tag Name", "Tag Comment"]].copy()
    )

    df_temp.loc[:, "PointName"] = df_temp["Tag Name"]
    df_temp.loc[:, "DrawingName"] = ""

    df_temp.loc[:, "Yoko_typical"] = "ANN"
    df_temp = df_temp.dropna(subset=["PointName"])
    df_temp = df_temp[~df_temp["PointName"].str.contains("%AN")]
    tag_typical_list = pd.concat([tag_typical_list, df_temp[all_columns]])

    # Global Switches
    df_temp = (
        data_Yoko["Global Switches"]
        .loc[:, ["Controller", "Tag Name", "Tag Comment"]]
        .copy()
    )

    df_temp.loc[:, "PointName"] = df_temp["Tag Name"]
    df_temp.loc[:, "DrawingName"] = ""

    df_temp.loc[:, "Yoko_typical"] = "GL_SWITCH"
    df_temp = df_temp.dropna(subset=["PointName"])
    tag_typical_list = pd.concat([tag_typical_list, df_temp[all_columns]])

    tag_typical_list = match_hon_typical(tag_typical_list, typical_match)
    return tag_typical_list


def match_hon_typical(tag_typical_list, typical_match):
    dprint("- Matching Yoko & Hon typicals...", "BLUE")
    tag_typical_list = pd.merge(
        tag_typical_list, typical_match, on="Yoko_typical", how="outer"
    )
    return tag_typical_list


def param_count(data_Yoko, typical_match):
    dprint("- Creating param overview Yoko...", "BLUE")
    lst_counts = []
    for sheet_name, sheet_data in data_Yoko.items():
        if sheet_name in typical_match["Yoko_typical"].tolist():
            cols = sheet_data.columns
            for col in cols:
                if sheet_data[col].nunique() <= 5:
                    lst_temp = sheet_data[col].astype(str).drop_duplicates().tolist()
                    lst_temp = [
                        "[EMPTY]" if element == "" else element for element in lst_temp
                    ]
                    lst_temp.sort()
                    unique_values = ", ".join(lst_temp)
                else:
                    unique_values = "More than 5 unique values"
                lst_counts.append(
                    [sheet_name, col, sheet_data[col].nunique(), unique_values]
                )

    df_counts = pd.DataFrame(
        lst_counts, columns=["Typical", "Parameter", "#Unique_Values", "Unique_Values"]
    )
    # TODO: general params (occurring in every typical)
    value_counts = df_counts["Parameter"].value_counts()

    # Map the counts to a new column
    df_counts["#Occurances_in_typicals"] = df_counts["Parameter"].map(value_counts)

    mapping = df_counts.groupby("Parameter")["Typical"].apply(list).to_dict()
    df_counts["found_in"] = df_counts["Parameter"].map(mapping)
    df_counts["not_found_in"] = df_counts.apply(
        lambda row: list(set(df_counts["Typical"]) - set(row["found_in"])), axis=1
    )
    del df_counts["found_in"]
    df_counts["not_found_in"] = df_counts["not_found_in"].apply(lambda x: ", ".join(x))
    df_counts.loc[df_counts["#Occurances_in_typicals"] < 50, "not_found_in"] = (
        "All typicals except " + df_counts.loc[:, "not_found_in"]
    )
    df_counts.loc[df_counts["#Occurances_in_typicals"] < 40, "not_found_in"] = ""
    df_counts.loc[
        df_counts["#Occurances_in_typicals"] == 50, "not_found_in"
    ] = "All typicals"
    df_counts = df_counts.rename(columns={"not_found_in": "Found_in"})

    return df_counts


def main():
    # Clear the screen
    os.system("cls" if os.name == "nt" else "clear")

    data_Yoko, typical_match = import_data()
    tag_typical_list = create_tag_typical_list(data_Yoko, typical_match)
    df_counts = param_count(data_Yoko, typical_match)
    # show_df(tag_typical_list)
    # show_data(tag_typical_list, typical_match)
    quick_excel(tag_typical_list, outputpath, "Typicals", True, revision=True)
    # quick_excel(df_counts, outputpath, "Yoko_params", True)
    # show_df(df_counts)


if __name__ == "__main__":
    main()
