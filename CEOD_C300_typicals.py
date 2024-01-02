# import libraries
import pandas as pd
import numpy as np
from routines import read_db, dprint, quick_excel
import os


# constants
path = "projects/CEOD/"
Hon_path = "projects/CEOD/Hon/"
outputpath = "projects/CEOD/Output/"

dict_Typical = {
    "AUTOMAN": {"file": "AUTOMAN_VALIDATION_1.xlsx", "sheet": "AUTOMAN_C300"},
    "DACA": {"file": "DAC_VALIDATION_3.xlsx", "sheet": "DACA_C300"},
    "DIGACQA": {"file": "DIGACQA_VALIDATION_4.xlsx", "sheet": "C300_DIGACQA"},
    "PID": {"file": "PID_VALIDATION_1.xlsx", "sheet": "PID_C300"},
    "DEVCTRL": {"file": "DEVCTRL_VALIDATION_1.xlsx", "sheet": "DEVCTL_C300"},
}


def import_data():
    dprint("Importing data", "YELLOW")
    dprint("- Loading Honeywell export...", "BLUE")
    data_HON = {}
    for typical in dict_Typical:
        data_HON[typical] = pd.read_excel(
            Hon_path + dict_Typical[typical]["file"],
            sheet_name=dict_Typical[typical]["sheet"],
            na_filter=False,
        )
    return data_HON


def param_count(data_Hon):
    dprint("- Creating param overview Honeywell...", "BLUE")
    lst_counts = []
    for sheet_name, sheet_data in data_Hon.items():
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
    df_counts.loc[df_counts["#Occurances_in_typicals"] < 00, "not_found_in"] = ""
    df_counts.loc[
        df_counts["#Occurances_in_typicals"] == len(data_Hon), "not_found_in"
    ] = "All typicals"
    df_counts = df_counts.rename(columns={"not_found_in": "Found_in"})

    return df_counts


def update_rev(df_counts, prefix):
    extension = ".xlsx"

    # Find all files starting with prefix and ending with "_0xxx"
    files = [
        filename
        for filename in os.listdir(outputpath)
        if filename.startswith(prefix) and filename.endswith(extension)
    ]
    new_filename = prefix + extension
    if files:
        max_number = -1
        max_file = ""

        for filename in files:
            print(filename)
            # Extract the number from the filename
            number_str = filename[
                len(prefix) + 1 : -len(extension)
            ]  # Assuming the number has a fixed length of 4 digits
            try:
                number = int(number_str)
                if number > max_number:
                    max_number = number
                    max_file = filename
                    print(f"maxfile: {max_file}")
            except ValueError:
                continue  # Skip files with invalid number format

        if max_file:
            new_filename = prefix + "_{:04d}".format(max_number + 1) + extension
            df_previous = pd.read_excel(outputpath + max_file)
            columns = df_counts.columns.tolist()
            df_counts = df_counts.merge(
                df_previous,
                how="outer",
                on=["Typical", "Parameter"],
                suffixes=("", "_old"),
            )
            columns.remove("Typical")
            columns.remove("Parameter")
            columns_old = [column + "_old" for column in columns]
            columns_new = [column + "" for column in columns]
            df_temp = df_counts[columns_old]
            df_temp.columns = columns_new
            df_counts = df_counts.combine_first(df_temp)
            df_counts = df_counts.drop(columns_old, axis=1)
            df_counts = df_counts[df_previous.columns.tolist()]
    return df_counts, new_filename


def main():
    # Clear the screen
    os.system("cls" if os.name == "nt" else "clear")

    data_Hon = import_data()
    df_counts = param_count(data_Hon)
    df_counts, new_filename = update_rev(df_counts, "Hon_params")
    print(new_filename)
    # TODO create file version update

    quick_excel(df_counts, outputpath, new_filename, True)


if __name__ == "__main__":
    main()
