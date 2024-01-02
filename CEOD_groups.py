# import libraries
import pandas as pd
import numpy as np
from routines import read_db, dprint, quick_excel
import os
from pandasgui import show as show_df
import re

# constants
path = "projects/CEOD/"
file_Yoko_export = "Yoko_Centum_CS3000_Parser_Output_3_3_2023_11_30_ 9_AM.xlsx"
outputpath = "projects/CEOD/Output/"


def import_Yoko_data():
    dprint("- Loading Yokogawa export...", "BLUE")
    data_Yoko = read_db(path, file_Yoko_export, sheet=None)

    return data_Yoko


def isolate_tagname(Tagname):
    colon_indices = [i for i, char in enumerate(Tagname) if char == ":"]

    if len(colon_indices) == 2:
        start_index = colon_indices[0] + 1
        end_index = colon_indices[1]
        Tagname = Tagname[start_index:end_index].strip()
    pattern = r"^(.*?)-(.{5,})$"
    match = re.match(pattern, Tagname)
    if match:
        Tagname = match.group(2)

    return Tagname.split(".")[0]


def get_connections(data_Yoko):
    dprint("Creating connections", "YELLOW")
    connections = {}
    last_group = 0
    tag_list = []
    for sheet_name, sheet_data in data_Yoko.items():
        columns = sheet_data.columns.to_list()
        connection_columns = [col for col in columns if col.startswith("Connection ")]
        if connection_columns != []:
            # create tag list
            for index, row in sheet_data.iterrows():
                tag_list = []
                tag_list.append(row["PointName"])
                for col in connection_columns:
                    if row[col] != "":
                        tag_list.append(isolate_tagname(row[col]))
                # connections[row["PointName"]] = tag_list

                # scan all groups
                found = []
                tag_list = list(set(tag_list))
                for group_name, group in connections.items():
                    for tag in tag_list:
                        if tag in group and group_name not in found:
                            connections[group_name].extend(tag_list)
                            connections[group_name] = list(set(connections[group_name]))
                            found.append(group_name)
                if len(found) > 1:
                    # combine groups
                    for group in found:
                        if group != found[0]:
                            connections[found[0]].extend(connections[group])
                            del connections[group]
                    connections[found[0]] = list(set(connections[found[0]]))

                if found == []:
                    # create new group
                    last_group += 1
                    connections["group_{:04d}".format(last_group)] = tag_list

    max_length = max(len(lst) for lst in connections.values())
    min_length = min(len(lst) for lst in connections.values())

    # Generate the new keys with leading zeros
    new_keys = ["GROUP_{:04d}".format(i + 1) for i in range(len(connections))]

    # Create a new dictionary with the renamed keys
    connections = {
        new_keys[i]: connections[key]
        for i, key in enumerate(
            sorted(connections.keys(), key=lambda k: len(connections[k]), reverse=True)
        )
    }

    for key in connections:
        if len(connections[key]) < max_length:
            connections[key].extend([None] * (max_length - len(connections[key])))

    return pd.DataFrame.from_dict(connections)


def main():
    # Clear the screen
    os.system("cls" if os.name == "nt" else "clear")
    dprint("Importing data", "YELLOW")
    data_Yoko = import_Yoko_data()
    connections = get_connections(data_Yoko)
    quick_excel(connections, outputpath, "Yoko_connections", True)
    # show_df(connections)


if __name__ == "__main__":
    main()
