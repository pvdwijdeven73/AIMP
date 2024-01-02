from routines import format_excel, show_df, dprint, quick_excel
import re
import pandas as pd

path = "projects/CEOD/Yoko export/"
file_Yoko_export = "Displays_Yoko_new.txt"
file_display_list = "DisplayList.txt"
outputpath = "projects/CEOD/Output/"
mode = ""

function_types = {
    "Call Data Input Dialog | Data",
    "Call Menu Dialog | Data",
    "Call Panel Set | Panel Set Name",
    "Call Window | Window Name",
    "Instrument Command Operation | Data",
    "Others | Parameter",
}

ack_types = {
    "No Acknowledgment | Command Data",
    "No Acknowledgment | Condition Formula",
    "With Acknowledgment | Command Data",
    "With Acknowledgment | Condition Formula",
    "With Acknowledgment | Condition Formula",
    "With Confirmation | Command Data",
    "With Confirmation | Condition Formula",
}

with open(path + file_Yoko_export, "r", encoding="latin-1") as file:
    lines = file.readlines()

display_list = []
with open(path + file_display_list, "r", encoding="latin-1") as file:
    for display in file:
        display_list.append(display[:-1])
# print(display_list)


def getCMI(line, display, curCMI):
    newCMI = []
    result = []
    if line.count(":") == 3:
        newCMI = [segment.strip() for segment in line.split(":")]
    elif "Execution Mode" in line:
        newCMI = (
            [display] + curCMI + [[segment.strip() for segment in line.split(":")][1]]
        )
    elif "Modifier Action" in line:
        result = [3]
    elif (
        "No Color Change Blink Continue Condition Formula Invert Modify String Character"
        in line
    ):
        result = [2]
    elif "No Color Change Blink Continue Condition Formula" in line:
        result = [1]
    else:
        result = [line]

    # curCMI + modifier 1 (with header)

    return result, newCMI


def check_CMI():
    return


CMI = []
curCMI = []
CMI_mode = 0
line_found = ""
mode = ""
dict_CMI = {}
LGNI = []
dict_LGNI = {}
dict_FLI = {}
dict_DLI = {}
prevline = ""
itemname = ""

display = "Display"
for line in lines:
    if "GE2K HIS0232" in line:
        display_old = display
        display = line[13:-1]
        if display in display_list:
            if display != display_old:
                dict_LGNI[display] = {}
                dict_FLI[display] = []
                dict_DLI[display] = []
            if mode == "Display":
                # dprint(f"Processing display '{display}'...","YELLOW")
                # mode = ""
                continue
            else:
                continue
        else:
            # dprint(f"Skipping display '{display}'...","YELLOW")
            mode = ""
            display = ""
            continue

    if "File Attribute Information" in line:
        # dprint(f"- File Attribute Information found in {display}...","BLUE")
        prevmode = mode
        mode = "Display"
        continue

    if "HIS Control Information" in line:
        # dprint(f"- HIS Control Information found in {display}...","BLUE")
        prevmode = mode
        mode = "HIS"
        continue

    if "Linked Parts Information" in line:
        # dprint(f"- Linked Parts Information found in {display}...","BLUE")
        prevmode = mode
        mode = "LPI"
        continue

    if "Data Link Information" in line and display != "":
        # dprint(f"- Data Link Information found in {display}...","BLUE")
        prevmode = mode
        mode = "DLI"

        continue

    if "Condition Modifier Information" in line and display != "":
        # dprint(f"- Condition Modifier Information found in {display}...","BLUE")
        prevmode = mode
        mode = "CMI"

        continue

    if "Local Generic Name Information" in line and display != "":
        # dprint(f"- Local Generic Name Information in {display}...","BLUE")
        prevmode = mode
        mode = "LGNI"

        continue

    if "Function Link Information" in line and display != "":
        # dprint(f"- Function Link Information in {display}...","BLUE")
        prevmode = mode
        mode = "FLI"

        continue

    if mode == "DLI":
        line = line.strip()
        dict_DLI[display].append(line)

    if mode == "FLI":
        line = line.strip()
        if len(line) > 3:
            if " : " not in line:
                dict_FLI[display][-1] += line
            else:
                dict_FLI[display].append(line)

    if mode == "LGNI":
        line = line.strip()
        totalline = ""
        if len(line) > 3 and "1 : :" not in line:
            if line[0] == "$" and line[-2:] != "No" and line[-3:] != "Yes":
                prevline = line
            else:
                if prevline != "":
                    totalline = prevline + line
                    prevline = ""
                else:
                    if "Generic Name Initial Value Character" not in line:
                        totalline = line
            if totalline != "":
                if totalline[0].isnumeric():
                    items = totalline.split(" : ")
                    if len(items) > 3:
                        cur_itemname = itemname
                        itemname = items[2] + "@" + items[3]
                        if cur_itemname != itemname:
                            dict_LGNI[display][itemname] = {}
                if totalline[0] == "$":
                    items = totalline.split(" ")
                    # print(display + " - " + itemname + " - " + items[0])
                    # print(dict_LGNI[display][itemname])
                    dict_LGNI[display][itemname][items[0]] = items[1]

    if mode == "CMI" and len(line) > 3:
        # print(line[:-1])
        CMI_line, curCMI = getCMI(line, display, curCMI)
        if CMI_line == []:
            if len(curCMI) < 5:
                pass
            else:
                if line_found != "":
                    # dprint(line_found, "GREEN")
                    # print(line_found + "- new CMI")
                    dict_CMI[curCMIabb][line_found[0]] = {}
                    dict_CMI[curCMIabb][line_found[0]]["properties1"] = line_found
                    line_found = ""
                # dprint(curCMI,"CYAN")
                # print(curCMI)
                curCMIabb = curCMI[0] + " * " + curCMI[3] + "@" + curCMI[4]
                dict_CMI[curCMIabb] = {}
        else:
            if CMI_line[0] == 1:
                CMI_mode = 1
                # dprint("Mode1","BLUE")
                if line_found != "":
                    # dprint(line_found,"GREEN")
                    # print(line_found + "- Mode1")
                    dict_CMI[curCMIabb][line_found[0]] = {}
                    dict_CMI[curCMIabb][line_found[0]]["properties1"] = line_found
                    line_found = ""
            elif CMI_line[0] == 2:
                CMI_mode = 2
                # print("MODE2 FOUND")
                # dprint("Mode2","BLUE")
                if line_found != "":
                    # print(line_found + "- Mode2")
                    dict_CMI[curCMIabb][line_found[0]] = {}
                    dict_CMI[curCMIabb][line_found[0]]["properties1"] = line_found
                line_found = ""
            elif CMI_line[0] == 3:
                CMI_mode = 3
                # print("MODE3 FOUND")
                # dprint("Mode3","BLUE")
                cur_mod = line_found[0]
                dict_CMI[curCMIabb][line_found[0]] = {}
                dict_CMI[curCMIabb][line_found[0]]["properties2"] = line_found
                # print(line_found + "- Mode3")
                line_found = ""
            else:
                # dprint(f"current mode:{CMI_mode}","BLUE")
                if CMI_line[0] != "":
                    cur_line = CMI_line[0][:-1]
                    # print(f"'{cur_line}'")
                    if CMI_mode == 1:
                        line_found += cur_line
                        CMI_mode = 11
                    elif CMI_mode == 11:
                        if cur_line[0].isnumeric():
                            # dprint (line_found,"GREEN")
                            # print(line_found + "- Mode11")
                            line_found += cur_line
                            CMI_mode = 1
                        else:
                            line_found += cur_line
                            dict_CMI[curCMIabb][line_found[0]] = {}
                            dict_CMI[curCMIabb][line_found[0]][
                                "properties1"
                            ] = line_found
                    elif CMI_mode == 2:
                        # print("MODE2")
                        line_found += cur_line
                        CMI_mode = 21
                    elif CMI_mode == 21:
                        if cur_line[0].isnumeric():
                            # print("MODE21")
                            # dprint (line_found,"GREEN")
                            # print(line_found + "- Mode21")
                            line_found += cur_line
                            CMI_mode = 1
                        else:
                            line_found += cur_line
                    elif CMI_mode == 3:
                        # print("MODE3")
                        line_found += cur_line
                        CMI_mode = 31
                    elif CMI_mode == 31:
                        if cur_line[0].isnumeric():
                            # print("MODE31")
                            # dprint (line_found,"GREEN")
                            # print(cur_line + "- Mode31")
                            if cur_line[0] == "1":
                                dict_CMI[curCMIabb][cur_mod][
                                    "Modifier1_property"
                                ] = cur_line[2 : cur_line.find(" ", 3)]
                                dict_CMI[curCMIabb][cur_mod][
                                    "Modifier1_value"
                                ] = cur_line[cur_line.find(" ", 3) :]
                                # leave empty if modifier2 is not there
                                dict_CMI[curCMIabb][cur_mod]["Modifier2_property"] = ""
                                dict_CMI[curCMIabb][cur_mod]["Modifier2_value"] = ""
                            else:
                                dict_CMI[curCMIabb][cur_mod][
                                    "Modifier2_property"
                                ] = cur_line[2 : cur_line.find(" ", 3)]
                                dict_CMI[curCMIabb][cur_mod][
                                    "Modifier2_value"
                                ] = cur_line[cur_line.find(" ", 3) :]
                            line_found = ""
                        else:
                            line_found += cur_line
        # CMI.append(CMI_line)

# print(dict_CMI)

dict_final = {}

# property_list = []
# for item1 in dict_CMI:
#     for item2 in dict_CMI[item1]:
#         item_name = item1 + " _ " + item2
#         for item3 in dict_CMI[item1][item2]:
#             property_list.append(item3)
# property_list = list(set(property_list))


for item1 in dict_CMI:
    for item2 in dict_CMI[item1]:
        item_name = item1 + " _ " + item2
        dict_final[item_name] = {}
        for item3 in dict_CMI[item1][item2]:
            dict_final[item_name]["Display"] = item1.split(" * ")[0]
            dict_final[item_name]["Shape"] = item1.split(" * ")[1]
            dict_final[item_name]["Modifier"] = item2
            # if 'transparent' in properties and properties= properties2, treat as properties1
            if item3 == "properties1":
                stuff1 = dict_CMI[item1][item2][item3]
            elif item3 == "properties2":
                if "Transparent" in dict_CMI[item1][item2][item3]:
                    stuff1 = dict_CMI[item1][item2][item3].strip()
                    stuff2 = ""
                else:
                    stuff = dict_CMI[item1][item2][item3]
                    third_last_space_index = len(" ".join(stuff.split()[:-3])) - 1
                    stuff1 = stuff[: third_last_space_index + 1].strip()
                    stuff2 = stuff[third_last_space_index + 1 :].strip()

                # Invert | Modify | String Character
                if stuff2 != "" and len(stuff2.split(" ")) == 3:
                    dict_final[item_name]["Invert"] = stuff2.split(" ")[0]
                    dict_final[item_name]["Modify_String"] = stuff2.split(" ")[1]
                    dict_final[item_name]["Character"] = stuff2.split(" ")[2]
            else:
                dict_final[item_name][item3] = dict_CMI[item1][item2][item3]
                stuff1 = ""
                # No | Color Change | Blink | Continue | Condition Formula
            if stuff1 != "":
                if "True" in stuff1:
                    stuff11 = stuff1.split("True")[0][2:]
                    Cond_form = stuff1.split("True")[1].strip()
                    Continue = "True"
                else:
                    stuff11 = stuff1.split("False")[0][2:]
                    Cond_form = stuff1.split("False")[1].strip()
                    Continue = "False"
                if "Alarm Specific Blinking" in stuff11:
                    blink = "Alarm Specific Blinking"
                elif stuff11.strip()[-1] == "o":
                    blink = "No"
                else:
                    blink = "Yes"
                color = stuff11[: stuff11.find(blink)]
                if color.strip() == "":
                    color = "No Color Change"

                dict_final[item_name]["Color_Change"] = color
                dict_final[item_name]["Blink"] = blink
                dict_final[item_name]["Continue"] = Continue
                dict_final[item_name]["Condition_Formula"] = Cond_form

                Cond_form_tags = Cond_form

                try:
                    for item in dict_LGNI[item1.split(" * ")[0]][item1.split(" * ")[1]]:
                        if item in Cond_form_tags:
                            Cond_form_tags = Cond_form_tags.replace(
                                item,
                                dict_LGNI[item1.split(" * ")[0]][item1.split(" * ")[1]][
                                    item
                                ],
                            )
                            # #print(
                            #     Cond_form_tags,
                            #     item,
                            #     dict_LGNI[item1.split(" * ")[0]][item1.split(" * ")[1]][
                            #         item
                            #     ],
                            # )
                    dict_final[item_name]["Condition_Formula_Tags"] = Cond_form_tags
                except:
                    dict_final[item_name]["Condition_Formula_Tags"] = Cond_form_tags

                # dict_final[item_name]["stuff1"] = stuff1


# quick_excel(df_CMI, path, "CMI_test", True, False)


def get_param(display, shape, line):
    try:
        for item in dict_LGNI[display][shape]:
            if item in line:
                line = line.replace(
                    item,
                    dict_LGNI[display][shape][item],
                )
                # #print(
                #     Cond_form_tags,
                #     item,
                #     dict_LGNI[item1.split(" * ")[0]][item1.split(" * ")[1]][
                #         item
                #     ],
                # )
        return line
    except:
        return line


FLI_dict = {}
item_name = ""
for item1 in dict_FLI:
    for item2 in dict_FLI[item1]:
        if item2[0].isnumeric():
            # new item found
            cur_item_name = item_name
            items = item2.split(" : ")
            item_name = f"{item1}#{items[2]}@{items[3]}"
            if item_name != cur_item_name:
                # print(item_name)
                FLI_dict[item_name] = {}
                FLI_dict[item_name]["Display"] = item1
                FLI_dict[item_name]["ID"] = items[0]
                FLI_dict[item_name]["Touch"] = items[1]
                FLI_dict[item_name]["Shape"] = items[2] + "@" + items[3]

        elif "Select Source" in item2:
            FLI_dict[item_name]["Source"] = item2.split(" : ")[1]
        elif "Function Type" in item2:
            for ft in function_types:
                ft_rep = ft.replace(" | ", " ")
                if ft_rep in item2:
                    cur_item = item2.replace(ft_rep, ft)
            FLI_dict[item_name]["Func_type"] = cur_item.split(" | ")[0].split(" : ")[1]
            param = cur_item.split(" | ")[1].split(" : ")[0]
            value = cur_item.split(" | ")[1].split(" : ")[1].replace("Parameter :", "")
            value = get_param(item_name.split("#")[0], item_name.split("#")[1], value)
            FLI_dict[item_name][param] = value
        elif "Acknowledgment" in item2 or "Acknowledgement" in item2:
            for at in ack_types:
                at_rep = at.replace(" | ", " ")
                if at_rep in item2:
                    cur_item = item2.replace(at_rep, at) + " "
            FLI_dict[item_name]["Acknowledgment"] = cur_item.split(" | ")[0].split(
                " : "
            )[1]
            param = cur_item.split(" | ")[1].split(" : ")[0]
            value = (
                cur_item.split(" | ")[1]
                .split(" : ")[1]
                .replace("Condition Formula", "")
            )

            value = get_param(item_name.split("#")[0], item_name.split("#")[1], value)
            FLI_dict[item_name][param] = value
        else:
            cur_item = item2.split(" : ")
            value = cur_item[1]
            value = get_param(item_name.split("#")[0], item_name.split("#")[1], value)
            FLI_dict[item_name][cur_item[0]] = value


# quick_excel(df_FLI, path, "CFLI_test", True, False)


with open(path + "DLI.txt", "w") as fp:
    for item1 in dict_DLI:
        for item2 in dict_DLI[item1]:
            fp.write(f"{item1},{item2}\n")
    print("Done")


DLI_dict = {}
item_name = ""
for item1 in dict_DLI:
    for item2 in dict_DLI[item1]:
        if item2[0].isnumeric():
            # new item found
            items = item2.split(" : ")
            item_name = f"{item1}#{items[2]}@{items[3]}"
            # print(item_name)
            DLI_dict[item_name] = {}
            DLI_dict[item_name]["Display"] = item1
            DLI_dict[item_name]["ID"] = items[0]
            DLI_dict[item_name]["Touch"] = items[1]
            DLI_dict[item_name]["Shape"] = items[2] + "@" + items[3]

        elif "Property" in item2:
            continue
        else:
            if (
                item2.split(" ")[0] == "Value"
                or "Limit" in item2.split(" ")[0]
                or "Content" in item2.split(" ")[0]
            ):
                param = item2.split(" ")[0]
                value = "".join(item2.split(" ")[1:])
                value = get_param(
                    item_name.split("#")[0], item_name.split("#")[1], value
                )
                DLI_dict[item_name][param] = value
            else:
                items = item2.split(" ")
                DLI_dict[item_name]["Property"] = items[0]
                # print(items)
                DLI_dict[item_name]["LowLimit"] = items[-4]
                DLI_dict[item_name]["HighLimit"] = items[-3]
                DLI_dict[item_name]["From"] = items[-2]
                DLI_dict[item_name]["To"] = items[-1]
                value = " ".join(items[1:-4])
                value = get_param(
                    item_name.split("#")[0], item_name.split("#")[1], value
                )
                DLI_dict[item_name]["Value"] = value

        # elif "Function Type" in item2:
        #     for ft in function_types:
        #         ft_rep = ft.replace(" | ", " ")
        #         if ft_rep in item2:
        #             cur_item = item2.replace(ft_rep, ft)
        #     FLI_dict[item_name]["Func_type"] = cur_item.split(" | ")[0].split(" : ")[1]
        #     param = cur_item.split(" | ")[1].split(" : ")[0]
        #     value = cur_item.split(" | ")[1].split(" : ")[1].replace("Parameter :", "")
        #     value = get_param(item_name.split("#")[0], item_name.split("#")[1], value)
        #     FLI_dict[item_name][param] = value
        # elif "Acknowledgment" in item2 or "Acknowledgement" in item2:
        #     for at in ack_types:
        #         at_rep = at.replace(" | ", " ")
        #         if at_rep in item2:
        #             cur_item = item2.replace(at_rep, at) + " "
        #     FLI_dict[item_name]["Acknowledgment"] = cur_item.split(" | ")[0].split(
        #         " : "
        #     )[1]
        #     param = cur_item.split(" | ")[1].split(" : ")[0]
        #     value = (
        #         cur_item.split(" | ")[1]
        #         .split(" : ")[1]
        #         .replace("Condition Formula", "")
        #     )

        #     value = get_param(item_name.split("#")[0], item_name.split("#")[1], value)
        #     FLI_dict[item_name][param] = value
        # else:
        #     cur_item = item2.split(" : ")
        #     value = cur_item[1]
        #     value = get_param(item_name.split("#")[0], item_name.split("#")[1], value)
        #     FLI_dict[item_name][cur_item[0]] = value


def get_loop(tagname):
    pattern = r"^(\d{1,4})([a-zA-Z]{1,10})(\d{1,4}).*"
    match = re.match(pattern, tagname)
    if match:
        # Extract the groups from the match
        loopname = match.group(1) + match.group(2)[0] + match.group(3)
        return loopname
    return tagname


def add_tag(item, tagname, dict, source):
    itemname = item["Display"] + "#" + item["Shape"] + "|" + tagname
    dict[itemname] = {}
    dict[itemname]["Display"] = item["Display"]
    dict[itemname]["Shape"] = item["Shape"]
    dict[itemname]["Tagname"] = tagname
    dict[itemname]["Loopname"] = get_loop(tagname)
    dict[itemname]["Source"] = source
    return dict


def remove_period(tagname):
    pattern = r"^([a-zA-Z0-9]{1,20})[.].*"
    match = re.match(pattern, tagname)
    if match:
        tagname = match.group(1)
        return tagname
    return ""


def get_tags(input_string):
    # Your list of characters to split on
    split_chars = ["=", ">", "<", ")", "(", "AND", "OR", " "]
    # split_chars = ['-', ':', ';', ',']
    # Your input string

    # Create a regular expression pattern that matches any of the characters
    pattern = "|".join(re.escape(char) for char in split_chars)

    # Split the string using the regular expression pattern as the delimiter
    result = re.split(pattern, input_string)
    tags = []
    for x in result:
        if "." in x:
            tags.append(x)
    # Print the result
    return tags


dict_Overview = {}
for item1 in FLI_dict:
    item = FLI_dict[item1]
    if "Parameter" in item:
        dict_Overview = add_tag(
            item, item["Parameter"], dict_Overview, "FunctionLink_Parameter"
        )
    if "Data" in item:
        tagname = remove_period(item["Data"])
        if tagname != "":
            dict_Overview = add_tag(item, tagname, dict_Overview, "FunctionLink_Data")
    if "Condition Formula" in item:
        if item["Condition Formula"] != "":
            tagname = remove_period(item["Condition Formula"])
            if tagname != "":
                dict_Overview = add_tag(
                    item, tagname, dict_Overview, "FunctionLink_Condition_Formula"
                )
for item1 in DLI_dict:
    item = DLI_dict[item1]
    if "Value" in item:
        tagname = remove_period(item["Value"])
        if tagname != "":
            dict_Overview = add_tag(item, tagname, dict_Overview, "DataLink_Value")
for item1 in dict_final:
    item = dict_final[item1]
    if "Condition_Formula_Tags" in item:
        taglist = get_tags(item["Condition_Formula_Tags"])
        for tag in taglist:
            tagname = remove_period(tag)
            if tagname != "":
                dict_Overview = add_tag(
                    item, tagname, dict_Overview, "ConditMod_Condition_Formula_Tags"
                )


df_overview = pd.DataFrame.from_dict(dict_Overview).transpose()
df_FLI = pd.DataFrame.from_dict(FLI_dict).transpose()
df_CMI = pd.DataFrame.from_dict(dict_final).transpose()
df_DLI = pd.DataFrame.from_dict(DLI_dict).transpose()

with pd.ExcelWriter(path + "displays_yoko_params.xlsx") as writer:
    print("- Writing Function Link Information")
    df_FLI.to_excel(writer, sheet_name="FunctionLink", index=False)
with pd.ExcelWriter(
    path + "displays_yoko_params.xlsx", engine="openpyxl", mode="a"
) as writer:
    print("- Writing Condition Modifier Information")
    df_CMI.to_excel(writer, sheet_name="CondMod", index=False)
    print("- Writing Data Link Information")
    df_DLI.to_excel(writer, sheet_name="DataLink", index=False)
    print("- Writing Overview")
    df_overview.to_excel(writer, sheet_name="Overview", index=False)

print("Formatting excel..")
format_excel(path, "displays_yoko_params.xlsx")
print("Done!")
