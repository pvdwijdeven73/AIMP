tabs = [
    "U3000R",
    "U3000_REAC",
    "TEMP_PROF",
    "HEAT",
    "U3000_UUR",
    "V3001",
    "V2304D",
    "KATALYS",
    "PO_KAT_CL",  #
    "PO_MCH_CL",  #
    "MP1_PO_CL",  #
    "U3000DRAIN",  #
    "DRAINEN",
    "SEQ_01",
    "SEQ_02",
    "SEQ_03",
    "SEQ_04",
    "SEQ_05",
    "SEQ_06_1",
    "SEQ_06_2",
    "SEQ_06_3",
    "SEQ_06_4",
    "SEQ_07",
    "SEQ_08",
    "SEQ_09",
    "SEQ_10",
    "IGNORE",
    "SETTINGS",
    "SWTCH_REQ",
    "STARTUP",
    "U300D1",
    "C3010",
    "C3010_UUR",
    "C3020",
    "C3030",
    "V3030",
    "C3030_UUR",
    "C3040",
    "C3040_UUR",
    "U3000_VAC",
    "VAC_DETAIL",
    "V3006",
    "V3007",
    "RNDN_IONOL",  #
    "ANL_MEPRX",  #
    "GC_STS_TBL",  #
    "ANL_EPROX",
    "U2300",
    "EOPO_OPSL",  #
    "EOPO_UUR",  #
    "A2302A/B",
    "UTILITIES",
    "U2300",
    "U3000",
    "WEC",
    "WRM_WTR",
    "CHILLWATER",
    "T1799",
    "T1799_UUR",
    "TANKNPRK-1",  #
    "TANKNPRK-2",  #
    "TANKNPRK-3",  #
    "TANKNPRK-4",  #
    "SEALS",
    "U3000SNF_I",  #
    "U3000SNF_F",  #
    "U2300SNF",  #
    "LAYOUT",
    "GASDET",  #
    "AUXCEOD",
    "991DCS430",
    "ANL_HOUSE",
    "OS_18G",
    "FAR",
    "U3000_SGS",
    "992DCS001/2",
    "CCC9_KLDR",
    "U3000_FGS",
    "990DCS004",
    "U2300_CUS",
    "U3000_CUS",
    "U2300_FLOWC",
    "U3000_FLOWC",
    "SP_RAMPNG",  #
    "MV_RAMPNG",  #
    "007_DETAIL",
    "005_DETAIL",
    "ALMSUPR1",
    "ALMSUPR2",
    "ALMSUPR3",
    "ALMSUPR4",
    "ALMSUPR5",
    "NMODE_1",
    "NMODE_2",
    "SCHAKLP-L",
    "SCHAKLP-E",
    "SCHAKLP-P",
    "SCHAKLP-F",
    "SCHAKLP-T",
    "UZ_OVZ",
    "TK1247-51",
    "TK1799",
    "121T-M126",
    "18UZ200",
    "23UZ100",
    "23UZ150",
    "23UZ300",
    "23UZ500",
    "U23BB",
    "VOORWRDN",  #
    "30UZ740",
    "30UZ500",
    "30UZ530",
    "30UZ720A",
    "30UZ720B",
    "30UZ710",
    "30UZ730",
    "30XZ003",
    "30XZ006",
    "30UZ700",
    "30TDZ010",
    "30TDZ010H",
    "30TYZ005",
    "30TYZ006",
    "30TDZ011",
    "30TYZ008",
    "30UZ810",
    "30UZ741A",
    "30UZ790A",
    "30UZ790B",
    "30UZ770A",
    "30UZ770B",
    "30UZ760",
    "30UZ780A",
    "30UZ780B",
    "30XZ752",
    "30XZ751",
    "30UZ800",
    "30XZ007",
    "30UZ741C",
    "30UZ100",
    "30UZ200",
    "30UZ300",
    "30UZ400",
    "30UZ860",
    "30UZ600",
    "MOS_U3000",
    "MOS_U2300",
    "DOS_OVZ",
    "FF_OVZ",
    "TANKNPRK",
    "BUILDINGS",
    "CUSUM",
    "FLOWCMP",
    "RAMPING",
    "ALMSUPPR",
    "NMODE",
    "SCHAKELP",
    "UZ_OVZ",
    "MOS/DOS",
]


def replace_and_create_files(strings, path, template_file):
    with open(path + template_file, "r") as f:
        template_content = f.read()

    batch_size = 16
    num_batches = (len(strings) + batch_size - 1) // batch_size

    for batch_num in range(num_batches):
        start_idx = batch_num * batch_size
        end_idx = min(start_idx + batch_size, len(strings))
        batch_strings = strings[start_idx:end_idx]

        modified_content = template_content
        for i, s in enumerate(batch_strings, start=1):
            placeholder = f"##TAB{i}##"
            modified_content = modified_content.replace(placeholder, s)

        output_filename = f"Tab_test{batch_num + 1:02d}.htm"
        with open(path + output_filename, "w") as f:
            f.write(modified_content)


if __name__ == "__main__":
    # Read your list of strings from a file or define it here
    strings = tabs

    template_file = "Tab_test.htm"
    path = "Projects\\CEOD\\Displays\\"
    replace_and_create_files(strings, path, template_file)

# import os


# def process_subfolder(subfolder_path):
#     file_list = os.listdir(subfolder_path)

#     if len(file_list) == 1 and file_list[0].upper() == "DUMMY.TMP":
#         return None  # Ignore subfolder with only DUMMY.TMP
#     else:
#         return os.path.basename(subfolder_path)  # Return subfolder name


# main_folders = [
#     "Projects\CEOD\Yoko export\GE2K\FCS0101\FUNCTION_BLOCK",
#     "Projects\CEOD\Yoko export\GE2K\FCS0102\FUNCTION_BLOCK",
#     "Projects\CEOD\Yoko export\GE2K\FCS0103\FUNCTION_BLOCK",
# ]

# valid_subfolders = []

# for main_folder in main_folders:
#     print(main_folder[31:38])
#     full_path = os.path.join(os.getcwd(), main_folder)
#     for root, dirs, files in os.walk(full_path):
#         for dir_name in dirs:
#             subfolder_path = os.path.join(root, dir_name)
#             subfolder_name = process_subfolder(subfolder_path)
#             if subfolder_name:
#                 valid_subfolders.append([main_folder[31:38], subfolder_name])

# print("Valid subfolders:")
# for subfolder in valid_subfolders:
#     print(subfolder)


# import os
# import xlwings as xw
# from openpyxl import Workbook, load_workbook


# def copy_csv_to_excel(csv_folder, output_file):
#     # Get a list of all CSV files in the given folder
#     csv_files = [f for f in os.listdir(csv_folder) if f.endswith(".csv")]

#     if not csv_files:
#         print("No CSV files found in the folder.")
#         return

#     # Create a new Excel workbook
#     wb = Workbook()
#     sheet_names = []

#     for csv_file in csv_files:
#         # Extract the name of the CSV file (without extension) to use as the tab name
#         tab_name = csv_file[:-4]  # Remove '.csv' extension
#         sheet_names.append(tab_name)

#         # Open the CSV file and read its content
#         csv_file_path = os.path.join(csv_folder, csv_file)
#         with open(csv_file_path, "r") as csv_file:
#             lines = csv_file.readlines()

#             # Create a new sheet in the workbook
#             sheet = wb.create_sheet(title=tab_name)

#             # Write the CSV data to the sheet
#             for row_idx, line in enumerate(lines, start=1):
#                 values = line.strip().split(",")
#                 for col_idx, value in enumerate(values, start=1):
#                     sheet.cell(row=row_idx, column=col_idx, value=value)

#     # Remove the default "Sheet" that is created by openpyxl
#     if "Sheet" in wb.sheetnames:
#         wb.remove(wb["Sheet"])

#     # Save the Excel file
#     wb.save(output_file)

#     # Close the workbook
#     wb.close()


# def open_excel_with_xlwings(file_path):
#     # Open the Excel file using xlwings to display it
#     app = xw.App(visible=True)
#     wb = xw.Book(file_path)
#     wb.api.Activate()  # Activate the workbook to bring it to the front
#     app.run()  # Run the xlwings application


# if __name__ == "__main__":
#     # Provide the folder containing the CSV files and the desired output Excel file name
#     csv_folder_path = "Projects\CEOD\IOCards"
#     output_excel_file = "output_file.xlsx"

#     copy_csv_to_excel(csv_folder_path, output_excel_file)
#     open_excel_with_xlwings(output_excel_file)


# # constants
# path = "projects/CEOD/"
# outputpath = "projects/CEOD/Output/"
# file_params = "C300_Params.txt"
# file_result = "C300_params_result.txt"
# file_output = "C300_params_parsed.txt"
# file_params_index = "C300_Params_index.txt"


# import re
# import pandas as pd
# from routines import show_df, quick_excel


# def process_text(file_path, output_file):
#     # Read the entire text file as a single string
#     with open(file_path, "r", encoding="utf-8") as file:
#         text = file.read()

#     # Find all occurrences of "This is a text" and replace with "***This is a text***"
#     pattern = r"Specific(?:\s*[\r\n]\s*|\s+)to(?:\s*[\r\n]\s*|\s+)Block"
#     processed_text = re.sub(pattern, r"***Specific to Block:::", text)
#     pattern = r"Description"
#     processed_text = re.sub(pattern, r"***Description:::", processed_text)
#     pattern = r"Data(?:\s*[\r\n]\s*|\s+)Type"
#     processed_text = re.sub(pattern, r"***Data Type:::", processed_text)
#     pattern = r"Range"
#     processed_text = re.sub(pattern, r"***Range:::", processed_text)
#     pattern = r"Default"
#     processed_text = re.sub(pattern, r"***Default:::", processed_text)
#     pattern = r"Config(?:\s*[\r\n]\s*|\s+)Load"
#     processed_text = re.sub(pattern, r"***Config Load:::", processed_text)
#     pattern = r"Active(?:\s*[\r\n]\s*|\s+)Loadable"
#     processed_text = re.sub(pattern, r"***Active Loadable:::", processed_text)
#     pattern = r"Access(?:\s*[\r\n]\s*|\s+)Lock"
#     processed_text = re.sub(pattern, r"***Access Lock:::", processed_text)
#     pattern = r"Residence"
#     processed_text = re.sub(pattern, r"***Residence:::", processed_text)
#     pattern = r"Related(?:\s*[\r\n]\s*|\s+)Parameters"
#     processed_text = re.sub(pattern, r"***Related Parameters:::", processed_text)
#     pattern = r"Remarks"
#     processed_text = re.sub(pattern, r"***Remarks:::", processed_text)

#     pattern = r":::(?![\r\n])"
#     processed_text = re.sub(pattern, "***\n", processed_text)
#     pattern = r":::"
#     processed_text = re.sub(pattern, "***", processed_text)
#     # Remove "(s) " at the start of a line
#     processed_text = re.sub(r"^\(s\) ", "", processed_text, flags=re.MULTILINE)
#     # Remove lines containing only "(s)"
#     processed_text = re.sub(r"^\(s\)$", "", processed_text, flags=re.MULTILINE)
#     pattern = r"^(\d{1,2}\.\d{1,3}\s+.+)$"
#     processed_text = re.sub(pattern, r"===\1", processed_text, flags=re.MULTILINE)

#     # Write the processed text back to the file
#     dict_param = {}
#     tagname = "DUMP"
#     ID = 0
#     param = ""
#     param_value = ""
#     for line in processed_text.split("\n"):
#         try:
#             if line[0:3] == "===":
#                 tagname = line.split()[1]
#                 ID += 1
#                 dict_param[ID] = {}
#                 dict_param[ID]["Parameter"] = tagname
#             elif line[0:3] == "***":
#                 # param found
#                 if param != "" and tagname != "DUMP":
#                     dict_param[ID][param] = param_value
#                     param_value = ""
#                 param = line.replace("***", "")
#                 param_value = ""
#             else:
#                 if param != "" and tagname != "DUMP":
#                     param_value += line
#         except:
#             if param != "" and tagname != "DUMP":
#                 param_value += line
#     df = pd.DataFrame(dict_param)
#     quick_excel(df.transpose(), path, "Test", False, False)
#     with open(output_file, "w", encoding="utf-8") as file:
#         file.write(processed_text)


# # Example usage:
# input_file_path = path + file_result
# process_text(input_file_path, path + file_output)


# #     # if line.strip() in [
# #     #     "Specific to Block(s)",
# #     #     "Description",
# #     #     "Data Type",
# #     #     "Range",
# #     #     "Default",
# #     #     "Config Load",
# #     #     "Active Loadable",
# #     #     "Access Lock",
# #     #     "Residence",
# #     #     "Related Parameters",
# #     #     "Remarks",
# #     # ]:
