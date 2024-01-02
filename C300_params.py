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
file_params = "C300_Params.txt"
file_result = "C300_params_result.txt"
file_params_index = "C300_Params_index.txt"


def ignore_patterns(line):
    # Define regular expression patterns for each case
    chapter_pattern = r"^Chapter .*$"
    dash_pattern = r"^- \d{2,4} -\s*$"
    xx_yyy_zaaa_pattern = r"\b\d{1,2}\.\d{1,3} .+ \d{2,4}\b"
    yXXXX_parameters_pattern = r"^[A-Z]{1}[A-Z0-9]{3} PARAMETERS$"
    chapter_z_pattern = r"^[A-Z]+ \d{1,2}$"

    # Check if the line matches any of the patterns
    if (
        not re.match(chapter_pattern, line)
        and not re.match(dash_pattern, line)
        and not re.search(xx_yyy_zaaa_pattern, line)
        and not re.match(yXXXX_parameters_pattern, line)
        and not re.match(chapter_z_pattern, line)
    ):
        return line


encodings_to_try = ["utf-8", "latin-1", "utf-16"]

for encoding in encodings_to_try:
    try:
        print(encoding)
        with open(path + file_params, "r", encoding=encoding) as file:
            lines = file.readlines()
        break
    except UnicodeDecodeError:
        continue
else:
    raise UnicodeDecodeError(
        "Failed to decode the file using any of the specified encodings."
    )


# Process each line and ignore lines that match the specified patterns
result_lines = [line for line in lines if ignore_patterns(line)]

with open(path + file_result, "w", encoding="utf-8") as file:
    file.writelines(result_lines)

# result_lines = []
# merge_next_line = False
# for line in lines:
#     if ignore_patterns(line):
#         result_lines.append(line.strip())


#     # if line.strip() in [
#     #     "Specific to Block(s)",
#     #     "Description",
#     #     "Data Type",
#     #     "Range",
#     #     "Default",
#     #     "Config Load",
#     #     "Active Loadable",
#     #     "Access Lock",
#     #     "Residence",
#     #     "Related Parameters",
#     #     "Remarks",
#     # ]:

# # Write the output file with UTF-8 encoding
# with open(path + file_result, "w", encoding="utf-8") as file:
#     file.writelines(result_lines)
