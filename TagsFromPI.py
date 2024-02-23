from glob import glob

folder_displays = "Test\\PI\\"

import os

files = os.listdir(path=folder_displays)
# for file in files:
file = files[1]
f = open(folder_displays + file, "r", errors="replace")
lines = f.readlines()
totalLine = ""
for line in lines:
    totalLine += line
totnew = ""
for x in totalLine:
    if ord(x) < 32 or ord(x) > 127:
        totnew += chr(32)
    else:
        totnew += x


split_line = totnew.split("COD:")
first = True
arr = []
for s_line in split_line:
    if not first:
        res_end0 = s_line.find(" ")
        if res_end0 == -1:
            res_end0 = 99999999
        res_end1 = s_line.find(".")
        if res_end1 == -1:
            res_end1 = 99999999
        res_end = min(res_end0, res_end1)
        arr.append(s_line[:res_end])
    else:
        first = False
arr = list(set(arr))
arr.sort()
print(arr)

# encodings_to_try = ["utf-8", "latin-1", "utf-16"]

# for encoding in encodings_to_try:
#     try:
#         print(encoding)
#         with open(file=folder_displays + file, mode="r", encoding=encoding) as file:
#             lines = file.readlines()
#         break
#     except UnicodeDecodeError:
#         continue
# else:
#     raise UnicodeDecodeError(
#         "Failed to decode the file using any of the specified encodings."
#     )
# print(lines)
