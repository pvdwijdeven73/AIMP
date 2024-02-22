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
print(len(split_line))
for s_line in split_line:
    res_end = s_line.find(chr(32))
    print(s_line[:res_end])


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
