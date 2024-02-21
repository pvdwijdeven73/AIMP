from glob import glob

folder_displays = "\\Test\\PI\\"

filenames = glob(folder_displays + "\\*.PDI")
for filename in filenames:
    print(filename)
