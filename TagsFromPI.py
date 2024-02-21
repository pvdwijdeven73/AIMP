from glob import glob

folder_displays = "\\Test\\PI\\"

filenames = glob(folder_displays + "\\*.pdi")
for filename in filenames:
    print(filename)
