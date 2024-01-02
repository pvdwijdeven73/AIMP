# import libraries
import pandas as pd
import numpy as np
import os
from simpledbf import Dbf5
import win32com.client as win32
from routines import format_excel, check_folder, read_db


class SOETopX:
    def __init__(self, project: str, phase: str = "Original", isPLC: bool = True):
        self.start(project)

    def get_files(self, input_path, output_file):
        all_files = []

        if all_files == []:
            files = os.listdir(path=input_path)
            for file in files:
                if file != output_file and "." in file:
                    all_files.append([file, file.split(".", 1)[0]])

        print(f"files found: \n{all_files}")

        return all_files

    def read_and_merge(self, all_files, input_path, output_path, output_file):
        lst_db = []

        db = pd.DataFrame()

        print("- Reading files")
        for curFile in all_files:
            print(f"  - Reading {curFile[0]}")
            curDB = read_db(input_path, curFile[0])

            curDB.insert(0, "File", curFile[1])
            lst_db.append(curDB)

        if len(lst_db) > 1:
            db = pd.concat(lst_db)
        elif len(lst_db) == 1:
            db = lst_db[0]

        print("- Done reading files")
        print("- Writing files")

        check_folder(output_path)

        db_counts = (
            db[["Tag Number", "Status Message", "Device #"]].value_counts().to_frame()
        )

        db_counts.columns = ["Count"]

        with pd.ExcelWriter(f"{output_path}{output_file}") as writer:
            print("  - Writing joined DB")
            db_counts.to_excel(writer, sheet_name="SOE analysis", index=True)
        print("- Done writing")
        print("- start formatting excel file")

        format_excel(output_path, output_file)

        print("Ready!!!")

    def start(self, project):
        if project == "Test":
            input_path = "Test\\"
            output_path = "Test\\"
            output_file = "Test_SOE_total.xlsx"
        else:
            input_path = f"Projects\\{project}\\SOE\\"
            output_path = f"Projects\\{project}\\SOE\\"
            output_file = f"{project}_total_SOE.xlsx"

        self.read_and_merge(
            self.get_files(input_path, output_file),
            input_path,
            output_path,
            output_file,
        )


def main():
    project = SOETopX("PSU30_35_45")


if __name__ == "__main__":
    main()
