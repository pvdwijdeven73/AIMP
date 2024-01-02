# This program converts an EB file to a readable flat table.

# Import libraries
import os
import pandas as pd
import numpy as np
from routines import format_excel, check_folder
from os import system
from colorama import Fore


class DBtoEB:
    def __init__(self, project: str, phase: str = "Original"):
        self.start(project, phase)

    def write_EB(self, EB_path, Export_output_path):

        files = os.listdir(path=Export_output_path)
        total = ""
        for file in files:
            if ".XLS" in file.upper():
                print(f"reading {file}...")
                Export_output_file = file

                all_sheets = pd.read_excel(
                    f"{Export_output_path}{Export_output_file}", None
                )
                for sheet in all_sheets:
                    if sheet != "TOTAL":
                        result = self.create_EB(all_sheets[sheet].fillna(""))
                        with open(
                            f"{EB_path}{Export_output_file[:Export_output_file.find('.')]}_{sheet}.EB",
                            "w",
                        ) as f:
                            f.write(result)
                        total += result
        with open(f"{EB_path}TOTAL.EB", "w") as f:
            f.write(total)
        print("done!")

    def create_EB(self, db_EB: pd.DataFrame) -> str:
        str_list = [
            "$CDETAIL",
            "ASSOCDSP",
            "EUDESC",
            "KEYWORD",
            "PTDESC",
            "STATE1",
            "STATE2",
            "STATETXT(0)",
            "STATETXT(1)",
            "STATETXT(2)",
        ]
        result = ""
        max_rows = db_EB.shape[0]
        for row in range(max_rows):
            row_contents = db_EB.iloc[row]
            for param in row_contents.keys():
                if param != "Source":
                    if param == "&T":
                        temp = f"{{SYSTEM ENTITY {db_EB.at[row,'&N']}( )"
                        temp += " " * (79 - len(temp)) + "}\n"
                        result += temp
                        result += "&T " + db_EB.at[row, "&T"] + "\n"
                    elif param == "&N":
                        result += "&N " + db_EB.at[row, "&N"] + "\n"
                    else:
                        if param in str_list:
                            result += (
                                param + ' = "' + str(db_EB.at[row, param]) + '" \n'
                            )
                        else:
                            result += param + " = " + str(db_EB.at[row, param]) + " \n"

        return result

    def get_params_list(self, EB_path):
        files = os.listdir(path=EB_path)

        params = []
        result = []
        for file in files:
            if ".EB" in file.upper():
                with open(EB_path + file) as f:
                    result += f.readlines()

        for line in result:
            #            print(line)
            if '"' in line:
                # print(line.find(" "))
                # print(line[: 6])
                params.append(line[: line.find(" ")])
        params = list(set(params))
        params.sort()
        print(params)
        return ""

    def start(
        self,
        project,
        phase,
    ):

        print(
            f"{Fore.MAGENTA}Creating EB files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )

        if project == "Test":
            EB_path = "Test\\"
            Export_output_path = "Test\\"
            Export_output_file = "PSU30_35_45_export_EB_total_Original.xlsx"
        else:
            EB_path = f"Projects\\{project}\\EB\\{phase}\\"
            Export_output_path = f"Projects\\{project}\\Exports\\"
            Export_output_file = f"{project}_DCS_export_TPStags.xlsx"
        # self.get_params_list(EB_path)
        self.write_EB(EB_path, Export_output_path)
        print(
            f"{Fore.MAGENTA}Finished creating EB files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def main():
    system("cls")
    project = DBtoEB("Test", "Original")


if __name__ == "__main__":
    main()
