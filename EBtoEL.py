# This program converts an EB file to a readable flat table.

# Import libraries
import os
import pandas as pd
import numpy as np
from routines import format_excel, check_folder
from os import system
from colorama import Fore


class EBtoEL:
    def __init__(self, project: str, phase: str = "Original"):
        self.start(project, phase)

    def get_EB(self, EB_path, Export_output_path, EL_output_file):
        EB = {}
        files = os.listdir(path=EB_path)

        pnttypes = []
        self.lines = []
        tags = []
        for file in files:
            if ".EB" in file.upper() and file != EL_output_file:
                print(f"reading:{file}")
            with open(EB_path + file, "r") as EBtext:
                self.lines += EBtext.readlines()
        self.lines = [s.replace("\x00", "") for s in self.lines]

        with open(Export_output_path + EL_output_file, "w") as file:
            for line in self.lines:
                if "&N" in line:
                    file.write(line[3:].replace(" ", ""))

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
            EL_output_file = "Test_EL.EL"
        else:
            EB_path = f"Projects\\{project}\\EB\\{phase}\\"
            Export_output_path = f"Projects\\{project}\\"
            EL_output_file = f"{project}_export_EL_{phase}.EL"

        self.get_EB(EB_path, Export_output_path, EL_output_file)

        print(
            f"{Fore.MAGENTA}Finished creating EB files for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def main():
    system("cls")
    project = EBtoEL("Test", "Optim")


if __name__ == "__main__":
    main()
