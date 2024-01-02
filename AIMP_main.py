from colorama import Fore
from os import system
from CheckFSC2SM import CheckFSC2SM
from CreateAIMPxlsx import CreateAIMPxlsx
from EBtoDB import EBtoDB
from iFATprepare import iFATprepare
from IOCount import IOCount
from SPIMatch import SPIMatch
from XLSMerge import XLSMerge
from XLSXEBtoXLSXtotal import XLSXEBtoXLSXtotal


def main():

    project = "CD6"
    phase = "Optim"
    proj_date = "2022-01-01"
    author = "Pascal van de Wijdeven"

    system("cls")
    print(
        f"{Fore.MAGENTA}Creating files for "
        f"{Fore.GREEN}{project}{Fore.MAGENTA}, "
        f"phase {Fore.GREEN}{phase}{Fore.RESET}"
    )

    # CheckFSC2SM(project, phase, False)
    # CreateAIMPxlsx(project=project, phase=phase, proj_date=proj_date, author=author)
    # EBtoDB(project=project, phase=phase)
    # iFATprepare(project=project, debug=False)
    # IOCount(project=project, phase=phase)
    # SPIMatch(project=project, phase=phase)
    # XLSMerge(project=project, phase=phase, isPLC=True)
    # XLSXEBtoXLSXtotal(project=project, phase=phase)

    print(
        f"{Fore.MAGENTA}Finished files for "
        f"{Fore.GREEN}{project}{Fore.MAGENTA}, "
        f"phase {Fore.GREEN}{phase}{Fore.RESET}"
    )


if __name__ == "__main__":
    main()
