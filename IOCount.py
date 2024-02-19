# import libraries
import pandas as pd
import numpy as np
from os import system, listdir
from simpledbf import Dbf5
from routines import (
    ErrorLog,
    format_excel,
    check_folder,
    show_df,
    read_db,
    dprint,
    file_exists,
)
from colorama import Fore, Back
import typing

pd.reset_option("mode.chained_assignment")

ERROR = Fore.WHITE + Back.RED
RESET = Fore.RESET + Back.RESET


class IOCount:
    def __init__(
        self, project: str, phase: str = "Original", proxes=["GBN", "GBO"], format=True
    ):
        self.start(project, phase, proxes, format)

    def get_PLCs(self, input_path, output_file):
        all_PLCs = []

        if all_PLCs == []:
            files = listdir(path=input_path)
            for file in files:
                if file != output_file and "." in file:
                    all_PLCs.append([file, file.split(".", 1)[0]])

        return all_PLCs

    def read_data(self, PLCs, input_path):

        lst_df = []
        for curPLC in PLCs:
            df_temp = read_db(input_path, curPLC[0])
            df_temp.insert(0, "PLC", curPLC[1])
            df_temp.rename(columns={"ChassisID/IOTAName": "RACK"}, inplace=True)
            df_temp.rename(columns={"Location": "LOC"}, inplace=True)
            df_temp.rename(columns={"PointType": "TYPE"}, inplace=True)
            df_temp.rename(columns={"SlotNumber": "POS"}, inplace=True)

            lst_df.append(df_temp)
        df = pd.concat(lst_df)

        return df

    def is_prox(self, df: pd.DataFrame, prox_name: list):
        unique_before = df.drop_duplicates(subset=["CARDNAME"], keep="first").size

        def func_prox(row):
            if (
                (prox_name[0] in row["TAGNUMBER"])
                or (prox_name[1] in row["TAGNUMBER"])
                and (row["TYPE"] == "I")
            ):
                return "I (PROX)"
            return row["TYPE"]

        df_temp: typing.Any = df.copy()

        df_temp["TYPE"] = df_temp.apply(lambda row: func_prox(row), axis=1)
        df_temp.loc[:, "CARDNAME"] = (
            df_temp["TYPE"].astype(str)
            + "_"
            + df_temp["PLC"].astype(str)
            + "_"
            + df_temp["RACK"].astype(str)
            + "_"
            + df_temp["POS"].astype(str)
        )
        unique_after = df_temp.drop_duplicates(subset=["CARDNAME"], keep="first").size

        if unique_after == unique_before:
            print(f"{Fore.GREEN}Message: prox cards determinded properly{Fore.RESET}")
            return df_temp
        else:
            print(
                f"{Fore.RED}Warning: unable to get prox cards amounts properly{Fore.RESET}"
            )
            return df

    def calculate_IO(self, input_path, output_path, output_file, proxes, format):

        PLCs = self.get_PLCs(input_path, output_file)
        df: typing.Any = self.read_data(PLCs, input_path)

        df_IO: typing.Any = df[
            ((df["RACK"] != "") & (df["RACK"] != 0) & (df["LOC"] != "SYS"))
        ]
        lst_types = df_IO[["PLC", "TYPE"]].drop_duplicates()
        IO_counts_PLC = df_IO[["PLC", "TYPE"]].value_counts()
        IO_counts_PLC = IO_counts_PLC.reset_index()
        IO_counts_PLC.rename(columns={0: "AMOUNT"}, inplace=True)

        IO_counts = df_IO[["TYPE"]].value_counts()
        IO_counts = IO_counts.reset_index()
        IO_counts.rename(columns={0: "AMOUNT"}, inplace=True)

        df_temp = pd.DataFrame(
            [["Total", IO_counts.sum()["AMOUNT"]]], columns=["TYPE", "AMOUNT"]
        )
        IO_counts = IO_counts.append(df_temp)
        IO_counts.sort_values(["TYPE"], inplace=True)

        for PLC in PLCs:
            df_temp = pd.DataFrame(
                [
                    [
                        PLC[1],
                        "Total",
                        IO_counts_PLC[IO_counts_PLC["PLC"] == PLC[1]].sum()["AMOUNT"],
                    ]
                ],
                columns=["PLC", "TYPE", "AMOUNT"],
            )
            IO_counts_PLC = IO_counts_PLC.append(df_temp)

        IO_counts_PLC.sort_values(["PLC", "TYPE"], inplace=True)

        cols = ["PLC", "RACK", "POS"]

        with pd.option_context("mode.chained_assignment", None):
            df_IO.loc[:, "CARDNAME"] = (
                df_IO[["PLC", "RACK", "POS"]].astype(str).agg("_".join, axis=1)
            )

        # df_IO.loc[:, "CARDNAME"] = (
        #     df_IO["PLC"].astype(str)
        #     + "_"
        #     + df_IO["RACK"].astype(str)
        #     + "_"
        #     + df_IO["POS"].astype(str)
        # )

        df_IO = self.is_prox(df_IO, proxes)
        df_cards = df_IO.drop_duplicates(subset=["CARDNAME"], keep="first")
        IO_card_counts: typing.Any = pd.DataFrame(df_cards["TYPE"].value_counts())
        IO_card_counts = IO_card_counts.reset_index()
        IO_card_counts.rename(columns={"index": "TYPE", "TYPE": "AMOUNT"}, inplace=True)
        df_temp = pd.DataFrame(
            [["Total", IO_card_counts.sum()["AMOUNT"]]], columns=["TYPE", "AMOUNT"]
        )
        IO_card_counts = IO_card_counts.append(df_temp)
        IO_card_counts.sort_values(["TYPE"], inplace=True)

        IO_card_counts_PLC = df_cards[["PLC", "TYPE"]].value_counts()
        IO_card_counts_PLC = IO_card_counts_PLC.reset_index()
        IO_card_counts_PLC.rename(columns={0: "AMOUNT"}, inplace=True)
        IO_card_counts_PLC.sort_values(["PLC", "TYPE"], inplace=True)

        for PLC in PLCs:
            df_temp = pd.DataFrame(
                [
                    [
                        PLC[1],
                        "Total",
                        IO_card_counts_PLC[IO_card_counts_PLC["PLC"] == PLC[1]].sum()[
                            "AMOUNT"
                        ],
                    ]
                ],
                columns=["PLC", "TYPE", "AMOUNT"],
            )
            IO_card_counts_PLC = IO_card_counts_PLC.append(df_temp)

        IO_card_counts_PLC.sort_values(["PLC", "TYPE"], inplace=True)

        check_folder(output_path)

        with pd.ExcelWriter(output_path + output_file) as writer:
            dprint("- Writing IO Counts per PLC", "YELLOW")
            IO_counts_PLC.to_excel(writer, sheet_name="IO Counts per PLC", index=False)
            dprint("- Writing IO Counts Overview", "YELLOW")
            IO_counts.to_excel(writer, sheet_name="IO Counts Overview", index=False)
            dprint("- Writing IO Counts per PLC", "YELLOW")
            IO_card_counts_PLC.to_excel(
                writer, sheet_name="IO Cards per PLC", index=False
            )
            dprint("- Writing IO Cards Overview", "YELLOW")
            IO_card_counts.to_excel(writer, sheet_name="IO Cards Overview", index=False)
        if format:
            dprint("- Formatting file", "YELLOW")
            format_excel(output_path, output_file)

    def start(self, project, phase, proxes, format):

        print(
            f"{Fore.MAGENTA}Creating IO overview for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )
        input_path = f"Projects\\{project}\\PLCs\\{phase}\\"
        output_path = f"Projects\\{project}\\IO_count\\"
        output_file = f"{project}_IO_count_{phase}.xlsx"
        self.calculate_IO(input_path, output_path, output_file, proxes, format)
        print(
            f"{Fore.MAGENTA}Done creating IO overview for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase}{Fore.RESET}"
        )


def get_IO_count(proj, phase) -> dict:
    projpath = f"Projects\\{proj}\\"
    filename = projpath + f"IO_count\\{proj}_IO_count_{phase}.xlsx"
    if not file_exists(filename):
        IOCount(proj, phase, format=False)
    try:
        dprint(f"- Loading IO count files {phase}", "CYAN")
        my_IO_counts = pd.read_excel(
            filename,
            sheet_name=None,
        )
        return my_IO_counts
    except FileNotFoundError:
        print(f"{ERROR}ERROR: file {filename} not found")
        exit(f"ABORTED: File not found{RESET}")


def main():
    system("cls")
    project = IOCount("PGPMODB")


if __name__ == "__main__":
    main()
