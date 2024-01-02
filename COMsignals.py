# import libraries

from EBtoDB import EBtoDB
from XLSMerge import XLSMerge
import pandas as pd
import numpy as np
from routines import (
    format_excel,
    quick_excel,
    show_df,
    ProjDetails,
    read_db,
    file_exists,
    dprint,
)
from colorama import Fore, Back
from os import system
import typing
from pandasgui import show as show_df

ERROR = Fore.WHITE + Back.RED
RESET = Fore.RESET + Back.RESET


class COMsignals:
    def __init__(self, project: str, phase0: str = "Original", phase1: str = "Optim"):
        self.start(project, phase0, phase1)

    def get_eb_files(self, proj, phase) -> pd.DataFrame:
        def add_columns(df, sheet, COM):
            if sheet.upper() == "TOTAL":
                columns = list(self.all_columns[COM][sheet]) + ["NODENUM", "NTWKNUM"]
                for column in columns:
                    if column != "" and column not in df[sheet].columns:
                        dprint(f"Empty column {column} added to {sheet}", "GREEN")
                        df[sheet][column] = ""
            return df

        if phase != "Original":
            if not file_exists(self.projpath + f"EB\\{phase}"):
                dprint(f"No EB files for phase {phase}", "IMPORTANT")
                return pd.DataFrame()
        if not file_exists(
            self.projpath + f"EB\\{phase}\\{proj}_export_EB_total_{phase}.xlsx"
        ):
            EBtoDB(proj, phase, False)
        try:
            dprint(f"- Loading EB files {phase}...", "CYAN")
            my_eb = pd.read_excel(
                self.projpath + f"\\EB\\{phase}\\{proj}_export_EB_total_{phase}.xlsx",
                sheet_name=None,
            )
            sheet = "TOTAL"
            if phase == "Original":
                add_columns(my_eb, sheet, "HG")
            else:
                add_columns(my_eb, sheet, "eUCN")
            my_eb[sheet].fillna(value="", inplace=True)
            return my_eb["TOTAL"]
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"'{proj}_export_EB_total_{phase}.xlsx' not found in "
                f"'Projects\\{proj}\\EB\\{phase}\\'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def get_all_columns(self, proj) -> dict:
        print("- Loading all columns...")
        df_temp = {}
        try:
            df_temp["HG"] = pd.read_excel(
                self.rulespath + "All_columns_HG.xlsx",
            ).fillna("")

        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"'All_columns_HG.xlsx' not found in "
                f"'{self.rulespath}'"
            )
            exit(f"ABORTED: File not found{RESET}")

        try:
            df_temp["eUCN"] = pd.read_excel(
                self.rulespath + "All_columns_eUCN.xlsx",
            ).fillna("")

        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"'All_columns_eUCN.xlsx' not found in "
                f"'{self.rulespath}'"
            )
            exit(f"ABORTED: File not found{RESET}")

        return df_temp

    def plc_names(self, proj):
        dict_PLC = {}
        dict_old = {}

        for PLC in proj.PLCs["Original"]:
            dict_old[PLC["PLCName"]] = PLC["details"]["PLCFile"].split(".")[0]
        for PLC in proj.PLCs["Migrated"]:
            dict_PLC[dict_old[PLC["PLCName"]]] = PLC["details"]["PLCFile"].split(".")[0]
        print(dict_PLC)
        return dict_PLC

    def get_plc(self, proj, phase) -> pd.DataFrame:
        if not file_exists(
            self.projpath + f"PLCs\\{phase}\\Merged\\{proj}_total_PLCs_{phase}.xlsx"
        ):
            XLSMerge(proj, phase, True, False)

        try:
            print(f"- Loading {phase} PLCs file")
            if phase == "Original":
                sn = "FSC"
            else:
                sn = "SM"

            return pd.read_excel(
                self.projpath
                + f"PLCs\\{phase}\\Merged\\{proj}_total_PLCs_{phase}.xlsx",
                sheet_name=sn,
            )
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"{proj}_total_PLCs_{phase}' not found in "
                f"'Projects\\{proj}\\PLCs\\{phase}\\Merged\\'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def combineBox(self, df):
        def func_box(row):
            if pd.isnull(row["BOXNUM"]):
                return row["OUTBOXNM"]
            else:
                return row["BOXNUM"]

        df["BOXNUM"] = df.apply(lambda row: func_box(row), axis=1)

        return df

    def compile_dcs(self, db_dcs, plc_list: list) -> pd.DataFrame:
        def func_removeLC(row: dict):
            try:
                return int(float(str(row["value"]).replace("!LC", "")))
            except:
                return ""

        def func_addAI(row: dict):
            try:
                if row["&T"] == "ANLINHG":
                    return row["value"] + 40000
                else:
                    return row["value"]
            except:
                return ""

        hwy_params = ["PCADDRI1", "PCADDRI2", "PCADDRO1", "PCADDRO2"]
        ucn_params = [
            "PLCADDR",
            "DISRC(1)",
            "DISRC(2)",
            "DODSTN(1)",
            "DODSTN(2)",
            "DODSTN(3)",
        ]
        id_params = [
            "&N",
            "PTDESC",
            "&T",
            "HWYNUM",
            "NODENUM",
            "NTWKNUM",
            "BOXNUM",
            "OUTBOXNM",
        ]
        first = True
        db_dcs_compiled = pd.DataFrame()

        for cur_plc in plc_list:
            # determine if HWY or SMM
            params = []
            if int(cur_plc[2]) != 0:

                # HWY
                db_cur = pd.melt(
                    db_dcs[
                        (db_dcs["HWYNUM"] == int(cur_plc[2]))
                        & (
                            (db_dcs["PNTBOXIN"] == int(cur_plc[3]))
                            | (db_dcs["PNTBOXOT"] == int(cur_plc[3]))
                        )
                    ],
                    id_vars=id_params,
                    value_vars=hwy_params,
                )
                db_cur["matchPLC"] = cur_plc[1]
            else:
                if cur_plc[4] != 0:
                    dprint(f"{cur_plc} is UCN", "BLUE")
                    # UCN
                    db_cur = pd.melt(
                        db_dcs[db_dcs["NODENUM"] == int(cur_plc[4])],
                        id_vars=id_params,
                        value_vars=ucn_params,
                    )
                    db_cur["matchPLC"] = cur_plc[1]
                    db_cur = db_cur[db_cur["NTWKNUM"].isin(cur_plc[6])]
                else:
                    continue

            db_cur = db_cur[(db_cur["value"] != "")]
            db_cur["value"] = db_cur.apply(lambda row: func_removeLC(row), axis=1)
            db_cur["PLCvalue"] = db_cur.apply(lambda row: func_addAI(row), axis=1)

            if first:
                db_dcs_compiled: pd.DataFrame = db_cur
                first = False
            else:
                db_dcs_compiled = pd.concat([db_dcs_compiled, db_cur])
            # db_dcs_compiled['PLCLINK'] = db_dcs_compiled['PLC'] + db_dcs_compiled['NTWKNUM'].astype(str)

            db_dcs_compiled = self.combineBox(db_dcs_compiled)
            db_dcs_compiled.drop("OUTBOXNM", inplace=True, axis=1)

        return db_dcs_compiled

    def match_HG(self):
        def func_checkLoop(row: dict):
            try:
                tag_DCS = row["&N"]
                tag_PLC = row["TAGNUMBER"]
                if tag_DCS[:4] + tag_DCS[-3:] == tag_PLC[:4] + tag_PLC[-3:]:
                    return "OK"
                else:
                    return f"Loop mismatch {tag_DCS[:4] + tag_DCS[-3:]} vs {tag_PLC[:4] + tag_PLC[-3:]}"
            except:
                if pd.isna(row["&N"]) or pd.isna(row["TAGNUMBER"]):
                    return ""
                else:
                    return "Something went wrong"

        def func_checkRemark(row: dict):
            try:
                if row["LoopCheck"] != "OK":
                    if row["Remark"] != "":
                        return row["Remark"] + " + check loopname"
                    else:
                        return "Check loopname"
                else:
                    return row["Remark"]
            except:
                return "Something went wrong"

        # TODO when SM HG is added as option (e.g. HV8), make sure there is a "ADDRESS" column added to the PLC file.

        df = self.main["EB_old"].merge(
            right=self.main["PLC_old_COM"],
            how="outer",
            left_on=["matchPLC", "PLCvalue"],
            right_on=["PLC", "ADDRESS"],
            indicator=True,
        )

        df["Remark"] = ""
        mask = (df["_merge"] == "both") & (
            df.duplicated(subset=["matchPLC", "PLCvalue"], keep=False)
        )
        df["LoopCheck"] = df.apply(lambda row: func_checkLoop(row), axis=1)

        df.loc[mask, "Remark"] = "Duplicate"
        mask = df["_merge"] == "left_only"
        df.loc[mask, "Remark"] = "Only in DCS, not in PLC"
        mask = df["_merge"] == "right_only"
        df.loc[mask, "Remark"] = "Only in PLC, not in DCS"

        df["Remark"] = df.apply(lambda row: func_checkRemark(row), axis=1)

        return df

    def filter_plc_old(self, PLCs, phase):

        df = pd.DataFrame()
        first = True
        columns = [
            "PLC",
            "TYPE",
            "TAGNUMBER",
            "SERVICE",
            "LOC",
            "SHEET",
            "COM",
            "CHANNEL",
            "ADDRESS",
        ]

        for PLC in PLCs:
            mask = (
                (self.main["PLC_old"]["PLC"] == PLC[1])
                & (self.main["PLC_old"]["COM"] == int(PLC[5][0]))
                & (self.main["PLC_old"]["CHANNEL"] == PLC[5][1])
            )
            if first:
                df: pd.DataFrame = self.main["PLC_old"][mask][columns]
                first = False
            else:
                df = pd.concat([df, self.main["PLC_old"][mask][columns]])
        return df

    def SM_address(self, df) -> pd.DataFrame:
        def func_eUCN_address(row: dict):
            try:
                for col in col_list:
                    if row[col] == "EUCN":
                        return row["PLCAddress" + col[-1:]]
            except:
                return "Something went wrong", "ohoh"

        def func_eUCN_comport(row: dict):
            try:
                for col in col_list:
                    if row[col] == "EUCN":
                        return col
            except:
                return "Something went wrong", "ohoh"

        col_list = []
        for col in df.columns:
            if "Master" in col:
                col_list.append(col)
        df["SM_address"] = df.apply(lambda row: func_eUCN_address(row), axis=1)
        df["COMport"] = df.apply(lambda row: func_eUCN_comport(row), axis=1)
        columns = [
            "PLC",
            "PointType",
            "TagNumber",
            "Description",
            "Location",
            "FLDNumber",
            "Description",
            "SM_address",
            "COMport",
        ]
        df = df[df["COMport"].notna()]
        return df[columns]

    def match_PLCs(self, df_old, df_new, d_plc_names) -> pd.DataFrame:
        def func_typeSM(row: dict):
            if row["TYPE"] == "I":
                return "DI"
            elif row["TYPE"] == "O":
                return "DO"
            else:
                return row["TYPE"]

        def func_plcSM(row: dict):
            return d_plc_names[row["PLC"]]

        df_old["TYPE_SM"] = df_old.apply(lambda row: func_typeSM(row), axis=1)
        df_old["PLCmatch"] = df_old.apply(lambda row: func_plcSM(row), axis=1)

        df = df_old.merge(
            right=df_new,
            how="outer",
            left_on=["PLCmatch", "TYPE_SM", "TAGNUMBER"],
            right_on=["PLC", "PointType", "TagNumber"],
            indicator=True,
        )

        return df

    def main_def(self, proj, phase0, phase1) -> typing.Any:
        my_proj = ProjDetails(proj)
        d_plc_names = self.plc_names(my_proj)
        self.rulespath = f"Projects\\Rules\\"
        self.all_columns = self.get_all_columns(proj)
        # newPLCs = my_proj.PLCs["Migrated"]
        self.projpath = f"Projects\\{proj}\\"
        self.main = {}
        self.main["EB_old"] = self.compile_dcs(
            self.get_eb_files(proj, phase0), my_proj.PLCs_prep
        )
        self.main["PLC_old"] = self.get_plc(proj, phase0)
        self.main["PLC_old_COM"] = self.filter_plc_old(my_proj.PLCs_prep, phase0)
        self.main["Original_match"] = self.match_HG()

        self.main["EB_new"] = self.get_eb_files(proj, phase1)
        if self.main["EB_new"].empty:
            del self.main["EB_new"]

        self.main["PLC_new"] = self.get_plc(proj, phase1)
        self.main["PLC_new_COM"] = self.SM_address(self.main["PLC_new"])
        self.main["PLC_match"] = self.match_PLCs(
            self.main["PLC_old_COM"], self.main["PLC_new_COM"], d_plc_names
        )

        filename = f"{proj}_COM_signals_{phase1}.xlsx"
        with pd.ExcelWriter(self.projpath + filename, mode="w") as writer:
            for sheet in self.main:
                dprint(f"writing sheet {sheet}", "YELLOW")
                self.main[sheet].to_excel(writer, sheet_name=sheet, index=False)

        format_excel(
            self.projpath,
            filename,
            first_time=True,
        )

        return

    def start(self, project, phase0, phase1):

        print(
            f"{Fore.MAGENTA}Matching HG with eUCN for COM signals{Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase1}{Fore.RESET}"
        )

        self.main_def(project, phase0, phase1)
        print(
            f"{Fore.MAGENTA}Finished matching HG with eUCN for COM signals for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase1}{Fore.RESET}"
        )


# TODO .............


def main():
    system("cls")
    project = COMsignals("PSU_JSON", "Original", "FAT")


if __name__ == "__main__":
    main()
