# import libraries
from math import log10
from copy import deepcopy
from distutils.log import warn
from sre_compile import isstring

from matplotlib import offsetbox
import pandas as pd
import numpy as np
from pyparsing import col
from simpledbf import Dbf5
from sqlalchemy import true
from routines import (
    format_excel,
    quick_excel,
    show_df,
    ProjDetails,
    read_db,
    file_exists,
    dprint,
)  # ,ErrorLog
from colorama import Fore, Back
from os import system
import typing
from pandasgui import show as show_df

ERROR = Fore.WHITE + Back.RED
RESET = Fore.RESET + Back.RESET

ALARM_VALUES = {"EMERGNCY": 5, "HIGH": 4, "LOW": 3, "JOURNAL": 2, "NOACTION": 1, "": 0}
ALARM_PRIOS = {5: "EMERGENCY", 4: "HIGH", 3: "LOW", 2: "JOURNAL", 1: "NOACTION", 0: ""}
ADDR_PARAMS = {
    "PLCADDR": "PCADDRI1",
    "DISRC(1)": "PCADDRI1",
    "DISRC(2)": "PCADDRI2",
    "DODSTN(1)": "PCADDRO1",
    "DODSTN(2)": "PCADDRO2",
}


def is_float(element: typing.Any) -> bool:
    try:
        float(element)
        return True
    except ValueError:
        return False


class EPLCG2SMM:
    def __init__(self, project: str, phase0: str = "Original", phase1: str = "Optim"):
        self.rulespath = ""
        self.eb_old = {}
        self.eb_new = {}
        self.main = {}
        self.pnttypes = []
        self.DCSmaster = {}
        self.start(project, phase0, phase1)

    def get_typicals(self, proj) -> pd.DataFrame:
        try:
            print("- Loading Typicals...")
            return pd.read_excel(f"Projects\\{proj}\\EB\\Rules\\Typicals.xlsx")
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"'Typicals.xlsx' not found in "
                f"'Projects\\{proj}\\EB\\Rules\\'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def get_rules(self, proj) -> pd.DataFrame:
        try:
            print("- Loading Rules...")
            df_temp = pd.read_excel(
                f"Projects\\{proj}\\EB\\Rules\\EPLCG2SMM_rules.xlsx",
            )
            return df_temp.fillna("")
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"'EPLCG2SMM_rules.xlsx' not found in "
                f"'Projects\\{proj}\\EB\\Rules\\'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def Dig_rules(self, proj, sheet) -> typing.Any:
        try:
            print("- Loading Rules...")
            df_temp = pd.read_excel(
                f"Projects\\{proj}\\EB\\Rules\\EPLCG2SMM_DDS_CTL_2.xlsx",
                sheet_name=sheet,
            )
            for sheet in df_temp:
                df_temp[sheet].fillna("", inplace=True)
            return df_temp
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"'EPLCG2SMM_DDS_CTL_2.xlsx' not found in "
                f"'Projects\\{proj}\\EB\\Rules\\'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def get_eb_files(self, proj, phase) -> typing.Dict[str, pd.DataFrame]:
        try:
            print(f"- Loading EB files {phase}...")
            my_eb = pd.read_excel(
                f"Projects\\{proj}\\EB\\{phase}\\{proj}_export_EB_total_{phase}.xlsx",
                sheet_name=None,
            )
            for sheet in my_eb:
                my_eb[sheet].fillna(value="", inplace=True)
            return my_eb
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"'{proj}_export_EB_total_{phase}.xlsx' not found in "
                f"'Projects\\{proj}\\EB\\{phase}\\'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def get_old_plc(self, proj) -> pd.DataFrame:
        try:
            print(f"- Loading sheet 'combined' from preparation file")
            return pd.read_excel(
                f"Projects\\{proj}\\{proj}_preparation.xlsx",
                sheet_name="Combined",
            )
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"{proj}_preparation.xlsx' not found in "
                f"'Projects\\{proj}'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def get_new_plc(self, proj, phase) -> pd.DataFrame:
        try:
            print(f"- Loading new PLCs file")
            return pd.read_excel(
                f"Projects\\{proj}\\PLCs\\{phase}\\Merged\\{proj}_total_PLCs_{phase}.xlsx",
            )
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"{proj}_total_PLCs_{phase}' not found in "
                f"'Projects\\{proj}\\PLCs\\{phase}\\Merged\\'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def get_PLC_match(self, proj, phase) -> pd.DataFrame:
        try:
            print(f"- Loading PLC match file")
            return pd.read_excel(
                f"Projects\\{proj}\\{proj}_PLCmatch_{phase}.xlsx",
            )
        except FileNotFoundError:
            print(
                f"{ERROR}ERROR: file "
                f"{proj}_PLCmatch_{phase}.xlsx' not found in "
                f"'Projects\\{proj}'"
            )
            exit(f"ABORTED: File not found{RESET}")

    def get_DCS_PLC_new(self):
        def func_check(row, param):
            if row["DCS_address_old"] == param:

                try:
                    for i in range(1, 12):
                        # return row[self.DCSmaster[row["PLC"]]]
                        if row[f"Master{i}"] == "EUCN":
                            return row[f"PLCAddress{i}"]
                except Exception as e:
                    return
            else:
                return

        plc_old = self.main["plc_old"].copy()
        plc_old = plc_old.set_index(["TYPE", "TAGNUMBER"])
        plc_match = self.main["plc_match"].copy()
        plc_match = plc_match.rename(
            columns={"TYPE_OLD": "TYPE", "TAGNUMBER_OLD": "TAGNUMBER"}
        )
        plc_match = plc_match.set_index(["TYPE", "TAGNUMBER"])
        result = pd.merge(
            plc_old,
            plc_match,
            left_index=True,
            right_index=True,
            how="outer",
        )
        result.reset_index(inplace=True)
        columns = [
            "Object Name",
            "variable",
            "TAGNUMBER",
            "TYPE",
            "TAGNUMBER_NEW",
            "TYPE_NEW",
        ]
        result = result[columns].dropna(inplace=False)
        columns = [
            "DCS_tag",
            "DCS_address_old",
            "PLC_tag_old",
            "PLC_type_old",
            "PLC_tag_new",
            "PLC_type_new",
        ]
        result.columns = columns

        result = pd.merge(
            result,
            self.main["plc_new"],
            right_on=["TagNumber", "PointType"],
            left_on=["PLC_tag_new", "PLC_type_new"],
            how="outer",
        )

        for address in ADDR_PARAMS.keys():
            param = ADDR_PARAMS[address]
            result[param] = result.apply(lambda row: func_check(row, param), axis=1)

        return result

    def getDCSMaster(self, PLCs):
        self.DCSmaster = {}
        for PLC in PLCs:
            plcname = PLC["PLCName"]
            address = PLC["details"]["eUCNParams"]["DCS"]
            self.DCSmaster[plcname] = f"PLCAddress{address}"

    def add_old(self):
        for EB in self.eb_old:
            cols = self.eb_old[EB].columns
            cols = [f"{param}_old" for param in cols]
            self.eb_old[EB].columns = cols

    def add_typical(self, df, index, typical, params):

        # !! "!!!" in typical means 'NOT' - this is
        # !! not done in original typical rules and avoids double assignments
        def func_typical(row, cur_params):

            for par in cur_params:
                typ_par = typical[par]
                if len(str(typ_par)) >= 4:
                    if str(typ_par)[0:3] == "!!!":
                        typ_par = typ_par[4:]
                        if typ_par == row[par]:
                            return 0
                        else:
                            continue
                # if row["&N"] == "602FZ021":
                #     dprint(
                #         f"{par}: typical = {typ_par} tag = {row[par]}",
                #         "IMPORTANT",
                #     )
                #     if typ_par != row[par]:
                #         dprint("NOT OK", "IMPORTANT")
                #         return 0
                #     else:
                #         dprint("OK", "IMPORTANT")

                # else:
                if typ_par != row[par]:
                    return 0

            return 1

        print(f"Assigning {Fore.GREEN}{typical['Typical']}{Fore.RESET}...")
        cur_params = []
        for param in params:
            curpar = typical[param]
            if curpar != "*" and curpar != "_":
                cur_params.append(param)

        pd.options.mode.chained_assignment = None
        df.loc[:, f"{typical['Typical']}"] = (
            df[:]
            .copy(deep=True)
            .apply(lambda row: func_typical(row, cur_params), axis=1)
        )
        pd.options.mode.chained_assignment = "warn"

        return df

    def assign_typicals(self):

        params = list(self.typ.columns)
        params.remove("Typical")
        params.remove("Rule")
        params.remove("PNTTYPE_NEW")
        params = params
        names = pd.DataFrame
        names = self.eb_old["TOTAL"][["&N"] + params]
        typicals = []
        for index, row in self.typ.iterrows():
            names = self.add_typical(names, index, row, params)
            typicals.append(row["Typical"])
        pd.options.mode.chained_assignment = None
        names["Typical_count"] = names.loc[:, typicals].sum(axis=1).copy(deep=True)

        if names[names["Typical_count"] != 1].size > 0:
            quick_excel(names, self.rulespath, "DEBUG")
            dprint("None or non-single typical assignment detected!!!", "Important")
            exit(
                f"{Fore.RED}Aborted - none or non-single typical assignment - check 'DEBUG.xlsx'{Fore.RESET}"
            )
        names["Typical"] = ""
        for typical in typicals:
            names["Typical"][names[typical] == True] = typical
        pd.options.mode.chained_assignment = "warn"
        names2 = names[["&N", "Typical"]].copy(deep=True)

        return names2, names

    def combine_old_new(self, df, phase0, phase1):
        df = pd.merge(
            df,
            self.eb_new["TOTAL"]["&N"],
            how="outer",
            on="&N",
            indicator=True,
            suffixes=(f"_{phase0}", f"_{phase1}"),
        )
        df_result = {}

        df_result["Combined"] = df[["&N", "Typical"]][df["_merge"] == "both"]
        df_result["Deleted"] = df[["&N", "Typical"]][df["_merge"] == "left_only"]
        df_result["New"] = df[["&N", "Typical"]][df["_merge"] == "right_only"]

        return df_result

    def moveColumn(self, df, column_ref, column_move):
        cols = df.columns.tolist()
        cols.insert(cols.index(column_ref) + 1, cols.pop(cols.index(column_move)))
        df = df[cols]
        df.columns = cols
        return df

    def ignore_rule(self, typical, index, row):
        def func_check(row, param):
            if row[param] == "":
                return "OK"
            else:
                return f"'' expected (not applicable)"

        param = row["parameter"]
        if param in self.main[f"{typical}_check"]:
            self.main[f"{typical}_check"][f"{param}_check"] = self.main[
                f"{typical}_check"
            ].apply(lambda row: func_check(row, param), axis=1)

            self.main[f"{typical}_check"] = self.moveColumn(
                self.main[f"{typical}_check"], f"{param}", f"{param}_check"
            )

    def link1_1(self, typical, index, row):
        def func_check(row, param):

            param_old = row[f"{param}_old"]
            if set(str(param_old)) == " ":
                param_old = ""
            if row[param] == param_old:
                return "OK"
            else:
                return f"{str(param_old)} expected (rule 1 on 1)"

        param = row["parameter"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_old"
        )

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"{param}_old", f"{param}_check"
        )

    def link1_1_round(self, typical, index, row):
        def func_check(row, param):
            param_old = float(row[f"{param}_old"])
            max_range = max(
                abs(float(row["PVEUHI_old"])), abs(float(row["PVEULO_old"]))
            )
            decimals = int(max(0, round(2.5 - (log10(max_range)), 0)))
            if decimals == 0:
                rounded = int(round(param_old, 0))
            else:
                rounded = round(param_old, decimals)

            if not is_float(row[param]):
                if row[param] == rounded:
                    return "OK"
                else:
                    return f"{rounded} expected (rule 1 on 1 with rounding - {decimals} decimals)"

            if float(row[param]) == rounded:
                return "OK"
            else:
                return f"{rounded} expected (rule 1 on 1 with rounding - {decimals} decimals)"

        param = row["parameter"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_old"
        )
        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"{param}_old", f"{param}_check"
        )

    def link1_1_round_TP(self, typical, index, row):
        def func_check(row, param: str):
            param_EU = param.replace("PV", "PVEU").replace("TP", "")
            param_old = float(row[f"{param}_old"])
            max_range = max(
                abs(float(row["PVEUHI_old"])), abs(float(row["PVEULO_old"]))
            )
            decimals = int(max(0, round(2.5 - (log10(max_range)), 0)))
            if decimals == 0:
                rounded_old = int(round(param_old, 0))
                rounded_EU = int(round(float(row[f"{param_EU}_old"]), 0))
            else:
                rounded_old = round(param_old, decimals)
                rounded_EU = round(float(row[f"{param_EU}_old"]), decimals)
            if "LO" in param:
                if rounded_old <= rounded_EU:
                    if row[param] == "--------":
                        return "OK"
                    else:
                        return f"'--------' expected ({param} outside range - {param_EU}: {rounded_EU})"

            if "HI" in param:
                if rounded_old >= rounded_EU:
                    if row[param] == "--------":
                        return "OK"
                    else:
                        return f"'-----' expected ({param} outside range - {param_EU}: {rounded_EU})"

            if not is_float(row[param]):
                if row[param] == rounded_old:
                    return "OK"
                else:
                    return f"{rounded_old} expected (rule 1 on 1 with rounding - {param_EU}: {rounded_EU})"

            if float(row[param]) == rounded_old:
                return "OK"
            else:
                return f"{rounded_old} expected (rule 1 on 1 with rounding - {param_EU}: {rounded_EU})"

        param = row["parameter"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_old"
        )
        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"{param}_old", f"{param}_check"
        )

    def link1_1_conditional_TP(self, typical, index, row):
        def func_check(row, param):
            # first check if new param exists
            param_available = param in row

            if "HI" in param:
                if row["PVHITP"] != "--------":
                    if not param_available:
                        return f"PVHITP defined, PVHIPR parameter expected"
                    param_old = row[f"{param}_old"]
                    if row[param] == param_old:
                        return f"OK"
                    else:
                        return f"{param_old} expected (rule 1 on 1)"
                else:
                    if not param_available:
                        return "OK"
                    if row[param] != "":
                        return f"Empty expected, no PVHITP defined"
                    else:
                        return "OK"

            if "LO" in param:
                if row["PVLOTP"] != "--------":
                    if not param_available:
                        return f"PVLOTP defined, PVLOPR parameter expected"
                    param_old = row[f"{param}_old"]
                    if row[param] == param_old:
                        return "OK"
                    else:
                        return f"{param_old} expected (rule 1 on 1)"
                else:
                    if not param_available:
                        return "OK"
                    if row[param] != "":
                        return f"Empty expected, no PVLOTP defined"
                    else:
                        return "OK"

            return f"Something weird happened here, param: {param}"

        param = row["parameter"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_old"
        )
        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"{param}_old", f"{param}_check"
        )

    def setDefault_conditional_TP(self, typical, index, row):
        def func_check(row, param):
            # first check if new param exists
            param_available = param in row

            if "HI" in param:
                if row["PVHITP"] != "--------":
                    if not param_available:
                        return f"{param} expected but not available"
                    if row[param] == "CUTOUT":
                        return f"OK"
                    else:
                        return f"'CUTOUT' expected (rule 1 on 1)"
                else:
                    if not param_available:
                        return "OK"
                    if row[param] != "":
                        return f"Empty expected, no PVHITP defined"
                    else:
                        return "OK"

            if "LO" in param:
                if row["PVLOTP"] != "--------":
                    if not param_available:
                        return f"{param} expected but not available"
                    if row[param] == "CUTOUT":
                        return f"OK"
                    else:
                        return f"'CUTOUT' expected (rule 1 on 1)"
                else:
                    if not param_available:
                        return "OK"
                    if row[param] != "":
                        return f"Empty expected, no PVLOTP defined"
                    else:
                        return "OK"

            return f"Something weird happened here, param: {param}"

        param = row["parameter"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_check"
        )

    def getPVRAW(self, typical, index, row_rule):
        def func_check(row, param):
            if param == "PVRAWHI":
                if row["PLC_type_new"] == "BO":
                    if row["PVEUHI"] == row[param]:
                        return "OK"
                    else:
                        return (
                            f"{row['PVEUHI']} expected (same as PVEUHI), source is BO"
                        )
                else:

                    if str(row[param]) == "3276" or str(row[param]) == "3276.0":
                        return "OK"
                    else:
                        return f"3276 expected (source is AI) - '{str(row[param])}' - {row['&N']}"
            if param == "PVRAWLO":
                if row["PLC_type_new"] == "BO":
                    if row["PVEULO"] == row["PVRAWLO"]:
                        return "OK"
                    else:
                        return (
                            f"{row['PVEULO']} expected (same as PVEULO), source is BO"
                        )
                else:
                    if str(row["PVRAWLO"]) == "655" or str(row["PVRAWLO"]) == "655.0":
                        return "OK"
                    else:
                        return f"655 expected (source is AI) - '{str(row['PVRAWLO'])}'"
            return f"Something went wrong with {param}"

        param = row_rule["parameter"]
        self.main[f"{typical}_check"][f"{param}_check"] = pd.merge(
            self.main[f"{typical}_check"],
            self.main["DCS_PLC_new"],
            left_on="&N",
            right_on="DCS_tag",
            suffixes=["", "_PLC"],
            how="left",
        ).apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_check"
        )

    def checkAlloc(self, typical, index, row_rule):
        def func_check(row, param):
            if row[param] == "":
                return "value expected"
            else:
                return "OK"

        param = row_rule["parameter"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_check"
        )

    def checkPLCADDR(self, typical, index, row_rule):
        def func_check(row, param):

            if param == "PLCADDR":
                try:
                    if str(int(row[ADDR_PARAMS[param]])) == str(int(row[param])):
                        return "OK"
                    else:
                        return f"{str(int(row[ADDR_PARAMS[param]]))} expected"
                except:
                    if row[param] != "":
                        return "EMPTY expected"
            else:
                try:
                    address = "!LC" + str(int(row[ADDR_PARAMS[param]]))
                    if address == row[param]:
                        return "OK"
                    else:
                        return f"{address} expected"
                except:
                    if row[param] != "":
                        return "EMPTY expected"

        param = row_rule["parameter"]
        temp = self.main["DCS_PLC_new"][
            ["DCS_tag", "PCADDRI1", "PCADDRI2", "PCADDRO1", "PCADDRO2"]
        ]
        temp = temp.groupby(["DCS_tag"]).agg(sum)

        self.main[f"{typical}_check"][f"{param}_check"] = pd.merge(
            self.main[f"{typical}_check"],
            temp,
            left_on="&N",
            right_on="DCS_tag",
            suffixes=["", "_PLC"],
            how="left",
        ).apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_check"
        )

    def setDefault(self, typical, index, row):
        def func_check(row, param, default):
            # && round of floats
            if is_float(default):
                if float(default).is_integer():
                    default = str(int(float(default)))

            cur_val = row[param]

            if is_float(cur_val):
                if float(cur_val).is_integer():
                    cur_val = str(int(float(cur_val)))

            if str(cur_val) == default:
                return "OK"
            else:
                return f"'{default}' expected (rule 2 - defaults)"

        param = row["parameter"]
        if param not in self.main[f"{typical}_check"]:
            return
        default = row["Default"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param, default), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_check"
        )

    def checkDefwCheck(self, typical, index, row):
        def func_check(row, param, default, param_available):

            extra = ""
            if row[f"{param}_old"] != default:
                extra = f" Check original! '{row[f'{param}_old']}' not default ('{default}')"
            if not param_available:
                if extra != "":
                    return extra
                else:
                    return "OK"
            if row[param] == default:
                return "OK" + extra
            else:
                return f"'{default}' expected (rule 98 - defaults)" + extra

        param = row["parameter"]
        default = row["Default"]
        param_available = param in self.main[f"{typical}_check"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param, default, param_available), axis=1)
        if param_available:
            self.main[f"{typical}_check"] = self.moveColumn(
                self.main[f"{typical}_check"], param, f"{param}_check"
            )
        else:
            self.main[f"{typical}_check"] = self.moveColumn(
                self.main[f"{typical}_check"], "&T", f"{param}_check"
            )

    def checkBadPV(self, typical, index, row):
        def func_check(row, param, ruleID):
            highest = 1
            badpvpr = ""
            if ruleID == "6l":
                for PR_param in row.keys():
                    if len(PR_param) > 2:
                        if PR_param[-2:] == "PR" and PR_param != "BADPVPR":
                            if row[PR_param] == row[PR_param]:
                                highest = max(highest, ALARM_VALUES[row[PR_param]])

                badpvpr = ALARM_PRIOS[highest]
                if row[param] == badpvpr:
                    return "OK"
                else:
                    return f"'{badpvpr}' expected (rule 6l - highest conf alarm)"
            elif ruleID == "6d":
                badpvpr = row["OFFNRMPR"]
                if row[param] == badpvpr:
                    return "OK"
                else:
                    return f"'{badpvpr}' expected (rule 6d - same as OFFNRMPR)"

        param = row["parameter"]
        ruleID = row["Processing_ID"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param, ruleID), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_check"
        )

        return

    def link1_1HG(self, typical, index, row):
        def func_check(row, param, paramHG):
            param_old = row[f"{paramHG}_old"]
            if row[param] == param_old:
                return "OK"
            else:
                return f"{param_old} expected (rule 1 on 1)"

        param = row["parameter"]
        paramHG = row["Link_par"]
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param, paramHG), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{paramHG}_old"
        )

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"{paramHG}_old", f"{param}_check"
        )

    def alarms_DIGINHG(self, typical):
        # params: ALMOPT, OFFNRMPR, PVNORMAL

        def func_check_ALMOPT(row, param):
            param_old = row[f"DIGALFMT_old"]
            if param_old == "STATE1" or param_old == "STATE2":
                if row[param] == "OFFNORML":
                    return "OK"
                else:
                    return f"'OFFNoRML' expected"
            if param_old == "CHNGOFST":
                if row[param] == "CHNGOFST":
                    return "OK"
                else:
                    return f"'CHNGOFST' expected"
            return "None of the checked options are valid"

        def func_check_PVNORMAL(row, param):
            param_old = row[f"DIGALFMT_old"]
            if param_old == "STATE1":
                if row[param] == row["STATE2_old"]:
                    return "OK"
                else:
                    return f"'{row['STATE2_old']}' expected"
            if param_old == "STATE2":
                if row[param] == row["STATE1_old"]:
                    return "OK"
                else:
                    return f"'{row['STATE1_old']}' expected"
            if param_old == "CHNGOFST":
                if row[param] == "":
                    return "OK"
                else:
                    return f"'' (empty) expected"
            return "None of the checked options are valid"

        def func_check_OFFNRMPR(row, param):
            param_old = row[f"DIGALFMT_old"]
            if param_old == "STATE1" or param_old == "STATE2":
                if row[param] == row["OFFNRMPR_old"]:
                    return "OK"
                else:
                    return f"'{row['OFFNRMPR_old']}' expected"
            if param_old == "CHNGOFST":
                if row[param] == row["CHOFSTPR_old"]:
                    return "OK"
                else:
                    return f"'{row['CHOFSTPR_old']}' expected"
            return "None of the checked options are valid"

        param = "ALMOPT"
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check_ALMOPT(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"DIGALFMT_old"
        )

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"DIGALFMT_old", f"{param}_check"
        )

        param = "PVNORMAL"
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check_PVNORMAL(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"PVNORMAL", f"{param}_check"
        )

        param = "OFFNRMPR"
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check_OFFNRMPR(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], "OFFNRMPR", f"OFFNRMPR_old"
        )
        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], "OFFNRMPR_old", f"CHOFSTPR_old"
        )

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], f"CHOFSTPR_old", f"{param}_check"
        )

    def checkAI_INPTDIR(self, typical):
        def func_check(row):
            if row["TopScale"] < row["BottomScale"]:
                if row["INPTDIR"] != "REVERSE":
                    return "REVERSE expected"
                else:
                    return "OK"
            else:
                if row["INPTDIR"] == "REVERSE":
                    return "DIRECT expected"
                else:
                    return "OK"

        param = "INPTDIR"
        self.main[f"{typical}_check"][f"{param}_check"] = pd.merge(
            self.main[f"{typical}_check"],
            self.main["DCS_PLC_new"],
            left_on="&N",
            right_on="DCS_tag",
            suffixes=["", "_PLC"],
            how="left",
        ).apply(lambda row: func_check(row), axis=1)

        # column already exists (rule 2 - so no need to move)
        # // self.main[f"{typical}_check"] = self.moveColumn(
        # //   self.main[f"{typical}_check"], param, f"{param}_check"
        # //)
        return

    def checkMOS_OFFNRMPR(self, typical):
        def func_check(row, param):
            if row[f"{param}_old"] != "" and row[f"{param}_old"] != "NOACTION":
                return f"CHECK old param: {row[f'{param}_old']}"
            else:
                if row[param] != "" and row[param] != "NOACTION":
                    return "EMPTY or NOACTION expected"
                else:
                    return "OK"

        param = "OFFNRMPR"
        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        self.main[f"{typical}_check"] = self.moveColumn(
            self.main[f"{typical}_check"], param, f"{param}_check"
        )

    def checkEIP(self, typical, index, row):
        def func_check(row, param):
            if param == "EIPPCODE":
                param_old = row[f"{param}_old"]
                if param_old == "--" or param_old == "":
                    if param in row:
                        if row[param] != "":
                            "No EIP expected"
                        else:
                            return "OK"
                    else:
                        return "OK"
                else:
                    if param in row:
                        if row[param] == param_old:
                            return "OK"
                        else:
                            return f"Check if EIP required"
                    else:
                        return f"Check if EIP required"
            elif param == "EVTOPT":
                if "EIPPCODE" in row:
                    if row[f"EIPPCODE"] != "":
                        if row[param] != "EIP":
                            return "'EIP' expected"
                        else:
                            return "OK"
                    else:
                        if row[param] != "NONE":
                            return "'NONE' expected"
                        else:
                            return "OK"
                else:
                    if row[param] != "NONE":
                        return "'NONE' expected"
                    else:
                        return "OK"
            return "something went wrong here"

        param = row["parameter"]

        self.main[f"{typical}_check"][f"{param}_check"] = self.main[
            f"{typical}_check"
        ].apply(lambda row: func_check(row, param), axis=1)

        if (
            param in self.main[f"{typical}_check"]
            and f"{param}_old" in self.main[f"{typical}_check"]
        ):
            self.main[f"{typical}_check"] = self.moveColumn(
                self.main[f"{typical}_check"], param, f"{param}_old"
            )
            self.main[f"{typical}_check"] = self.moveColumn(
                self.main[f"{typical}_check"], f"{param}_old", f"{param}_check"
            )
        elif param in self.main[f"{typical}_check"]:
            self.main[f"{typical}_check"] = self.moveColumn(
                self.main[f"{typical}_check"], param, f"{param}_check"
            )

    def checkDO_rules(self, typical):
        def func_check(row, param):
            param_rules = row[f"{param}_rules"]
            try:
                if str(param_rules)[0] == "[":
                    param_rules = row[param_rules[1:-1] + "_old"]
            except:
                pass
            if row[param] == param_rules:
                return "OK"
            else:
                return f"{param_rules} expected (DO rules)"

        rules = self.main["DigOut_rules"]
        rules = rules[rules["Typical"] == typical]
        match_params = rules.columns[1:4].tolist()
        check_params = rules.columns[4:].to_list()

        df_check = pd.merge(
            self.main[f"{typical}_check"],
            rules,
            left_on=[sub + "_old" for sub in match_params],
            right_on=match_params,
            suffixes=["", "_rules"],
            how="left",
        )

        for param in check_params:
            if f"{param}_rules" in df_check:
                self.main[f"{typical}_check"][f"{param}_check"] = df_check.apply(
                    lambda row: func_check(row, param), axis=1
                )

                self.main[f"{typical}_check"] = self.moveColumn(
                    self.main[f"{typical}_check"], f"{param}", f"{param}_check"
                )

    def checkDI_rules(self, typical):
        def func_check(row, param):
            # dprint(param, "IMPORTANT")
            param_rules = row[f"{param}_rules"]
            try:
                if str(param_rules)[0] == "[":
                    param_rules = row[param_rules[1:-1] + "_old"]
            except:
                pass
            if row[param] == param_rules:
                return "OK"
            else:
                return f"{param_rules} expected (DI rules)"

        rules = self.main["DigIn_rules"]
        rules = rules[rules["Typical"] == typical]
        match_params = rules.columns[1:3].tolist()
        check_params = rules.columns[3:-1].to_list()

        df_check = pd.merge(
            self.main[f"{typical}_check"],
            rules,
            left_on=[sub + "_old" for sub in match_params],
            right_on=match_params,
            suffixes=["", "_rules"],
            how="left",
        )

        for param in check_params:
            self.main[f"{typical}_check"][f"{param}_check"] = df_check.apply(
                lambda row: func_check(row, param), axis=1
            )

            self.main[f"{typical}_check"] = self.moveColumn(
                self.main[f"{typical}_check"], f"{param}", f"{param}_check"
            )

    def check_typical(self, typical, pnttype_old, pnttype_new):
        # setting up typical check dataframe
        dprint(f"Checking {typical}", "GREEN")

        self.main[f"{typical}_check"] = pd.merge(
            self.main["Combined"][self.main["Combined"]["Typical"] == typical],
            self.eb_new[pnttype_new],
            on="&N",
            how="inner",
        )

        self.main[f"{typical}_check"]["DUMP"] = "DUMP_COLUMN"

        self.main[f"{typical}_check"] = pd.merge(
            self.main[f"{typical}_check"],
            self.eb_old[pnttype_old],
            left_on="&N",
            right_on="&N_old",
        )

        # Ignoring unused typicals
        if self.main[f"{typical}_check"].size == 0:
            dprint(f"{typical} is unused", "RED")
            self.main.pop(f"{typical}_check", None)
            return False

        # rule0 - ignore
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "0")
                | (self.rules["Processing_ID"] == "3a")
                | (self.rules["Processing_ID"] == "3b")
                | (self.rules["Processing_ID"] == "3c")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)
        for index, row in cur_check.iterrows():
            self.ignore_rule(typical, index, row)

        # rule1 = link 1-1
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "1")
                | (self.rules["Processing_ID"] == "1t")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)
        for index, row in cur_check.iterrows():
            self.link1_1(typical, index, row)

        # rule1a/b/e = link 1-1 with rounding
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "1a")
                | (self.rules["Processing_ID"] == "1b")
                | (self.rules["Processing_ID"] == "1e")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)
        for index, row in cur_check.iterrows():
            self.link1_1_round(typical, index, row)

        # rule1c/d = link 1-1 with rounding or -----
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "1c")
                | (self.rules["Processing_ID"] == "1d")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)
        for index, row in cur_check.iterrows():
            self.link1_1_round_TP(typical, index, row)

        # rule1h/j = link 1-1 if TP <> -----
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "1h")
                | (self.rules["Processing_ID"] == "1j")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)
        for index, row in cur_check.iterrows():
            self.link1_1_conditional_TP(typical, index, row)

        # rule1i/k = DEFAULT if TP <> -----
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "1i")
                | (self.rules["Processing_ID"] == "1k")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)
        for index, row in cur_check.iterrows():
            self.setDefault_conditional_TP(typical, index, row)

        # rule1f/g = PVRAW from IPS
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "1f")
                | (self.rules["Processing_ID"] == "1g")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)
        for index, row in cur_check.iterrows():
            self.getPVRAW(typical, index, row)

        # rule2 = set default
        cur_check = self.rules[
            (self.rules["Processing_ID"] == "2") & (self.rules["Typical"] == typical)
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.setDefault(typical, index, row)

        # rule3abc = Allocation
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "3a")
                | (self.rules["Processing_ID"] == "3b")
                | (self.rules["Processing_ID"] == "3c")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.checkAlloc(typical, index, row)

        # # rule3defgh = PLCADDR
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "3d")
                | (self.rules["Processing_ID"] == "3e")
                | (self.rules["Processing_ID"] == "3f")
                | (self.rules["Processing_ID"] == "3g")
                | (self.rules["Processing_ID"] == "3h")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.checkPLCADDR(typical, index, row)

        # rule4: link 1-1 with HG param
        cur_check = self.rules[
            ((self.rules["Processing_ID"] == "4") & (self.rules["Typical"] == typical))
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.link1_1HG(typical, index, row)

        # rule5: EIP
        cur_check = self.rules[
            (
                (
                    (self.rules["Processing_ID"] == "5a")
                    | (self.rules["Processing_ID"] == "5b")
                )
                & (self.rules["Typical"] == typical)
            )
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.checkEIP(typical, index, row)

        # rule6l = badpv highest + 6d = badpv pvnormalpr
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "6l")
                | (self.rules["Processing_ID"] == "6d")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.checkBadPV(typical, index, row)

        # rule8abcde = digital inputs
        cur_check = self.rules[
            (
                (self.rules["Processing_ID"] == "8a")
                | (self.rules["Processing_ID"] == "8b")
                | (self.rules["Processing_ID"] == "8c")
                | (self.rules["Processing_ID"] == "8d")
                | (self.rules["Processing_ID"] == "8e")
            )
            & (self.rules["Typical"] == typical)
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.checkDI_rules(typical)

        # rule98 = default with checking
        cur_check = self.rules[
            (self.rules["Processing_ID"] == "98") & (self.rules["Typical"] == typical)
        ].copy(deep=True)

        for index, row in cur_check.iterrows():
            self.checkDefwCheck(typical, index, row)

        # ** rule6* - rule does not follow database!
        if typical == "DIGINHG" or typical == "DC11ST":
            self.alarms_DIGINHG(typical)

        # && Extra checks

        # ANLINHG INPTDIR
        if typical == "ANLINHG":
            self.checkAI_INPTDIR(typical)

        # MOS OFFNRMPR
        if typical == "MOS":
            self.checkMOS_OFFNRMPR(typical)

        if typical in list(self.main["DigOut_rules"]["Typical"]):
            self.checkDO_rules(typical)

        # find last 'DUMP' column, then remove all remaining '_old'
        columns = self.main[f"{typical}_check"].columns
        keep_columns = []
        for column in columns:
            if column == "DUMP":
                break
            keep_columns.append(column)
        self.main[f"{typical}_check"] = self.main[f"{typical}_check"][keep_columns]

        return True

    def is_in_rules(self, typical, param):
        result = self.rules[
            (self.rules["Typical"] == typical) & (self.rules["parameter"] == param)
        ]
        return result.empty

    def check_all(self):
        for index, row in self.typ.iterrows():
            if "IGNORE" not in str(row["Typical"]):
                typical = row["Typical"]
                if self.check_typical(typical, row["&T"], row["PNTTYPE_NEW"]):
                    cols = self.main[f"{typical}_check"].columns.tolist()
                    for col in cols:
                        if "_old" not in col and "_check" not in col:
                            if f"{col}_check" not in cols:
                                if col not in ["&N", "&T", "Typical", "Source"]:
                                    if self.is_in_rules(typical, col):
                                        self.ignore_rule(typical, 0, {"parameter": col})
                                        result = self.main[f"{typical}_check"][
                                            f"{col}_check"
                                        ]
                                        if len(set(result)) > 1:
                                            dprint(
                                                f"Values found in not checked {col} for {typical}",
                                                "IMPORTANT",
                                            )
                                    else:
                                        dprint(
                                            f"{col} not checked for {typical}",
                                            "IMPORTANT",
                                        )

    def main_def(self, proj, phase0, phase1) -> typing.Any:
        my_proj = ProjDetails(proj)

        newPLCs = my_proj.PLCs["Migrated"]

        self.rulespath = f"Projects\\{proj}\\EB\\Rules\\"
        self.typ = self.get_typicals(proj)
        self.typ = self.typ.rename(columns={"PNTTYPE": "&T"})
        self.rules = self.get_rules(proj)
        self.eb_old = self.get_eb_files(proj, phase0)
        self.eb_new = self.get_eb_files(proj, phase1)
        self.names, self.typicals = self.assign_typicals()
        self.main = self.combine_old_new(self.names, phase0, phase1)
        # gui = show_df(self.main["Combined"])
        # !! Following requires having run iFATprepare succesfully to combine old DCS/IPS
        self.main["plc_old"] = self.get_old_plc(proj)
        self.main["TypicalAss."] = self.typicals
        # !! Following requires having database with matching old/new PLC tag+type combi
        # !! see CD6 as (manually produced) example
        self.main["plc_match"] = self.get_PLC_match(proj, phase1)
        self.getDCSMaster(newPLCs)
        self.main["plc_new"] = self.get_new_plc(proj, phase1)
        self.main["DCS_PLC_new"] = self.get_DCS_PLC_new()
        self.main["EB_old"] = self.eb_old["TOTAL"]
        self.main["EB_new"] = self.eb_new["TOTAL"]
        self.main["Typicals"] = self.typ
        self.main["Rules"] = self.rules
        self.main["DigIn_rules"] = self.Dig_rules(proj, "DigIn").fillna("")
        self.main["DigAlarm_rules"] = self.Dig_rules(proj, "DigAlarm").fillna("")
        self.main["DigOut_rules"] = self.Dig_rules(proj, "DigOut").fillna("")
        self.add_old()

        # !! Make sure PNTTYPE_NEW is added to typicals file
        self.check_all()

        quick_excel(self.main, f"Projects\\{proj}\\EB\\Rules\\", "Combined")

        # gui = show_df(self.main["MOS_check"])
        # gui = show_df(self.main["DC11ST_check"])
        filename = "check_file.xlsx"
        with pd.ExcelWriter(self.rulespath + filename, mode="w") as writer:
            for sheet in self.main:
                # if "_check" in sheet:
                dprint(f"writing sheet {sheet}", "YELLOW")
                self.main[sheet].to_excel(writer, sheet_name=sheet, index=False)

        format_excel(
            self.rulespath,
            filename,
            first_time=True,
            different_red=True,
            different_blue=False,
            check_existing_red=False,
            how="check",
        )

        return

    def start(self, project, phase0, phase1):

        print(
            f"{Fore.MAGENTA}Comparing HG to UCN for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase1}{Fore.RESET}"
        )

        self.main_def(project, phase0, phase1)
        print(
            f"{Fore.MAGENTA}Finished Comparing HG to UCN for {Fore.GREEN}{project}{Fore.MAGENTA}, phase {Fore.GREEN}{phase1}{Fore.RESET}"
        )


# TODO add rules 98/99 checks
# TODO EIP check
# TODO pre-filter EB-files
# TODO fix when column is missing (e.g. PVLOPR SWS6)


def main():
    system("cls")
    project = EPLCG2SMM("PSU_JSON", "Original", "FAT")


if __name__ == "__main__":
    main()
