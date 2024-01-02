# Step 1:
# create tagmatch DCS - IPS (existing config)
# required:
# 	export DOC4000 TPS tags (existing config)
# 	export FSC (existing config)
#
# Step 2:
# create list of drawings (existing config)
# required:
# 	export DOC4000 display reference (existing config)
#
# Step 3:
# create list of CL files (existing config)
# required:
# 	export DOC4000 CL reference (existing config)
#
# Step 4:
# create tagmatch DCS - IP: (new config)
# required:
# 	export DOC4000 TPS tags (new config)
# 	export SM (new config)
#
# Step 5:
# Connect old/new IPS
# Compare old/new DCS

# # Import files description
# HMI import (from DOC4000):
#
# ! Do not use complete reference import as the export to XLS will be limited to max 65.536 rows, so filter on 'Native Window Display - tag' before exporting.
#
# CL import (from DOC4000): see example query below
#

# TODO export EB from AM and search for tags.

# import libraries
from typing import Any
import pandas as pd
import numpy as np
from simpledbf import Dbf5

# import win32com.client as win32
from routines import format_excel, show_df, ProjDetails, read_db  # ,ErrorLog
import json

# This function is used by the function readOldPLC and makes sure that when no shuffle has taken place, or only partly, the empty values are filled.


class iFATprepare:
    def __init__(self, project: str, debug: bool = False):
        self.debug = False
        self.start(project)

    def fill_shuffle(self, df: pd.DataFrame) -> pd.DataFrame:
        # fill in values that are not in shufflelist.
        df.OLD_PLC.replace("", np.nan, inplace=True)
        df.OLD_FLD.replace("", np.nan, inplace=True)
        df.NEW_PLC.replace("", np.nan, inplace=True)
        df.NEW_FLD.replace("", np.nan, inplace=True)

        df.OLD_PLC = df.PLC.where(pd.isna(df.OLD_PLC), df.OLD_PLC)
        df.OLD_FLD = df.SHEET.where(pd.isna(df.OLD_FLD), df.OLD_FLD)
        df.NEW_PLC = df.PLC.where(pd.isna(df.NEW_PLC), df.NEW_PLC)
        df.NEW_FLD = df.SHEET.where(pd.isna(df.NEW_FLD), df.NEW_FLD)

        return df

    def create_match_tag_new(self, df: pd.DataFrame, PLCname: str) -> pd.DataFrame:
        # this function creates a new column "MatchTag" for SM database

        df.insert(0, "PLC", PLCname)
        df["tempPT"] = df["PointType"]
        df.loc[df.PointType == "DI", "tempPT"] = "I"
        df.loc[df.PointType == "DO", "tempPT"] = "O"
        df.insert(3, "MatchTag", df.PLC + "_" + df.tempPT + "_" + df.TagNumber)
        del df["tempPT"]
        return df

    # The MatchTag is a combination of Tagname, pointtype and PLC, for matching the old and new tags.

    def create_match_tag_old(
        self, df: pd.DataFrame, PLCname: str, db_TagChange: pd.DataFrame
    ) -> pd.DataFrame:
        # this function creates a new column "MatchTag" for FSC database

        def func_tag_change(row: dict) -> str:
            # subfunction to change tagnames
            # in this case for both FSC and COM signals the "-" will be removed and _F and _DCS are removed
            temp_tag = row["NewTag"]

            # some tags are renamed and can be found in rename-database
            if temp_tag in db_TagChange.Old.values:
                x: Any = db_TagChange.loc[db_TagChange.Old == temp_tag, "New"]
                return x.iloc[0]

            # Disabled for FGS:
            if "FGS" not in PLCname:
                if row["LOC"] == "FSC" or row["LOC"] == "COM":
                    # removal of "-"
                    if len(temp_tag) >= 6:
                        if temp_tag[5] == "-":
                            temp_tag = temp_tag[0:5] + temp_tag[6:]
                    # removal of "_F" and "_DCS"
                    temp_tag = temp_tag.replace("_F", "")
                    temp_tag = temp_tag.replace("_DCS", "")
                    temp_tag = temp_tag.replace("_MOS", "MOS")
                    # replace 5-KOPxx with 5-KOP-xx:
            return temp_tag

        # if no NEW_PLC is filled in, use old PLC. Same for FLD (sheet), this is when a shuffle has taken place
        df["NEW_PLC"].fillna(df["PLC"], inplace=True)
        df["NEW_FLD"].fillna(df["SHEET"], inplace=True)
        df.insert(3, "NewTag", df.TAGNUMBER)
        df["NewTag"] = df.apply(lambda row: func_tag_change(row), axis=1)
        df.insert(3, "MatchTag", df.PLC + "_" + df.TYPE + "_" + df.TAGNUMBER)
        return df

    def combineBox(self, df):
        def func_box(row):
            if pd.isnull(row["BOXNUM"]):
                return row["OUTBOXNM"]
            else:
                return row["BOXNUM"]

        df["BOXNUM"] = df.apply(lambda row: func_box(row), axis=1)

        return df

    def compile_dcs(self, db_dcs: Any, plc_list: list) -> pd.DataFrame:
        def func_removeLC(row: dict) -> Any:
            try:
                return int(float(str(row["value"]).replace("!LC", "")))
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
            "Object Name",
            "PTDESC",
            "_TYPE",
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
                db_cur: Any = pd.melt(
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
                db_cur["PLC"] = cur_plc[1]
            else:
                if cur_plc[4] != 0:
                    # UCN
                    db_cur = pd.melt(
                        db_dcs[db_dcs["NODENUM"] == int(cur_plc[4])],
                        id_vars=id_params,
                        value_vars=ucn_params,
                    )
                    db_cur["PLC"] = cur_plc[1]
                    db_cur = db_cur[db_cur["NTWKNUM"].isin(cur_plc[6])]
                else:
                    continue

            db_cur = db_cur[(db_cur["value"] != "")]
            db_cur["value"] = db_cur.apply(lambda row: func_removeLC(row), axis=1)
            if self.debug:
                print(db_cur)

            if first:
                db_dcs_compiled: pd.DataFrame = db_cur
                first = False
            else:
                db_dcs_compiled = pd.concat([db_dcs_compiled, db_cur])
            # db_dcs_compiled['PLCLINK'] = db_dcs_compiled['PLC'] + db_dcs_compiled['NTWKNUM'].astype(str)

            db_dcs_compiled = self.combineBox(db_dcs_compiled)
            db_dcs_compiled.drop("OUTBOXNM", inplace=True, axis=1)

        return db_dcs_compiled

    def match_dcs(
        self, db_dcs_old: pd.DataFrame, db_plc_old: pd.DataFrame, plc_list: list
    ):
        def func_AIHG(row: dict) -> str:
            add = 0
            if row["_TYPE"] == "ANLINHG":
                add = 40000
            try:
                return row["PLC"] + "_" + str(int(row["value"]) + add)
            except:
                return row["PLC"] + "_"

        db_dcs_old["MatchMB"] = db_dcs_old.apply(lambda row: func_AIHG(row), axis=1)

        db_plc_old["MatchMB"] = ""
        cur_plc = plc_list[0]
        for cur_plc in plc_list:
            # determine if HWY or UCN
            params = []

            if cur_plc[2] != 0:
                mask = (
                    (db_plc_old["PLC"] == (cur_plc[1]))
                    & (db_plc_old["COM"] == int(cur_plc[5][0]))
                    & (db_plc_old["CHANNEL"] == cur_plc[5][1])
                )
                if self.debug:
                    print("mask:")
                    print(mask)
                cols = ["PLC", "ADDRESS"]
                db_plc_old.loc[mask, "MatchMB"] = db_plc_old.loc[mask, :][cols].apply(
                    lambda row: "_".join(row.values.astype(str)), axis=1
                )
            else:
                if cur_plc[4] != 0:
                    mask = (db_plc_old["PLC"] == cur_plc[1]) & (
                        db_plc_old["DCS_ADDR"] != -1
                    )
                    if self.debug:
                        print("mask:")
                        print(mask)
                    cols = ["PLC", "DCS_ADDR"]
                    db_plc_old.loc[mask, "MatchMB"] = db_plc_old.loc[mask, :][
                        cols
                    ].apply(lambda row: "_".join(row.values.astype(str)), axis=1)
                else:
                    continue

        # mask = (db_plc_old['COM'] == int(cur_plc[5][0])) & (db_plc_old['CHANNEL'] == cur_plc[5][1])
        mask = db_plc_old["MatchMB"] != ""
        # db_cur['value']= db_cur.apply (lambda row: func_removeLC(row), axis=1)
        # combined = pd.merge(DB_DCS_old, DB_old,how='inner',on='MatchMB')
        combined = pd.merge(
            db_dcs_old, db_plc_old[mask], how="outer", on="MatchMB", indicator="Exist"
        )

        return db_plc_old, db_dcs_old, combined

    def filter_COM_CHAN(self, IPS_remaining, old_plc_list):
        def func_filter(row):

            for PLC in old_plc_list:
                if (
                    int(row["COM"]) == int(PLC[5][0])
                    and row["CHANNEL"] == PLC[5][1]
                    and row["PLC_y"] == PLC[1]
                ):
                    return "DCS"
            return ""

        IPS_remaining.loc[:, "DCS_COM_CHAN"] = IPS_remaining[
            IPS_remaining["LOC"] == "COM"
        ].apply(lambda row: func_filter(row), axis=1)

        return IPS_remaining[IPS_remaining["DCS_COM_CHAN"] == "DCS"]

    # This function reads the old PLC database.
    # - If `shuffle_file` is provided, new PLC/FLD positions will be added to the database.
    # - If `tag_change_file` is provided, new PLC/FLD positions will be added to the database.

    def createCLOverview(self, df):
        def func_tag(row):
            if row["Output Object Type"] == "Tag":
                return row["Output Object Name"]
            else:
                return row["Input Object Name"]

        def func_CL(row):
            if row["Output Object Type"] == "Tag":
                return row["Input Object Name"]
            else:
                return row["Output Object Name"]

        df["Tagname"] = df.apply(lambda row: func_tag(row), axis=1)
        df["CL"] = df.apply(lambda row: func_CL(row), axis=1)

        return df[["Tagname", "CL"]]

    def read_old_plc(
        self,
        path: str,
        old_plc_list: list,
        shuffle_file: str,
        tag_change_file: str,
        dcs_export_old: str,
        hmi_export_old: str,
        cl_export_old: str,
        output_file: str,
    ):
        print("Reading old databases:")
        db_old = pd.DataFrame()
        if tag_change_file != "":
            print(f"- Reading {tag_change_file}")
            db_tag_change = read_db(path + "Extra\\", tag_change_file)
        else:
            db_tag_change = pd.DataFrame(columns=["Old", "New"])

        if shuffle_file != "":
            db_shuffle = read_db(path + "Extra\\", shuffle_file)
            print(f"- Reading {shuffle_file}")
        else:
            db_shuffle = pd.DataFrame(
                columns=["OLD_PLC", "OLD_FLD", "NEW_PLC", "NEW_FLD"]
            )

        print(f"- Reading {dcs_export_old}")
        db_dcs = read_db(path + "Exports\\", dcs_export_old)
        # to use an EB (combined) export, we have to modify 2 column names
        if "&N" in db_dcs.columns:
            db_dcs = db_dcs.rename(columns={"&N": "Object Name", "&T": "_TYPE"})

        params = [
            "PCADDRI1",
            "PCADDRI2",
            "PCADDRO1",
            "PLCADDR",
            "DISRC(1)",
            "DISRC(2)",
            "DODSTN(1)",
            "DODSTN(2)",
            "DODSTN(3)",
            "HWYNUM",
            "NODENUM",
            "NTWKNUM",
            "BOXNUM",
            "OUTBOXNM",
        ]
        for param in params:
            if param not in db_dcs.columns:
                db_dcs[param] = ""
        params = [
            "HWYNUM",
            "NODENUM",
            "NTWKNUM",
            "BOXNUM",
            "OUTBOXNM",
            "PNTBOXIN",
            "PNTBOXOT",
        ]

        for param in params:
            db_dcs[param].replace("", np.nan, inplace=True)
        db_dcs[params] = db_dcs[params].astype(float).astype("Int32")

        first = True
        for curPLC in old_plc_list:
            print(f"- Reading {curPLC[0]}")
            db_old_temp = read_db(path + "PLCs\\Original\\", curPLC[0])
            db_old_temp.insert(0, "PLC", curPLC[1])
            db_old_temp = pd.merge(
                db_old_temp,
                db_shuffle,
                how="left",
                left_on=["PLC", "SHEET"],
                right_on=["OLD_PLC", "OLD_FLD"],
            )

            db_old_temp = self.create_match_tag_old(
                db_old_temp, curPLC[1], db_tag_change
            )
            if first:
                db_old: pd.DataFrame = db_old_temp
                first = False
            else:
                db_old = pd.concat([db_old, db_old_temp])

        db_old = self.fill_shuffle(db_old)

        print(f"- Compiling DCS")
        db_dcs_old = self.compile_dcs(db_dcs, old_plc_list)
        if self.debug:
            with pd.ExcelWriter(path + "debug.xlsx") as writer:
                print("\n- Writing db_dcs_old")
                db_dcs_old.to_excel(writer, sheet_name="db_dcs_old", index=False)
                print("\n- Writing db_old")
                db_old.to_excel(writer, sheet_name="db_old", index=False)
                print("\n- Writing db_dcs")
                db_dcs.to_excel(writer, sheet_name="db_dcs", index=False)

        print(f"- matching tags")
        db_old, db_dcs_old, db_combined = self.match_dcs(
            db_dcs_old, db_old, old_plc_list
        )

        print(f"- Reading {hmi_export_old}")
        db_ref = read_db(path + "Exports\\", hmi_export_old)
        db_displays: Any = pd.DataFrame
        if db_ref.iloc[1, 7] == "Native Window Display":
            db_displays = db_ref[db_ref["Input Object Type"] == "Native Window Display"]
        else:
            db_displays = db_ref[db_ref["Input Object Type"] == "HMIWebDisplay"]
        db_tags_unique = (
            db_combined["Object Name"].drop_duplicates().dropna().sort_values()
        )
        db_result = pd.merge(
            db_displays,
            db_tags_unique,
            how="right",
            on="Object Name",
            indicator="Exist",
        )
        db_hmi_unique = (
            db_result["Input Object Name"].drop_duplicates().dropna().sort_values()
        )

        print(f"- Reading {cl_export_old}")
        db_cl = read_db(path + "Exports\\", cl_export_old)
        db_cl_match = pd.merge(
            db_cl, db_tags_unique, how="right", on="Object Name", indicator="Exist"
        )
        db_cl_match = db_cl_match[db_cl_match["Exist"] == "both"].drop("Exist", axis=1)
        db_cl_match = self.createCLOverview(db_cl_match)

        # DB_old = DB_old.drop_duplicates(subset ="MatchTag", keep = 'first')
        with pd.ExcelWriter(path + output_file) as writer:
            print("\n- Writing old PLC")
            db_old.to_excel(writer, sheet_name="PLC_old", index=False)
        with pd.ExcelWriter(path + output_file, engine="openpyxl", mode="a") as writer:
            if not db_tag_change.empty:
                print("- Writing TagChange")
                db_tag_change.to_excel(writer, sheet_name="TagChange", index=False)
            if not db_shuffle.empty:
                print("- Writing shuffle list")
                db_shuffle.to_excel(writer, sheet_name="ShuffleList", index=False)
            print("- Writing DCS Export Old")
            db_dcs.to_excel(writer, sheet_name="DCS_old", index=False)
            print("- Writing Compiled DCS Export Old")
            db_dcs_old.sort_values("Object Name", axis=0).to_excel(
                writer, sheet_name="DCS_old_compiled", index=False
            )
            print("- Writing Combined Old")
            db_combined[db_combined["Exist"] == "both"].drop("Exist", axis=1).to_excel(
                writer, sheet_name="Combined", index=False
            )
            print("- Writing remaining DCS Old")
            db_combined[db_combined["Exist"] == "left_only"].drop(
                "Exist", axis=1
            ).to_excel(writer, sheet_name="Remaining DCS Old", index=False)
            print("- Writing remaining IPS Old")
            self.filter_COM_CHAN(
                db_combined[db_combined["Exist"] == "right_only"].drop("Exist", axis=1),
                old_plc_list,
            ).to_excel(writer, sheet_name="Remaining IPS Old", index=False)
            print("- Writing HMI old")
            db_result[db_result["Exist"] == "both"].drop("Exist", axis=1).to_excel(
                writer, sheet_name="HMI Old", index=False
            )
            print("- Writing HMI unique old")
            db_hmi_unique.to_excel(writer, sheet_name="HMI unique old", index=False)
            print("- Writing tags unique old")
            db_tags_unique.to_excel(writer, sheet_name="tags unique Old", index=False)
            print("- Writing tags with no HMI")
            db_excel: Any = db_result["Object Name"][db_result["Exist"] == "right_only"]
            db_excel.to_excel(writer, sheet_name="Old - tags not in HMI", index=False)
            print("- Writing CL old per tag")
            db_cl_match.to_excel(writer, sheet_name="CL old per tag", index=False)
            print("- Writing unique CL")
            db_cl_match[["CL"]].drop_duplicates().sort_values("CL", axis=0).to_excel(
                writer, sheet_name="CL old unique", index=False
            )
            print("Reading old database ready\n")
        return (
            db_old,
            db_dcs_old,
            db_combined[db_combined["Exist"] == "both"].drop("Exist", axis=1),
        )

    def check_i_fat(
        self,
        path: str,
        output_file: str,
        old_plc_list: list,
        shuffle_file: str,
        tag_change_file: str,
        old_dcs_export: str,
        old_hmi_export: str,
        old_cl_export: str,
    ):
        db_old, db_dcs, db_combined = self.read_old_plc(
            path,
            old_plc_list,
            shuffle_file,
            tag_change_file,
            old_dcs_export,
            old_hmi_export,
            old_cl_export,
            output_file,
        )
        print("Formatting excel..")
        format_excel(path, output_file)
        print("Done!")
        return db_old, db_dcs, db_combined

    def start(self, project):
        debug = False

        my_proj = ProjDetails(project)

        my_path = my_proj.path
        outputFile = my_proj.outputFile["prep"]
        oldPLCs = my_proj.PLCs_prep
        shuffleList = my_proj.shuffleList  # ignored if "" - placed in path+"Extra\\"
        # ignored if "" - placed in path+"Extra\\"
        tagChangeList = my_proj.tagChangeList
        DCSExport = my_proj.DCSExport  # - placed in path+"Exports\\"
        HMIExport = my_proj.HMIExport  # - placed in path+"Exports\\"
        CLExport = my_proj.CLExport  # - placed in path+"Exports\\"

        df_PLC_old, df_DCS_old, df_combined = self.check_i_fat(
            my_path,
            outputFile,
            oldPLCs,
            shuffleList,
            tagChangeList,
            DCSExport,
            HMIExport,
            CLExport,
        )

    # in case of assertion error opening file, save as XLSX in stead of XLS


def main():
    project = iFATprepare("SWS6")


if __name__ == "__main__":
    main()
