# import libraries
import pandas as pd
import numpy as np
from simpledbf import Dbf5
from routines import format_excel, ProjDetails, read_db  # ,ErrorLog
from colorama import Fore
from pandasgui import show as show_df


class CheckFSC2SM:
    def __init__(self, project: str, phase: str, debug: bool = False):
        self.debug = debug
        self.start(project, phase)

    def createMatchTagNew(self, df, PLCname):
        # this function creates a new column "MatchTag" for SM database

        df.insert(0, "PLC", PLCname)
        df["tempPT"] = df["PointType"]
        df.loc[df.PointType == "DI", "tempPT"] = "I"
        df.loc[df.PointType == "DO", "tempPT"] = "O"
        df.insert(3, "MatchTag", df.PLC + "_" + df.tempPT + "_" + df.TagNumber)
        del df["tempPT"]
        return df

    def createMatchTagOld(self, df, PLCname, db_TagChange):
        # this function creates a new column "MatchTag" for FSC database

        def func_TagChange(row):
            # subfunction to change tagnames
            # in this case for both FSC and COM signals the "-" will be removed and _F and _DCS are removed
            temptag = row["NewTag"]

            # some tags are renamed and can be found in rename-database
            if temptag in db_TagChange.Old.values:
                x = db_TagChange.loc[db_TagChange.Old == temptag, "New"]
                return x.iloc[0]

            # Disabled for FGS:
            if False:  # not "FGS" in PLCname:
                if row["LOC"] == "FSC" or row["LOC"] == "COM":
                    # removal of "-"
                    if len(temptag) >= 6:
                        if temptag[5] == "-":
                            temptag = temptag[0:5] + temptag[6:]
                    # removal of "_F" and "_DCS"
                    temptag = temptag.replace("_F", "")
                    temptag = temptag.replace("_DCS", "")
                    # replace 5-KOPxx with 5-KOP-xx:
            return temptag

        # if no NEW_PLC is filled in, use old PLC. Same for FLD (sheet), this is when a shuffle has taken place
        df["NEW_PLC"].fillna(df["PLC"], inplace=True)
        df["NEW_FLD"].fillna(df["SHEET"], inplace=True)
        df.insert(3, "NewTag", df.TAGNUMBER)
        df["NewTag"] = df.apply(lambda row: func_TagChange(row), axis=1)
        df.insert(3, "MatchTag", df.NEW_PLC + "_" + df.TYPE + "_" + df.NewTag)
        return df

    def moveColumn(self, df, column_old, column_new):
        cols = df.columns.tolist()
        cols.insert(cols.index("DUMP"), cols.pop(cols.index(column_old)))
        df = df[cols]
        cols[cols.index(column_old)] = column_new
        df.columns = cols
        return df

    def fillShuffle(self, df):
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

    def fillIO(self, df):
        # fill in values that are not in shufflelist.
        df.IO_Cabinet_Old.replace("", np.nan, inplace=True)
        df.IO_Cabinet_New.replace("", np.nan, inplace=True)
        df.IO_RACK_Old.replace("", np.nan, inplace=True)
        df.IO_POS_Old.replace("", np.nan, inplace=True)
        df.IO_CHAN_Old.replace("", np.nan, inplace=True)
        df.IO_RACK_New.replace("", np.nan, inplace=True)
        df.IO_POS_New.replace("", np.nan, inplace=True)
        df.IO_CHAN_New.replace("", np.nan, inplace=True)

        df.IO_Cabinet_Old = df.PLC.where(pd.isna(df.IO_Cabinet_Old), df.IO_Cabinet_Old)
        df.IO_Cabinet_New = df.PLC.where(pd.isna(df.IO_Cabinet_New), df.IO_Cabinet_New)
        df.IO_RACK_Old = df.RACK.where(pd.isna(df.IO_RACK_Old), df.IO_RACK_Old)
        df.IO_RACK_New = df.RACK.where(pd.isna(df.IO_RACK_New), df.IO_RACK_New)
        df.IO_POS_Old = df.POS.where(pd.isna(df.IO_POS_Old), df.IO_POS_Old)
        df.IO_POS_New = df.POS.where(pd.isna(df.IO_POS_New), df.IO_POS_New)
        df.IO_CHAN_Old = df.CHAN.where(pd.isna(df.IO_CHAN_Old), df.IO_CHAN_Old)
        df.IO_CHAN_New = df.CHAN.where(pd.isna(df.IO_CHAN_New), df.IO_CHAN_New)

        return df

    # This function reads the old PLC database.
    # - If `shuffleList` is provided, new PLC/FLD positions will be added to the database.
    # - If `IOList` is provided, new PLC/FLD positions will be added to the database.
    # - If `TagChangeList` is provided, new PLC/FLD positions will be added to the database.

    def readOldPLC(
        self, path, PLCpath, oldPLCs, shuffleList, tagChangeList, IOList, outputFile
    ):
        print("Reading old databases:")
        DB_old = pd.DataFrame()
        if tagChangeList != "":
            DB_TagChange = read_db(path, tagChangeList)
        else:
            DB_TagChange = pd.DataFrame(columns=["Old", "New"])

        if IOList != "":
            DB_IO = read_db(path, IOList)
            DB_IO = DB_IO[
                [
                    "TAGNUMBER",
                    "OLD Cabinet",
                    "NEW Cabinet",
                    "RACK Old",
                    "POS Old",
                    "CHAN Old",
                    "RACK New",
                    "POS New",
                    "CHAN New",
                ]
            ]
            DB_IO.columns = [
                "IO_TAGNUMBER",
                "IO_Cabinet_Old",
                "IO_Cabinet_New",
                "IO_RACK_Old",
                "IO_POS_Old",
                "IO_CHAN_Old",
                "IO_RACK_New",
                "IO_POS_New",
                "IO_CHAN_New",
            ]
        else:
            DB_IO = pd.DataFrame(
                columns=[
                    "IO_TAGNUMBER",
                    "IO_Cabinet_Old",
                    "IO_Cabinet_New",
                    "IO_RACK_Old",
                    "IO_POS_Old",
                    "IO_CHAN_Old",
                    "IO_RACK_New",
                    "IO_POS_New",
                    "IO_CHAN_New",
                ]
            )

        if shuffleList != "":
            DB_shuffle = read_db(path, shuffleList)
        else:
            DB_shuffle = pd.DataFrame(
                columns=["OLD_PLC", "OLD_FLD", "NEW_PLC", "NEW_FLD"]
            )

        first = True
        for curPLC in oldPLCs:
            print("- Reading ", curPLC[0])
            DB_old_temp = read_db(PLCpath, curPLC[0])
            DB_old_temp.insert(0, "PLC", curPLC[2])
            DB_old_temp = pd.merge(
                DB_old_temp,
                DB_shuffle,
                how="left",
                left_on=["PLC", "SHEET"],
                right_on=["OLD_PLC", "OLD_FLD"],
            )
            DB_old_temp = pd.merge(
                DB_old_temp,
                DB_IO,
                how="left",
                left_on=["PLC", "RACK", "POS", "CHAN"],
                right_on=["IO_Cabinet_Old", "IO_RACK_Old", "IO_POS_Old", "IO_CHAN_Old"],
            )

            DB_old_temp = self.createMatchTagOld(DB_old_temp, curPLC[1], DB_TagChange)
            if first:
                DB_old: pd.DataFrame = DB_old_temp
                first = False
            else:
                DB_old = pd.concat([DB_old, DB_old_temp])

            DB_old = self.fillShuffle(DB_old)
            DB_old = self.fillIO(DB_old)

            # DB_old = DB_old.drop_duplicates(subset ="MatchTag", keep = 'first')
        with pd.ExcelWriter(path + outputFile) as writer:
            print("- Writing old DB")
            DB_old.to_excel(writer, sheet_name="DB_old", index=False)
        with pd.ExcelWriter(path + outputFile, engine="openpyxl", mode="a") as writer:
            print("- Writing IO list")
            DB_IO.to_excel(writer, sheet_name="IO List", index=False)
            print("- Writing TagChange")
            DB_TagChange.to_excel(writer, sheet_name="TagChange", index=False)
            print("- Writing shuffle list")
            DB_shuffle.to_excel(writer, sheet_name="ShuffleList", index=False)
        print("Reading old database ready\n")
        return DB_old

    def readNewPLC(self, path, PLCpath, newPLCs, outputFile):
        print("Reading new databases:")
        first = True
        DB_new = pd.DataFrame()
        for curPLC in newPLCs:
            print("- Reading ", curPLC[0])
            DB_new_temp = read_db(PLCpath, curPLC[0])
            DB_new_temp = self.createMatchTagNew(DB_new_temp, curPLC[2])
            if first:
                DB_new = DB_new_temp
                first = False
            else:
                DB_new = pd.concat([DB_new, DB_new_temp])
        with pd.ExcelWriter(path + outputFile, engine="openpyxl", mode="a") as writer:
            print("- Writing new DB")
            DB_new.to_excel(writer, sheet_name="DB_new", index=False)
        print("Reading new databases ready\n")
        return DB_new

    def createCheckDB(self, path, DB_old, DB_new, outputFile):
        print("Creating check database:")
        print("- combining old/new inner join")
        DB_check = pd.merge(DB_old, DB_new, how="inner", on="MatchTag")

        print("- combining old/new outer join")
        DB_combined = pd.merge(DB_old, DB_new, how="outer", on="MatchTag")
        DB_combined["TagNumber"].replace("", np.nan, inplace=True)
        DB_combined["TAGNUMBER"].replace("", np.nan, inplace=True)

        print("- creating old remaining")
        DB_old_remaining = DB_combined[pd.isnull(DB_combined.TagNumber)]
        print("- creating new remaining")
        DB_new_remaining = DB_combined[pd.isnull(DB_combined.TAGNUMBER)]
        with pd.ExcelWriter(path + outputFile, engine="openpyxl", mode="a") as writer:
            print("- writing combined inner DB")
            DB_check.to_excel(writer, sheet_name="Combined inner", index=False)
            print("- writing combined outer DB")
            DB_combined.to_excel(writer, sheet_name="Combined outer", index=False)
            print("- writing remaining old DB")
            DB_old_remaining.to_excel(
                writer, sheet_name="DB_old_remaining", index=False
            )
            print("- writing remaining new DB")
            DB_new_remaining.to_excel(
                writer, sheet_name="DB_new_remaining", index=False
            )
        print("- preparing check database")
        DB_check.insert(loc=0, column="DUMP", value="DUMP")
        DB_check = self.moveColumn(DB_check, "MatchTag", "MatchTag")
        DB_check = self.moveColumn(DB_check, "PLC_x", "PLC")
        DB_check = self.moveColumn(DB_check, "TAGNUMBER", "TagNumber_old")
        DB_check = self.moveColumn(DB_check, "TagNumber", "TagNumber_new")
        DB_check = self.moveColumn(DB_check, "TYPE", "PointType_old")
        DB_check = self.moveColumn(DB_check, "PointType", "PointType_new")
        print("Done creating check database\n")

        # gui = show_df(DB_check)

        return DB_check

    # The following functions are all used to compare 1 type of databasefield. Details can be found per function.

    # This function checks the FLD field (sheet number). When a shuffle has taken place, the proposed NEW FLD (from the shufflelist) will be compared with the FLD number in the SM.
    # Differences are typically expected for functionblocks and references to functionblocks (FLD 2400+) and system alarms (FLD 60-).

    def checkFLD(self, dftotal, f_old="NEW_FLD", f_new="FLDNumber"):
        def func_FLD(row):
            if row["FLD_old"] == row["FLD_new"]:
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "FLD_old")
        dftotal = self.moveColumn(dftotal, f_new, "FLD_new")
        dftotal["FLD_check"] = dftotal.apply(lambda row: func_FLD(row), axis=1)
        dftotal = self.moveColumn(dftotal, "FLD_check", "FLD_check")

        return dftotal

    # The old and new descriptions are compared. Differences need to be reviewed manually. These differences are typically conversions of strange characters. Markers, Registers and Timers are ignored.

    def checkDescription(self, dftotal, f_old="SERVICE", f_new="Description"):
        def func_desc(row):
            if row["PointType_old"] in ["M", "R", "T"]:
                if row["Description_old"] == row["Description_new"]:
                    return "OK"
                else:
                    return "IGNORE M/R/T"

            if row["Description_old"] == row["Description_new"]:
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "Description_old")
        dftotal = self.moveColumn(dftotal, f_new, "Description_new")
        dftotal["Description_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "Description_check", "Description_check")

        return dftotal

    # This function compares the State1Text. Differences need to be reviewed manually.

    def checkState1Text(self, dftotal, f_old="QUALIFICAT", f_new="State1Text"):
        def func_desc(row):
            if row["State1Text_old"] == row["State1Text_new"]:
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "State1Text_old")
        dftotal = self.moveColumn(dftotal, f_new, "State1Text_new")
        dftotal["State1Text_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "State1Text_check", "State1Text_check")

        return dftotal

    # This function compares the old and new location. An example of an allowed conversion is included in the function, but this might be expanded with other comparable replacements.

    def checkLocation(self, dftotal, f_old="LOC", f_new="Location"):
        def func_desc(row):
            if row["Location_old"] == row["Location_new"]:
                return "OK"
            # example of allowed conversion
            if row["Location_old"] == "FSC" and row["Location_new"] == "COM":
                return "IGNORE FCS->DCS"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "Location_old")
        dftotal = self.moveColumn(dftotal, f_new, "Location_new")
        dftotal["Location_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "Location_check", "Location_check")

        return dftotal

    # This function compares the Unit. Differences need to be reviewed manually. Markers, Registers and Timers are ignored.

    def checkUnit(self, dftotal, f_old="UNIT", f_new="Unit"):
        def func_desc(row):
            if row["PointType_old"] in ["M", "R", "T"]:
                if row["Description_old"] == row["Description_new"]:
                    return "OK"
                else:
                    return "IGNORE M/R/T"
            if row["Unit_old"] == row["Unit_new"]:
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "Unit_old")
        dftotal = self.moveColumn(dftotal, f_new, "Unit_new")
        dftotal["Unit_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "Unit_check", "Unit_check")

        return dftotal

    # This function compares the SubUnit. Differences need to be reviewed manually. Markers, Registers and Timers are ignored.

    def checkSubUnit(self, dftotal, f_old="SUB_UNIT", f_new="SubUnit"):
        def func_desc(row):
            if row["PointType_old"] in ["M", "R", "T"]:
                if row["Description_old"] == row["Description_new"]:
                    return "OK"
                else:
                    return "IGNORE M/R/T"
            if row["SubUnit_old"] == row["SubUnit_new"]:
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "SubUnit_old")
        dftotal = self.moveColumn(dftotal, f_new, "SubUnit_new")
        dftotal["SubUnit_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "SubUnit_check", "SubUnit_check")

        return dftotal

    # This function compares the field SafetyRelated in both old and new situation.
    # - For Markers, Registers and Timers SafetyRelated cannot be filled in for SM, so these are ignored.
    # - FSC signals that have become COM signals will have SafetyRelated "No", and will therefore be ignored.

    def checkSafetyRelated(self, dftotal, f_old="SAFETY", f_new="SafetyRelated"):
        def func_desc(row):
            if row["PointType_new"] in [
                "M",
                "R",
                "T",
            ]:
                return "IGNORE"
            if (
                row["Location_old"] == "FSC"
                and row["Location_new"] == "COM"
                and row["SafetyRelated_new"] == "No"
            ):
                return "IGNORE"

            if row["Location_old"] == "":
                return "IGNORE"

            if row["SafetyRelated_old"][0] == row["SafetyRelated_new"][0]:
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "SafetyRelated_old")
        dftotal = self.moveColumn(dftotal, f_new, "SafetyRelated_new")
        dftotal["SafetyRelated_check"] = dftotal.apply(
            lambda row: func_desc(row), axis=1
        )
        dftotal = self.moveColumn(dftotal, "SafetyRelated_check", "SafetyRelated_check")

        return dftotal

    # This function compares the field ForceEnable in both old and new situation.
    # Here are 2 options:
    # 1. Compare against previous implementation and then compare differences with basic rules. (`NewRulesOnly=False`)
    # 2. Only compare against basic rules. (`NewRulesOnly=True`).
    #
    # Basic rules need to be copied from RT20-046 document
    #
    #

    def checkForceEnable(
        self, dftotal, NewRulesOnly=True, f_old="FORCE", f_new="ForceEnable"
    ):
        def func_desc(row):
            if not NewRulesOnly:
                if row["ForceEnable_old"] == "N" and not row["ForceEnable_new"]:
                    return "OK"
                if row["ForceEnable_old"] == "Y" and row["ForceEnable_new"]:
                    return "OK"

            # markers, timers and registers should not be forceble
            if row["PointType_new"] in ["M", "R", "T"]:
                if not row["ForceEnable_new"]:
                    return "OK"
                else:
                    return "INCORRECT M/R/T should be False"

            # field signals all force enabled, except for Flame Eyes (DI)
            if row["Location_new"] in [
                "FLD",
                "MCC",
                "PNL",
                "ADP",
                "LP",
                "CAB",
                "MDF",
                "SM",
                "PLC",
            ]:
                if row["ForceEnable_new"]:
                    return "OK"
                else:
                    return "INCORRECT FLD signals should be True (except Flame eyes)"

            # COM signals DI/BO force enabled, BI/DO not force enabled (except for DO systemtags)
            if row["Location_new"] == "COM":
                if row["PointType_new"] in ["DI", "BO"]:
                    if row["ForceEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT DI/BO COM should be True"
                if row["PointType_new"] in ["DO", "BI"]:
                    if not row["ForceEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT DO/BI COM should be False (except DO system tags)"

            # FSC signals DI/DO force enabled, BI/BO not force enabled
            if row["Location_new"] == "FSC":
                if row["PointType_new"] in ["DI", "DO"]:
                    if row["ForceEnable_new"]:
                        return "OK"
                    else:
                        return "DI/DO FSC should be True"
                if row["PointType_new"] in ["BI", "BO"]:
                    if not row["ForceEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT BI/BO FSC should be False"

            # ANN and SYS signals not force enabled
            if row["Location_new"] in ["ANN", "SYS"] and not row["ForceEnable_new"]:
                return "OK"
            else:
                return "INCORRECT ANN/SYS should be False"

        dftotal = self.moveColumn(dftotal, f_old, "ForceEnable_old")
        dftotal = self.moveColumn(dftotal, f_new, "ForceEnable_new")
        dftotal["ForceEnable_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "ForceEnable_check", "ForceEnable_check")

        return dftotal

    # This function compares the field WriteEnable in both old and new situation.
    # Here are 2 options:
    # 1. Compare against previous implementation and then compare differences with basic rules. (`NewRulesOnly=False`)
    # 2. Only compare against basic rules. (`NewRulesOnly=True`).
    #
    # Basic rules need to be copied from RT20-046 document

    def checkWriteEnable(
        self, dftotal, NewRulesOnly=True, f_old="WRITE", f_new="WriteEnable"
    ):
        def func_desc(row):
            if not NewRulesOnly:
                if row["WriteEnable_old"] == "N" and not row["WriteEnable_new"]:
                    return "OK"
                if row["WriteEnable_old"] == "Y" and row["WriteEnable_new"]:
                    return "OK"

            if row["Location_new"] == "COM" and row["PointType_new"] in ["DI", "BI"]:
                if row["WriteEnable_new"]:
                    return "OK"
                else:
                    return "INCORRECT, COM DI/BI should be TRUE"

            if not row["WriteEnable_new"]:
                return "OK"
            else:
                return "INCORRECT, Should be False for not COM DI/BI"

        dftotal = self.moveColumn(dftotal, f_old, "WriteEnable_old")
        dftotal = self.moveColumn(dftotal, f_new, "WriteEnable_new")
        dftotal["WriteEnable_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "WriteEnable_check", "WriteEnable_check")

        return dftotal

    # This function checks the difference between datatype of the old and new database.
    # - It is assumed that all AI, AO, BI and BO are converted to floats.
    # - Registers can be word, long or float, differences with old version need to be verified

    def checkDataType(self, dftotal, f_old="REGTYPE", f_new="DataType"):
        def func_desc(row):
            if row["PointType_new"] in ["DI", "DO", "M", "T"]:
                if row["DataType_new"] == "":
                    return "OK"
                else:
                    return "INCORRECT, Should be empty"

            if row["PointType_new"] in ["AI", "AO", "BI", "BO"]:
                if row["DataType_new"] == "Float":
                    return "OK"
                else:
                    return "INCORRECT, Float expected"

            # only registers are remaining now (can be word, long or float)
            if row["PointType_old"] == "":
                return "OK"
            if row["DataType_old"] == "":
                if row["DataType_new"] == "":
                    return "OK"
                else:
                    return "DIFFERENT"
            if row["DataType_new"] == "":
                if row["DataType_old"] == "":
                    return "OK"
                else:
                    return "DIFFERENT"

            if row["DataType_old"][0] == row["DataType_new"][0]:
                return "OK"
            else:
                return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "DataType_old")
        dftotal = self.moveColumn(dftotal, f_new, "DataType_new")
        dftotal["DataType_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "DataType_check", "DataType_check")

        return dftotal

    # This function checks the signal type. This simply is 4-20mA or empty. In the old version there is a space between '4-20' and 'mA'.

    def checkSignalType(self, dftotal, f_old="SIGNALTYPE", f_new="SignalType"):
        def func_desc(row):
            if row["SignalType_old"].replace(" ", "") == row["SignalType_new"].replace(
                " ", ""
            ):
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "SignalType_old")
        dftotal = self.moveColumn(dftotal, f_new, "SignalType_new")
        dftotal["SignalType_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "SignalType_check", "SignalType_check")

        return dftotal

    # This function checks the bottom scale. This only applies to AI. In the case of conversion from raw to EU, top and bottom scale will have changed and need to be checked manually.

    def checkBottomScale(self, dftotal, f_old="ANBOTTOM", f_new="BottomScale"):
        def func_desc(row):
            if row["PointType_new"] == "AI":
                if round(float(row["BottomScale_old"]), 8) == round(
                    float(row["BottomScale_new"]), 8
                ):
                    return "OK"
                else:
                    return "DIFFERENT"
            return "IGNORE"

        dftotal = self.moveColumn(dftotal, f_old, "BottomScale_old")
        dftotal = self.moveColumn(dftotal, f_new, "BottomScale_new")
        dftotal["BottomScale_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "BottomScale_check", "BottomScale_check")

        return dftotal

    # This function checks the bottom top. This only applies to AI. In the case of conversion from raw to EU, top and bottom scale will have changed and need to be checked manually.

    def checkTopScale(self, dftotal, f_old="ANTOP", f_new="TopScale"):
        def func_desc(row):
            if row["PointType_new"] == "AI":
                if round(float(row["TopScale_old"]), 8) == round(
                    float(row["TopScale_new"]), 8
                ):
                    return "OK"
                else:
                    return "DIFFERENT"
            return "IGNORE"

        dftotal = self.moveColumn(dftotal, f_old, "TopScale_old")
        dftotal = self.moveColumn(dftotal, f_new, "TopScale_new")
        dftotal["TopScale_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "TopScale_check", "TopScale_check")

        return dftotal

    # This function checks the engineering units.
    #
    # Differences need to be checked manually, but may be changes from degC to °C.
    # In the case of conversion from raw to float, the old situation had no engineering unit at all, but the new will have.

    def checkEngineeringUnits(
        self, dftotal, f_old="AENGUNIT", f_new="EngineeringUnits"
    ):
        def func_desc(row):
            if row["EngineeringUnits_old"] == row["EngineeringUnits_new"]:
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "EngineeringUnits_old")
        dftotal = self.moveColumn(dftotal, f_new, "EngineeringUnits_new")
        dftotal["EngineeringUnits_check"] = dftotal.apply(
            lambda row: func_desc(row), axis=1
        )
        dftotal = self.moveColumn(
            dftotal, "EngineeringUnits_check", "EngineeringUnits_check"
        )

        return dftotal

    #
    # This function compares the field SOEEnable in both old and new situation.
    # Here are 2 options:
    # 1. Compare against previous implementation and then compare differences with basic rules. (`NewRulesOnly=False`)
    # 2. Only compare against basic rules. (`NewRulesOnly=True`).
    #
    # Basic rules need to be copied from RT20-046 document

    def checkSOEEnable(
        self, dftotal, NewRulesOnly=True, f_old="SER", f_new="SOEEnable"
    ):
        def func_desc(row):
            if not NewRulesOnly:
                if row["SOEEnable_old"] == "N" and not row["SOEEnable_new"]:
                    return "OK"
                if row["SOEEnable_old"] == "Y" and row["SOEEnable_new"]:
                    return "OK"
                if row["SOEEnable_old"] == row["SOEEnable_new"]:
                    return "OK"

            if row["PointType_new"] in ["AI", "AO", "BI", "BO"]:
                if not row["SOEEnable_new"]:
                    return "OK"
                else:
                    return "INCORRECT, FALSE expected"

            if row["Location_new"] in [
                "FLD",
                "MCC",
                "PNL",
                "ADP",
                "LP",
                "CAB",
                "MDF",
                "SM",
                "PLC",
                "COM",
                "FSC",
                "ANN",
            ]:
                if row["PointType_new"] == "DI":
                    if row["SOEEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT, TRUE expected for FLD/ANN/COM/FSC DI"

            if row["Location_new"] in [
                "FLD",
                "MCC",
                "PNL",
                "ADP",
                "LP",
                "CAB",
                "MDF",
                "SM",
                "PLC",
            ]:
                if row["PointType_new"] == "DO":
                    if row["SOEEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT, TRUE expected for FLD DO, except for LEDs"

            if row["Location_new"] == "COM":
                if row["PointType_new"] == "DO":
                    if not row["SOEEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT, FALSE expected, only TRUE for COM DO with no hardware COM"

            if row["Location_new"] == "FSC":
                if row["PointType_new"] == "DO":
                    if not row["SOEEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT, FALSE expected for FSC DO"

            if row["Location_new"] == "ANN":
                if row["PointType_new"] == "DO":
                    if row["SOEEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT, TRUE expected for ANN DO"

            if row["Location_new"] == "SYS":
                if row["PointType_new"] == "DI":
                    if not row["SOEEnable_new"]:
                        return "OK"
                    else:
                        return "INCORRECT, FALSE expected for SYS DI, except for SOE buffer and controller fault"

            return "Check manually"

        dftotal = self.moveColumn(dftotal, f_old, "SOEEnable_old")
        dftotal = self.moveColumn(dftotal, f_new, "SOEEnable_new")
        dftotal["SOEEnable_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "SOEEnable_check", "SOEEnable_check")

        return dftotal

    # This function checks the transmitter low alarm (in mA). Normally this value should not change, so differences need to be checked manually. Non-AI point can be ignored.

    def checkTransmitterAlarmLow(
        self, dftotal, f_old="TRMALSETPL", f_new="TransmitterAlarmLow"
    ):
        def func_desc(row):
            if row["PointType_new"] != "AI":
                return "IGNORE"
            if (
                row["TransmitterAlarmLow_old"] == -1
                and row["TransmitterAlarmLow_new"] == 0
            ):
                return "OK"
            if round(row["TransmitterAlarmLow_old"], 2) == round(
                row["TransmitterAlarmLow_new"], 2
            ):
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "TransmitterAlarmLow_old")
        dftotal = self.moveColumn(dftotal, f_new, "TransmitterAlarmLow_new")
        dftotal["TransmitterAlarmLow_check"] = dftotal.apply(
            lambda row: func_desc(row), axis=1
        )
        dftotal = self.moveColumn(
            dftotal, "TransmitterAlarmLow_check", "TransmitterAlarmLow_check"
        )

        return dftotal

    # This function checks the transmitter high alarm (in mA). Normally this value should not change, so differences need to be checked manually. Non-AI point can be ignored.

    def checkTransmitterAlarmHigh(
        self, dftotal, f_old="TRMALSETPH", f_new="TransmitterAlarmHigh"
    ):
        def func_desc(row):
            if row["PointType_new"] != "AI":
                return "IGNORE"
            if (
                row["TransmitterAlarmHigh_old"] == -1
                and row["TransmitterAlarmHigh_new"] == 25
            ):
                return "OK"
            if round(row["TransmitterAlarmHigh_old"], 4) == round(
                row["TransmitterAlarmHigh_new"], 4
            ):
                return "OK"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "TransmitterAlarmHigh_old")
        dftotal = self.moveColumn(dftotal, f_new, "TransmitterAlarmHigh_new")
        dftotal["TransmitterAlarmHigh_check"] = dftotal.apply(
            lambda row: func_desc(row), axis=1
        )
        dftotal = self.moveColumn(
            dftotal, "TransmitterAlarmHigh_check", "TransmitterAlarmHigh_check"
        )

        return dftotal

    #
    # This function checks the fault reaction.
    #
    # Note that in FSC some values did not need to be filled in, while in SM they should.
    #
    # * hardware:
    #     * AO should be fail-safe ('0mA').
    #     * AI should be towards trip direction - needs to be checked manually, especially when changed.
    #     * DO (hardware) should be fail-safe ('Low')
    #     * DI should be fail-safe ('Low').
    # * COM:
    #     * BI and DI set to 'Freeze'.
    #     * BO does not have fault reaction ('')
    #     * DO does not have fault reaction ('Undefined')
    # * FSC:
    #     * BO does not have fault reaction ('')
    #     * DO does not have fault reaction ('Undefined')
    #     * BI/DI: Application dependend - checked versus FSC application
    # * ANN:
    #     * should be 'Undefined'
    # * Markers/Registers/Timers do not have a fault reaction ('')

    def checkFaultReaction(
        self,
        dftotal,
        NewRulesOnly=True,
        f_old="FAULTREACT",
        f_new="FaultReaction",
    ):
        def func_desc(row):
            if NewRulesOnly:
                if row["Location_new"] in [
                    "FLD",
                    "MCC",
                    "PNL",
                    "ADP",
                    "LP",
                    "CAB",
                    "MDF",
                    "SM",
                    "PLC",
                ]:
                    if row["PointType_new"] == "DO":
                        if row["FaultReaction_new"] == "Low":
                            return "OK"
                        else:
                            return "'INCORRECT, Low' expected for FLD DO"
                    if row["PointType_new"] == "AO":
                        if row["FaultReaction_new"] == "0 mA":
                            return "OK"
                        else:
                            return "INCORRECT, '0 mA' expected for FLD AO"
                    if row["PointType_new"] == "DI":
                        if row["FaultReaction_new"] == "Low":
                            return "OK"
                        else:
                            return "INCORRECT, 'low' expected for FLD DI"
                    if row["PointType_new"] == "AI":
                        if row["FaultReaction_old"] == row["FaultReaction_new"]:
                            return "OK"
                        return "DIFFERENT"

                if row["Location_new"] == "COM":
                    if row["PointType_new"] == "DO":
                        if row["FaultReaction_new"] == "Undefined":
                            return "OK"
                        else:
                            return "INCORRECT, 'Undefined' expected for COM DO"
                    if row["PointType_new"] == "BO":
                        if row["FaultReaction_new"] == "":
                            return "OK"
                        else:
                            return "INCORRECT, empty expected for COM BO"
                    if row["PointType_new"] == "DI":
                        if row["FaultReaction_new"] == "Freeze":
                            return "OK"
                        else:
                            return "INCORRECT, 'Freeze' expected for COM DI"
                    if row["PointType_new"] == "BI":
                        if row["FaultReaction_new"] == "Freeze":
                            return "OK"
                        else:
                            return "INCORRECT, 'Freeze' expected for COM BI"

                if row["Location_new"] == "FSC":
                    if row["PointType_new"] == "DO":
                        if row["FaultReaction_new"] == "Undefined":
                            return "OK"
                        else:
                            return "INCORRECT, 'Undefined' expected for FSC DO"
                    if row["PointType_new"] == "BO":
                        if row["FaultReaction_new"] == "":
                            return "OK"
                        else:
                            return "INCORRECT, empty expected for FSC BO"
                    if row["PointType_new"] in ["DI", "BI"]:
                        if row["FaultReaction_old"] == row["FaultReaction_new"]:
                            return "OK"
                        return "DIFFERENT"

                if row["Location_new"] == "ANN":
                    if row["FaultReaction_new"] == "Undefined":
                        return "OK"
                    else:
                        return "INCORRECT, 'Undefined' expected for ANN"

                if row["Location_new"] == "":
                    if row["FaultReaction_new"] == "":
                        return "OK"
                    else:
                        return "INCORRECT, empty expected for R/T/M"

                if row["FaultReaction_old"] == row["FaultReaction_new"]:
                    return "OK"
                return "DIFFERENT"
            else:
                if row["FaultReaction_old"] == row["FaultReaction_new"]:
                    return "OK"
                if row["FaultReaction_old"] == "N.a.":
                    return "IGNORE"
                else:
                    return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "FaultReaction_old")
        dftotal = self.moveColumn(dftotal, f_new, "FaultReaction_new")
        dftotal["FaultReaction_check"] = dftotal.apply(
            lambda row: func_desc(row), axis=1
        )
        dftotal = self.moveColumn(dftotal, "FaultReaction_check", "FaultReaction_check")

        return dftotal

    def checkRack(self, dftotal, f_old="IO_RACK_New", f_new="ChassisID/IOTAName"):
        def func_desc(row):
            if row["PointType_new"] not in ["AI", "DI", "DO", "AO"]:
                return "IGNORE"
            if row["Location_new"] in ["SYS", "FSC", "COM", "ANN"]:
                return "IGNORE"
            # don't have example of RUSIO in FSC, so just ignoring RUSIO in SM for now
            if "RUSIO" in row["Rack_new"]:
                if "RUSIO" in row["Rack_old"]:
                    return "OK"
                else:
                    return "DIFFERENT (RUSIO)"
            try:
                if int(str(row["Rack_old"])) == int(str(row["Rack_new"])[3:]):
                    return "OK"
            except:
                print(row["Rack_new"], row["MatchTag"])
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "Rack_old")
        dftotal = self.moveColumn(dftotal, f_new, "Rack_new")
        dftotal["Rack_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "Rack_check", "Rack_check")

        return dftotal

    def checkSlotNumber(self, dftotal, f_old="IO_POS_New", f_new="SlotNumber"):
        def func_desc(row):
            if row["PointType_new"] not in ["AI", "DI", "DO", "AO"]:
                return "IGNORE"
            if row["Location_new"] in ["SYS", "FSC", "COM", "ANN"]:
                return "IGNORE"
            try:
                if int(str(row["SlotNumber_old"])) == int(str(row["SlotNumber_new"])):
                    return "OK"
            except:
                return "DIFFERENT"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "SlotNumber_old")
        dftotal = self.moveColumn(dftotal, f_new, "SlotNumber_new")
        dftotal["SlotNumber_check"] = dftotal.apply(lambda row: func_desc(row), axis=1)
        dftotal = self.moveColumn(dftotal, "SlotNumber_check", "SlotNumber_check")

        return dftotal

    def checkChannelNumber(self, dftotal, f_old="IO_CHAN_New", f_new="ChannelNumber"):
        def func_desc(row):
            if row["PointType_new"] not in ["AI", "DI", "DO", "AO"]:
                return "IGNORE"
            if row["Location_new"] in ["SYS", "FSC", "COM", "ANN"]:
                return "IGNORE"
            try:
                if int(str(row["ChannelNumber_old"])) == int(
                    str(row["ChannelNumber_new"])
                ):
                    return "OK"
            except:
                return "DIFFERENT"
            return "DIFFERENT"

        dftotal = self.moveColumn(dftotal, f_old, "ChannelNumber_old")
        dftotal = self.moveColumn(dftotal, f_new, "ChannelNumber_new")
        dftotal["ChannelNumber_check"] = dftotal.apply(
            lambda row: func_desc(row), axis=1
        )
        dftotal = self.moveColumn(dftotal, "ChannelNumber_check", "ChannelNumber_check")

        return dftotal

    # This function executes all checks and writes the result to the outputfile.

    def doChecks(self, path, DB_check, outputFile, phase):
        print("Executing checks:")
        print("- checking FLD")
        DB_check = self.checkFLD(DB_check)
        print("- checking Description")
        DB_check = self.checkDescription(DB_check)
        print("- checking State1Text")
        DB_check = self.checkState1Text(DB_check)
        print("- checking Location")
        DB_check = self.checkLocation(DB_check)
        print("- checking Unit")
        DB_check = self.checkUnit(DB_check)
        print("- checking SubUnit")
        DB_check = self.checkSubUnit(DB_check)
        print("- checking SafetyRelated")
        DB_check = self.checkSafetyRelated(DB_check)
        print("- checking ForceEnabled")
        DB_check = self.checkForceEnable(DB_check, phase != "TUV")
        print("- checking WriteEnabled")
        DB_check = self.checkWriteEnable(DB_check, phase != "TUV")
        print("- checking DataType")
        DB_check = self.checkDataType(DB_check)
        print("- checking SignalType")
        DB_check = self.checkSignalType(DB_check)
        print("- checking BottomScale")
        DB_check = self.checkBottomScale(DB_check)
        print("- checking TopScale")
        DB_check = self.checkTopScale(DB_check)
        print("- checking EngineeringUnits")
        DB_check = self.checkEngineeringUnits(DB_check)
        print("- checking SOEEnable")
        DB_check = self.checkSOEEnable(DB_check, phase != "TUV")
        print("- checking TransmitterAlarmLow")
        DB_check = self.checkTransmitterAlarmLow(DB_check)
        print("- checking TransmitterAlarmHigh")
        DB_check = self.checkTransmitterAlarmHigh(DB_check)
        print("- checking FaultReaction")
        DB_check = self.checkFaultReaction(DB_check, phase != "TUV")
        print("- checking Rack")
        DB_check = self.checkRack(DB_check)
        print("- checking Slot")
        DB_check = self.checkSlotNumber(DB_check)
        print("= checking Channel")
        DB_check = self.checkChannelNumber(DB_check)

        print("- dropping remaining columns")
        cols = DB_check.columns.tolist()
        DB_check = DB_check.drop(cols[cols.index("DUMP") :], axis=1)

        with pd.ExcelWriter(path + outputFile, engine="openpyxl", mode="a") as writer:
            print("Writing check DB\n")
            DB_check.to_excel(writer, sheet_name="Check", index=False)
        return DB_check

    def check_PLC(
        self,
        my_path,
        phase,
        outputFile,
        oldPLCs,
        newPLCs,
        shuffleList,
        tagChangeList,
        IOList,
        colouring=True,
    ):
        DB_old = self.readOldPLC(
            my_path,
            my_path + "PLCs\\Original\\",
            oldPLCs,
            shuffleList,
            tagChangeList,
            IOList,
            outputFile,
        )
        DB_new = self.readNewPLC(
            my_path, my_path + rf"PLCs\\{phase}\\", newPLCs, outputFile
        )
        DB_check = self.createCheckDB(my_path, DB_old, DB_new, outputFile)
        DB_check = self.doChecks(my_path, DB_check, outputFile, phase)
        if colouring:
            print("Formatting Excel...")
            format_excel(
                path=my_path,
                filename=outputFile,
                first_time=True,
                different_red=True,
                different_blue=False,
                check_existing_red=False,
            )

        print(f"{Fore.YELLOW}Done!{Fore.RESET}")

    # TODO Moved COM signals
    # TODO IO check when moved IOs

    def start(self, project: str, phase: str = "TUV"):
        my_proj = ProjDetails(project)
        my_path = my_proj.path
        outputFile = my_proj.outputFile[phase]
        oldPLCs = my_proj.PLC_list["Original"]
        print(oldPLCs)
        newPLCs = my_proj.PLC_list[phase]
        print(newPLCs)

        shuffleList = my_proj.shuffleList  # ignored if "" - placed in path+"Extra\\"
        # ignored if "" - placed in path+"Extra\\"
        tagChangeList = my_proj.tagChangeList
        IOList = my_proj.ioList  # ignored if "" - placed in path+"Extra\\"

        dfcheck = self.check_PLC(
            my_path=my_path,
            phase=phase,
            outputFile=outputFile,
            oldPLCs=oldPLCs,
            newPLCs=newPLCs,
            shuffleList=shuffleList,
            tagChangeList=tagChangeList,
            IOList=IOList,
            colouring=True,
        )


def main():
    project = CheckFSC2SM("CVP_MOD9", "FAT")


if __name__ == "__main__":
    main()
