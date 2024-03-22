# import libraries
import sys
import pandas as pd
import numpy as np
from simpledbf import Dbf5
from routines import format_excel, read_db


class XLSCompare:
    def __init__(self, path, output_file, file_old, unique_old, file_new, unique_new):
        self.start(path, output_file, file_old, unique_old, file_new, unique_new)

    def addComparingKey(self, DB, unique, suffix):
        DB["Comparing_Key"] = ""
        for item in unique:
            DB["Comparing_Key"] = (
                DB["Comparing_Key"].astype(str) + "___" + DB[item + suffix].astype(str)
            )
        if DB["Comparing_Key"].is_unique:
            print("- comparing key added")
            return DB
        else:
            sys.exit("Comparing key is not unique, aborting...")

    def checkColumnNames(self, DB_old_columns, DB_new_columns):
        col_both = []
        col_old = []
        col_new = []
        for col in DB_old_columns:
            if col in DB_new_columns:
                col_both.append(col)
            else:
                col_old.append(col)
        for col in DB_new_columns:
            if col not in DB_old_columns:
                col_new.append(col)
        return col_both, col_old, col_new

    def addSuffix(self, df, suffix):
        cols = df.columns.tolist()
        cols = ["{}_{}".format(a, b) for b in [suffix] for a in cols]
        df.columns = cols
        return df, cols

    def moveColumn(self, df, column_old, column_new):
        cols = df.columns.tolist()
        cols.insert(cols.index("DUMP"), cols.pop(cols.index(column_old)))
        df = df[cols]
        cols[cols.index(column_old)] = column_new
        df.columns = cols
        return df

    def createCheckDB(
        self,
        path,
        DB_old,
        DB_new,
        DB_old_key,
        DB_new_key,
        cols_old,
        cols_new,
        output_file,
    ):
        print("Creating empty check database:")
        print("- combining old/new inner join")
        DB_check = pd.merge(DB_old, DB_new, how="inner", on="Comparing_Key")

        print("- combining old/new outer join")
        DB_combined = pd.merge(DB_old, DB_new, how="outer", on="Comparing_Key")

        DB_combined[DB_old_key + "_old"].replace("", np.nan, inplace=True)
        DB_combined[DB_new_key + "_new"].replace("", np.nan, inplace=True)

        print("- creating old remaining")
        DB_old_remaining = DB_combined[pd.isnull(DB_combined[DB_new_key + "_new"])]

        print("- creating new remaining")
        DB_new_remaining = DB_combined[pd.isnull(DB_combined[DB_old_key + "_old"])]
        with pd.ExcelWriter(path + output_file, engine="openpyxl", mode="a") as writer:
            print("- writing combined inner DB")
            DB_check.to_excel(writer, sheet_name="Combined inner", index=False)
            print("- writing combined outer DB")
            DB_combined.to_excel(writer, sheet_name="Combined outer", index=False)
            print("- writing remaining old DB")
            DB_old_remaining[cols_old].to_excel(
                writer, sheet_name="DB_old_remaining", index=False
            )
            print("- writing remaining new DB")
            DB_new_remaining[cols_new].to_excel(
                writer, sheet_name="DB_new_remaining", index=False
            )
        print("- preparing check database")
        DB_check.insert(loc=0, column="DUMP", value="DUMP")
        DB_check = self.moveColumn(DB_check, "Comparing_Key", "Comparing_Key")
        print("Done creating check database\n")
        return DB_check

    def compareColumns(self, DB_compare, col_both, output_file, path):
        def func_desc(row):
            if row[col + "_old"] == row[col + "_new"]:
                return False
            else:
                if col == "SHORTCODE":
                    return False
                return "DIFFERENT"

        print("Doing column checks:")
        for col in col_both:
            print("- checking " + col)
            DB_compare = self.moveColumn(DB_compare, col + "_old", col + "_old")
            DB_compare = self.moveColumn(DB_compare, col + "_new", col + "_new")
            DB_compare[col + "_different"] = DB_compare.apply(
                lambda row: func_desc(row), axis=1
            )
            DB_compare = self.moveColumn(
                DB_compare, col + "_different", col + "_different"
            )

        col_check = ["{}_{}".format(a, b) for b in ["different"] for a in col_both]

        DB_compare["Diff"] = DB_compare[col_check].any(axis=1)
        cols = DB_compare.columns.tolist()
        cols.insert(0, cols.pop(cols.index("Diff")))
        DB_compare = DB_compare[cols]

        print("- writing check  DB")
        with pd.ExcelWriter(path + output_file, engine="openpyxl", mode="a") as writer:
            DB_compare.to_excel(writer, sheet_name="Difference", index=False)

        print("Done doing column checks.")

        return DB_compare

    def ExcelCompare(
        self, path, output_file, file_old, unique_old, file_new, unique_new
    ):
        print("Reading old database")
        DB_old = read_db(path, file_old)
        assert isinstance(DB_old, pd.DataFrame)
        DB_old_columns = DB_old.columns

        print("- Writing old database")
        with pd.ExcelWriter(path + output_file) as writer:
            DB_old.to_excel(writer, sheet_name="DB_old", index=False)

        DB_old, cols_old = self.addSuffix(DB_old, "old")
        DB_old = self.addComparingKey(DB_old, unique_old, "_old")

        print("Reading new database")
        DB_new = read_db(path, file_new)
        assert isinstance(DB_new, pd.DataFrame)
        DB_new_columns = DB_new.columns

        print("- Writing new database")
        with pd.ExcelWriter(path + output_file, engine="openpyxl", mode="a") as writer:
            DB_new.to_excel(writer, sheet_name="DB_new", index=False)

        DB_new, cols_new = self.addSuffix(DB_new, "new")
        DB_new = self.addComparingKey(DB_new, unique_new, "_new")

        col_both, col_old, col_new = self.checkColumnNames(
            DB_old_columns, DB_new_columns
        )

        DB_check = self.createCheckDB(
            path,
            DB_old,
            DB_new,
            unique_old[0],
            unique_new[0],
            cols_old,
            cols_new,
            output_file,
        )
        DB_check = self.compareColumns(DB_check, col_both[:-1], output_file, path)

        print("Formatting file")
        format_excel(
            path,
            output_file,
            first_time=True,
            different_red=True,
            different_blue=False,
            check_existing_red=False,
        )
        return DB_check

    def start(self, path, output_file, file_old, unique_old, file_new, unique_new):
        # path = "Test\\"
        # output_file = "results_DC.xlsx"
        # file_old = "DuringFAT DC.xlsx"
        # unique_old = ["&N"]
        # file_new = "AfterFAt DC.xlsx"
        # unique_new = ["&N"]
        self.ExcelCompare(path, output_file, file_old, unique_old, file_new, unique_new)


def main():
    # Requires file old, file new and a unique comparison column
    path = "Test\\"
    output_file = "Differences.xlsx"
    file_old = "RVC_HV8 (Original).xls"
    unique_old = ["PointType", "TagNumber"]
    file_new = "RVC_HV8.xls"
    unique_new = ["PointType", "TagNumber"]
    project = XLSCompare(path, output_file, file_old, unique_old, file_new, unique_new)


if __name__ == "__main__":
    main()
