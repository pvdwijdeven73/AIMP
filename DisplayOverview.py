import pandas as pd
from pandasgui import show as show_df
from os import system
from colorama import Fore
from routines import format_excel


class DispOverview:
    def __init__(self, project: str, phase="Original", perPLC: bool = False):
        self.start(project, phase, perPLC)

    def write_Overview(self, file_HMI, file_output, project, perPLC):
        def insert_tagname(row):
            for column in row.index:
                # print(column)
                if row[column] == 1:
                    row[column] = row.name
            return row

        # df_combined = pd.read_excel(file_CB, sheet_name="Combined")
        df_HMI = pd.read_excel(file_HMI, sheet_name="Overview")
        df_HMI["Count"] = 1
        if perPLC:
            PLCs = sorted(df_HMI["PLC"].unique())
            print(PLCs)
            with pd.ExcelWriter(
                f"Projects\\{project}\\{file_output}", engine="openpyxl", mode="w"
            ) as writer:
                for PLC in PLCs:
                    df_piv = pd.pivot_table(
                        df_HMI[df_HMI["PLC"] == PLC],
                        values="Count",
                        index=["Tagname"],
                        columns=["Display"],
                        fill_value="",
                    )
                    df_piv.fillna("")
                    df_result = df_piv.apply(lambda row: insert_tagname(row), axis=1)
                    print(f"- Writing DCS Export")
                    df_result.to_excel(
                        writer, sheet_name=f"Display_Overview_{PLC}", index=True
                    )
            print("Formatting Excel")
            format_excel(f"Projects\\{project}\\", file_output)
        else:
            df_piv = pd.pivot_table(
                df_HMI,
                values="Count",
                index=["Tagname"],
                columns=["Display"],
                fill_value="",
            )
            df_piv.fillna("")
            df_result = df_piv.apply(lambda row: insert_tagname(row), axis=1)
            with pd.ExcelWriter(
                f"Projects\\{project}\\{file_output}", engine="openpyxl", mode="w"
            ) as writer:
                print(f"- Writing DCS Export")
                df_result.to_excel(writer, sheet_name="Display_Overview", index=True)
            print("Formatting Excel")
            format_excel(f"Projects\\{project}\\", file_output)
        # show_df(df_result)
        return

    def start(self, project, phase, perPLC=False):

        print(
            f"{Fore.MAGENTA}Creating display overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )

        if project == "Test":
            Export_input_file = "Test\\HMIExport.xlsx"
            Export_output_file = "TestOverview.xlsx"
        else:
            Export_input_file = (
                f"Projects\\{project}\\{project}_Display_Params_{phase}.xlsx"
            )
            Export_output_file = f"{project}_Display_Overview_{phase}.xlsx"

        # self.get_params_list(EB_path)
        self.write_Overview(Export_input_file, Export_output_file, project, perPLC)
        print(
            f"{Fore.MAGENTA}Finished creating display overview for {Fore.GREEN}{project}{Fore.MAGENTA}{Fore.RESET}"
        )


def main():
    system("cls")
    # Requires: Excel file of display params for this phase: Projects\\{project}\\{project}_Display_Params_{phase}.xlsx
    # Requires: if per PLC, an extra column PLC needs to be added to the Overview sheet.
    project = DispOverview("PSU30_35_45", "FAT", True)


if __name__ == "__main__":
    main()
