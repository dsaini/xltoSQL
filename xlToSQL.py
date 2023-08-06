import numpy as np
import pandas as pd
from tkinter.filedialog import askopenfilename
from os import startfile
import ctypes


def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def get_sheetnames(filename):
    dataframe1 = pd.read_excel(filename, None)
    table_name_list = list(dataframe1.keys())
    dataframe1 = {}
    return table_name_list


def get_xlsheet_into_dataframe(filename, n):
    dataframe1 = pd.read_excel(filename, sheet_name=n)
    dataframe1 = dataframe1.replace(np.NaN, 'NULL').applymap(str)
    return dataframe1


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    print("Please select an .xls or .xlsx file...")

    fn = askopenfilename()

    try:
        if fn:
            print("Opening file...")
            list_of_tables = get_sheetnames(fn)
            print("Table names found: " + ' '.join(list_of_tables))
            for i in range(len(list_of_tables)):
                df0 = get_xlsheet_into_dataframe(fn, i)
                list_col_names = list(df0.keys())

                insert_statement = "INSERT INTO " + list_of_tables[i] + " (" + ",".join(list_col_names) + ") VALUES \n"

                list_of_rows = []
                for j, row in df0.iterrows():
                    lr = row.tolist()
                    val = ("('" + "','".join(lr) + "')").replace("'NULL'", "NULL").replace(".0", "")
                    list_of_rows.append(val)

                insert_statement = insert_statement + ",\n".join(list_of_rows)
                f = open(list_of_tables[i] + ".txt", "w")
                # print(insert_statement)
                f.write(insert_statement)
                f.close()
                startfile(list_of_tables[i] + ".txt")
                print("Success! See text file that was opened just now.")
    except Exception as err:
        print("Error:", err)
        Mbox("Error", "Cannot parse file", 0)

