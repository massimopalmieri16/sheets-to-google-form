from openpyxl.utils.exceptions import InvalidFileException
from openpyxl import load_workbook
from zipfile import BadZipFile
import tkinter as tk
from tkinter import filedialog
import itertools
import requests

root = tk.Tk()
root.withdraw()

ENTRY_ROW = 2
# Put here google form id
FORM_ID = ""


# every field of the form can be filled adding params using this schema
# entry.XXXXXXXX=value
# example: entry.123456=test_value
#
# send a GET request with this URL to send the form
url = f"https://docs.google.com/forms/d/e/{FORM_ID}/formResponse?entry.1=Test&entry.2=Test&entry.3=100"


def main():

    sheets_file_name = filedialog.askopenfilename()
    if not sheets_file_name:
        print("No file selected")
        return

    # load xlsx
    try:
        print("Opening " + sheets_file_name + "...")
        wb = load_workbook(filename = sheets_file_name, read_only = True)
        print("File succesfully opened")

        try:
            ws = wb.worksheets[0]
        except IndexError:
            print("Sheet not found")
            return

        # read entry row
        entry = []
        entry_row = next(itertools.islice(ws.rows, ENTRY_ROW - 1, None))
        for cell in entry_row:
            entry.append(int(cell.value))

        # read every row
        values_rows = []
        value_rows = itertools.islice(ws.rows, ENTRY_ROW, len(list(ws.rows)))
        for row in value_rows:
            values_row = []
            for cell in row:
                values_row.append(cell.value)
            
            values_rows.append(values_row)

        print(entry)
        print(values_rows)

        for values_row in values_rows:
            str_values = ""
            for i in range(0, len(values_row)):
                str_values += "entry.{entry}={value}&".format(entry=entry[i], value=values_row[i])

            url = "https://docs.google.com/forms/d/e/{form_id}/formResponse?{param}".format(form_id = FORM_ID, param = str_values)
            response = requests.get(url)
            print(response.status_code)

    except BadZipFile:
        print("File not valid")
        return
    except InvalidFileException:
        print("File format not valid")
        return


if __name__ == "__main__":
    main()