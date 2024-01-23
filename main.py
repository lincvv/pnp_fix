import PySimpleGUI as sg
from pathlib import Path
import openpyxl
import csv
import os

NEGATIVE_SIGN = '-'
ROTATION_270 = 270
ROTATION_90 = 90
ROTATION_180 = 180
ROTATION_0 = 0
CONVERTED_FACTOR = 0.0254


def csv_to_excel(csv_file, excel_file):
    csv_data = []
    with open(csv_file) as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)
    wb = openpyxl.Workbook()
    sheet = wb.active
    for row in csv_data:
        sheet.append(row)
    wb.save(excel_file)


def excel_to_csv(csv_file, excel_file):
    print("to csv")
    wb = openpyxl.load_workbook(filename=excel_file)
    ws = wb.active
    with open(csv_file, 'w', newline="") as file:
        writer = csv.writer(file)
        for line in ws.iter_rows():
            writer.writerow([cell.value for cell in line])


def read_col(sh, col_name):
    list_data_col = list()
    for col_cell in sh.iter_cols(1, sh.max_column):
        if col_cell[0].value == col_name:
            for cell_data in col_cell[1:]:
                list_data_col.append(cell_data.value)
    print(list_data_col)
    return list_data_col


def change_data(sh, col_name, list_val, prefix=None):
    for col_cell in sh.iter_cols(1, sh.max_column):  # iterate column cell
        if col_cell[0].value == col_name:
            for it in range(len(list_val)):
                if prefix:
                    if list_val[it][0] == NEGATIVE_SIGN:
                        col_cell[it + 1].value = list_val[it][1:]
                    else:
                        col_cell[it + 1].value = prefix + list_val[it]
                else:
                    col_cell[it + 1].value = list_val[it]
            break


def convert_to_mm(name, column_data):
    if column_data[0].value == name:
        for data_cell in column_data[1:]:
            print(format(float(data_cell.value), ".2f",))
            data_cell.value = format(float(data_cell.value) * CONVERTED_FACTOR, ".2f")


def calculate_new_rot_value(value):
    return 0 if float(value) == ROTATION_270 or float(value) > 360 else float(value) + ROTATION_90


def strip_negative_sign(pos_list):
    return [item[1:] if item[0] == NEGATIVE_SIGN else item for item in pos_list]


def main():
    name_pos_x = "x"
    name_pos_y = "y"
    name_rot = 'r'
    rotated = ROTATION_0
    excel_file, csv_file, wb, sheet = None, None, None, None

    sg.theme("DarkGrey9")
    sg.set_options(font=("Microsoft JhengHei", 16))

    layout_title = [
        [sg.Input(enable_events=True, key='-INPUT-', size=(25, 1)),
         sg.FileBrowse(key='-FILE-', file_types=(("CSV Files", "*.csv"), ("ALL Files", "*.*"))),
         ],
        [sg.Button("Flip-X", enable_events=True, key='-FLIP-X-', font='Helvetica 16', pad=(45, 10),),
         sg.Button("Flip-Y", enable_events=True, key='-FLIP-Y-', font='Helvetica 16', pad=(15, 0),),
         sg.Checkbox("to mm", key="-CHECKBOX-", default=False, enable_events=True, metadata=False, size=(35, 1))
         ],

        [sg.Text("Rotated: {}°".format(rotated), size=(35, 0), key='-text-', font='Helvetica 16', pad=(95, 5))],
        [sg.Button('', enable_events=True, key='-ROTATE-', font='Helvetica 16', image_filename='arrow.png',
                   pad=(110, 0), button_color=('white', 'white')),
         ],
    ]

    window = sg.Window('Fix PnP', layout_title, resizable=False, icon="01_nti_blue.ico", auto_size_buttons=True,
                       size=(450, 250))

    while True:
        event, values = window.read()
        if event == '-INPUT-':
            print(values['-INPUT-'])
            csv_file = values['-INPUT-']
            excel_file = os.path.splitext(csv_file)[0]
            excel_file = excel_file + ".xlsx"
            if Path(csv_file).is_file():
                try:
                    if not Path(excel_file).is_file():
                        csv_to_excel(csv_file, excel_file)
                    wb = openpyxl.load_workbook(excel_file)
                    sheet = wb.active
                    text_elem = window['-text-']
                    text_elem.update("Rotated: {}°".format(ROTATION_0))
                except Exception as e:
                    print("Error: ", e)
        # TODO: check open file
        elif event == "-CHECKBOX-":
            print("check")
            try:
                for column_cell in sheet.iter_cols(1, sheet.max_column):
                    convert_to_mm(name_pos_x, column_cell)
                    convert_to_mm(name_pos_y, column_cell)
                wb.save(excel_file)
                excel_to_csv(csv_file=csv_file, excel_file=excel_file)
            except Exception as e:
                print("Error: ", e)

        elif event == '-FLIP-X-':
            try:
                change_data(sheet, name_pos_x, read_col(sheet, name_pos_x), prefix='-')
                wb.save(excel_file)
                excel_to_csv(csv_file=csv_file, excel_file=excel_file)
            except Exception as e:
                print("Error: ", e)

        elif event == '-FLIP-Y-':
            try:
                change_data(sheet, name_pos_y, read_col(sheet, name_pos_y), prefix='-')
                wb.save(excel_file)
                excel_to_csv(csv_file=csv_file, excel_file=excel_file)
            except Exception as e:
                print("Error: ", e)

        elif event == '-ROTATE-':
            rotated += ROTATION_90
            if rotated > ROTATION_270:
                rotated = ROTATION_0

            text_elem = window['-text-']
            text_elem.update("Rotated: {}°".format(rotated))

            try:
                for column_cell in sheet.iter_cols(1, sheet.max_column):
                    if column_cell[0].value == name_rot:
                        for data in column_cell[1:]:
                            data.value = calculate_new_rot_value(data.value)
                        wb.save(excel_file)
                        excel_to_csv(csv_file=csv_file, excel_file=excel_file)
                        break

                list_pos_x = read_col(sheet, name_pos_x)
                list_pos_y = read_col(sheet, name_pos_y)
                change_data(sheet, name_pos_x, list_pos_y, prefix='-')
                change_data(sheet, name_pos_y, list_pos_x)
                wb.save(excel_file)
                excel_to_csv(csv_file=csv_file, excel_file=excel_file)
            except Exception as e:
                print("Error: ", e)

        elif event in (sg.WIN_CLOSED, 'Exit'):
            break

    window.close()


if __name__ == '__main__':
    main()