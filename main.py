import PySimpleGUI as sg
from pathlib import Path
import openpyxl

NEGATIVE_SIGN = '-'
ROTATION_270 = 270
ROTATION_90 = 90
ROTATION_180 = 180
ROTATION_0 = 0
CONVERTED_FACTOR = 0.0254


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
    return 0 if int(value) == ROTATION_270 else int(value) + ROTATION_90


def strip_negative_sign(pos_list):
    return [item[1:] if item[0] == NEGATIVE_SIGN else item for item in pos_list]


def main():
    name_pos_x = "PosX"
    name_pos_y = "PosY"
    name_rot = 'Rot'
    rotated = ROTATION_0
    filename, wb, sheet = None, None, None

    sg.theme("DarkGrey9")
    sg.set_options(font=("Microsoft JhengHei", 16))

    layout_title = [
        [sg.Input(enable_events=True, key='-INPUT-', size=(25, 1)),
         sg.FileBrowse(key='-FILE-', file_types=(("XLSX Files", "*.xlsx"), ("ALL Files", "*.*"))),
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
            filename = values['-INPUT-']
            if Path(filename).is_file():
                try:
                    wb = openpyxl.load_workbook(filename)
                    sheet = wb.active
                except Exception as e:
                    print("Error: ", e)

        elif event == "-CHECKBOX-":
            try:
                for column_cell in sheet.iter_cols(1, sheet.max_column):
                    convert_to_mm(name_pos_x, column_cell)
                    convert_to_mm(name_pos_y, column_cell)
                wb.save(filename)
                print('mm')
            except Exception as e:
                print("Error: ", e)

        elif event == '-FLIP-X-':
            try:
                change_data(sheet, name_pos_x, read_col(sheet, name_pos_x), prefix='-')
                wb.save(filename)
            except Exception as e:
                print("Error: ", e)

        elif event == '-FLIP-Y-':
            try:
                change_data(sheet, name_pos_y, read_col(sheet, name_pos_y), prefix='-')
                wb.save(filename)
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
                        wb.save(filename)
                        break

                list_pos_x = read_col(sheet, name_pos_x)
                list_pos_y = read_col(sheet, name_pos_y)
                change_data(sheet, name_pos_x, list_pos_y, prefix='-')
                change_data(sheet, name_pos_y, list_pos_x)
                wb.save(filename)
            except Exception as e:
                print("Error: ", e)

        elif event in (sg.WIN_CLOSED, 'Exit'):
            break

    window.close()


if __name__ == '__main__':
    main()