import PySimpleGUI as sg
from pathlib import Path
import openpyxl

NEGATIVE_SIGN = '-'
ROTATION_270 = 270
ROTATION_90 = 90
ROTATION_180 = 180
ROTATION_0 = 0
rotated = ROTATION_0
name_pos_x = "PosX"
name_pos_y = "PosY"
name_Rot = 'Rot'
Position = {name_pos_x: list(), name_pos_y: list()}


def change_data(sh, col_name, list_val, prefix=None):
    for col_cell in sh.iter_cols(1, sh.max_column):  # iterate column cell
        # print(col_cell)
        if col_cell[0].value == col_name:    # check for your column
            for it in range(len(list_val)):
                col_cell[it + 1].value = list_val[it] if prefix is None else prefix + list_val[it]
            wb.save(filename)
            break


def append_data_to_position(name, column_data):
    if column_data[0].value == name:
        for data in column_data[1:]:
            Position[name].append(data.value)


def calculate_new_rot_value(value):
    return 0 if int(value) == ROTATION_270 else int(value) + ROTATION_90


def strip_negative_sign(pos_list):
    return [item[1:] if item[0] == NEGATIVE_SIGN else item for item in pos_list]


def popup_text(filename, text):

    layout = [
        [sg.Multiline(text, size=(80, 25)),],
    ]
    win = sg.Window(filename, layout, modal=True, finalize=True)

    while True:
        event, values = win.read()
        if event == sg.WINDOW_CLOSED:
            break
    win.close()


sg.theme("DarkBlue3")
sg.set_options(font=("Microsoft JhengHei", 16))

layout_title = [
    [sg.Input(key='-INPUT-', size=(35, 1)),
     sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"), ("ALL Files", "*.*"))),
     sg.Button("Read")
     ],

    [sg.Text('Rotated:', size=(25, 1), key='-text-', font='Helvetica 16')],
    [sg.Button('Rotate ->', enable_events=True, key='-ROTATE-', font='Helvetica 16')],
]

window = sg.Window('Title', layout_title)

# window = sg.Window('Title', layout_title, size=(350,100))
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    elif event == 'Read':
        filename = values['-INPUT-']
        if Path(filename).is_file():
            try:
                wb = openpyxl.load_workbook(filename)
                sheet = wb.active
                for column_cell in sheet.iter_cols(1, sheet.max_column):
                    append_data_to_position(name_pos_x, column_cell)
                    append_data_to_position(name_pos_y, column_cell)

                list_pos_x = Position[name_pos_x]
                list_pos_y = Position[name_pos_y]

                copied_pos_x = list_pos_x.copy()
                if list_pos_x[0][0] == NEGATIVE_SIGN and list_pos_y[0][0] == NEGATIVE_SIGN:
                    Position[name_pos_x] = strip_negative_sign(list_pos_x)
                    Position[name_pos_y] = strip_negative_sign(list_pos_y)
                    rotation = ROTATION_180
                elif list_pos_x[0][0] == NEGATIVE_SIGN:
                    Position[name_pos_x] = list_pos_y
                    Position[name_pos_y] = strip_negative_sign(copied_pos_x)
                    rotation = ROTATION_90
                elif list_pos_y[0][0] == NEGATIVE_SIGN:
                    Position[name_pos_x] = strip_negative_sign(list_pos_y)
                    Position[name_pos_y] = copied_pos_x
                    rotation = ROTATION_270
                else:
                    rotated = ROTATION_0

                text_elem = window['-text-']
                text_elem.update("Rotated: {}".format(rotated))

            except Exception as e:
                print("Error: ", e)
    elif event == '-ROTATE-':
        rotated += ROTATION_90
        if rotated > ROTATION_270:
            rotated = ROTATION_0
        text_elem = window['-text-']
        text_elem.update("Rotated: {}".format(rotated))

        filename = values['-INPUT-']
        if Path(filename).is_file():
            try:
                wb = openpyxl.load_workbook(filename)
                sheet = wb.active
                for column_cell in sheet.iter_cols(1, sheet.max_column):
                    if column_cell[0].value == name_Rot:
                        cell = column_cell[0]
                        for data in column_cell[1:]:
                            data.value = calculate_new_rot_value(data.value)
                        wb.save(filename)
                        break

                if rotated == 90:
                    change_data(sheet, name_pos_x, Position[name_pos_y], prefix='-')
                    change_data(sheet, name_pos_y, Position[name_pos_x])
                    print(" Change data rotated to 90\n")
                elif rotated == 180:
                    change_data(sheet, name_pos_x, Position[name_pos_x], prefix='-')
                    change_data(sheet, name_pos_y, Position[name_pos_y], prefix='-')
                    print(" Change data rotated to 180\n")
                elif rotated == 270:
                    change_data(sheet, name_pos_x, Position[name_pos_y])
                    change_data(sheet, name_pos_y, Position[name_pos_x], prefix='-')
                    print(" Change data rotated to 270\n")
                else:
                    change_data(sheet, name_pos_x, Position[name_pos_x])
                    change_data(sheet, name_pos_y, Position[name_pos_y])
                    print(" Change data rotated to 0\n")

            except Exception as e:
                print("Error: ", e)

window.close()
