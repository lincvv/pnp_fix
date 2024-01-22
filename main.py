import PySimpleGUI as sg
from pathlib import Path
import openpyxl

NEGATIVE_SIGN = '-'
ROTATION_270 = 270
ROTATION_90 = 90
ROTATION_180 = 180
ROTATION_0 = 0
rotated = ROTATION_0
CONVERTED_FACTOR = 0.0254
name_pos_x = "PosX"
name_pos_y = "PosY"
name_Rot = 'Rot'
# Position = {name_pos_x: list(), name_pos_y: list()}


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
        # print(col_cell)
        if col_cell[0].value == col_name:    # check for your column
            for it in range(len(list_val)):
                if prefix:
                    if list_val[it][0] == NEGATIVE_SIGN:
                        col_cell[it + 1].value = list_val[it][1:]
                    else:
                        col_cell[it + 1].value = prefix + list_val[it]
                else:
                    col_cell[it + 1].value = list_val[it]
            wb.save(filename)
            break


def convert_to_mm(name, column_data):
    if column_data[0].value == name:
        for data_cell in column_data[1:]:
            print(format(float(data_cell.value), ".2f"))
            data_cell.value = format(float(data_cell.value) * CONVERTED_FACTOR, ".2f")
        wb.save(filename)


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


sg.theme("DarkGrey9")
sg.set_options(font=("Microsoft JhengHei", 16))

layout_title = [
    [sg.Input(key='-INPUT-', size=(25, 1)),
     sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"), ("ALL Files", "*.*"))),
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

# window = sg.Window('Title', layout_title, size=(350,100))
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break

    elif event == "-CHECKBOX-":
        filename = values['-INPUT-']
        if Path(filename).is_file():
            if values["-CHECKBOX-"]:
                print("mm")
                try:
                    wb = openpyxl.load_workbook(filename)
                    sheet = wb.active
                    for column_cell in sheet.iter_cols(1, sheet.max_column):
                        convert_to_mm(name_pos_x, column_cell)
                        convert_to_mm(name_pos_y, column_cell)
                except Exception as e:
                    print("Error: ", e)

    elif event == 'Read':
        filename = values['-INPUT-']
        if Path(filename).is_file():
            try:
                wb = openpyxl.load_workbook(filename)
                sheet = wb.active
                #
                # copied_pos_x = list_pos_x.copy()
                # if list_pos_x[0][0] == NEGATIVE_SIGN and list_pos_y[0][0] == NEGATIVE_SIGN:
                #     # Position[name_pos_x] = strip_negative_sign(list_pos_x)
                #     # Position[name_pos_y] = strip_negative_sign(list_pos_y)
                #     rotated = ROTATION_180
                # elif list_pos_x[0][0] == NEGATIVE_SIGN:
                #     # Position[name_pos_x] = list_pos_y
                #     # Position[name_pos_y] = strip_negative_sign(copied_pos_x)
                #     rotated = ROTATION_90
                # elif list_pos_y[0][0] == NEGATIVE_SIGN:
                #     # Position[name_pos_x] = strip_negative_sign(list_pos_y)
                #     # Position[name_pos_y] = copied_pos_x
                #     rotated = ROTATION_270
                # else:
                #     rotated = ROTATION_0

                text_elem = window['-text-']
                text_elem.update("Rotated: {}°".format(rotated))

            except Exception as e:
                print("Error: ", e)
    elif event == '-ROTATE-':
        rotated += ROTATION_90
        if rotated > ROTATION_270:
            rotated = ROTATION_0

        text_elem = window['-text-']
        text_elem.update("Rotated: {}°".format(rotated))

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

                list_pos_x = read_col(sheet, name_pos_x)
                list_pos_y = read_col(sheet, name_pos_y)
                change_data(sheet, name_pos_x, list_pos_y, prefix='-')
                change_data(sheet, name_pos_y, list_pos_x)

            except Exception as e:
                print("Error: ", e)

window.close()
