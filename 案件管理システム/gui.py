import PySimpleGUI as sg
import font

sg.theme('Reddit')
appicon = b'iVBORw0KGgoAAAANSUhEUgAAACEAAAAhCAYAAABX5MJvAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAnFJREFUeNrsl01rE0Ecxp/dbLMxtZvEHIpZqrd4aAUFtVg8FKy0IhSDn8GDB0/qJxA89V6oF8GD9OShB3sQFLQeLB7EVmxA1KQ5BI027Sa7m+yLO1Oiid1mX7J5OTiwZJglM799nmf+s8ssbmxNo4/t1sSpl5z1+wL9bQyLAWj/IQYKgrMblN4+xeVvi0jxSst4SJwEn3nsaYGirOD59x3/SoRiYwifmKIX6ftp2uZHyPPXgSfL/iFGZu4jfuMRuNGJjiTXtZo/CLIwNzoOQy2D5QVfiw+JScRuX4N+Ng1195d3CL2cRy33BpVXC5A/LHsGME0TIbWAc/FVjAgS6nLVFqQtxJClgqnsgk9f9QVQrUgwdL1l3A6kLUS9uGkBzNFgeslEM8A2Unhg3MOPcPpQEMdgVtcfUku04oYvBfbUGt593SY3DlWkLQQbERA9fxNMRPBtQZgL4Xic/J85qLQF4hzMnTzNBAkoscRPBpLDUcyfGbeOStZbxfxjR3yMArB8DOGTU3i/9gxfVlZsAKynqqmYvTITXNluDubwpbu0X11fQmmvjk9b2QMK6PX9QtQVCLJF1ewq7ZNwXkxPYjpzx9GCQE/RxhYllrjNQNeU+PfwyuVy+FkqwTSMtpOLYqpzCINWyzlashsgRIHksQQSMXfbtixJnUE0bCC7g0BomkYtyFtKyLLiOLlbJRwrZuX1AoVRrWDKqhFIBjwpQbYltcC6lBq736FPKLpeoGM7mgEM869shUIhUDscITSdoQCtkwerxOC+bR+9kMHa6Vkodq9j2c+9++7gjkQRERL9//jpFYhjJnoB4iqY3Qb5LcAA2jc4XKnlrsIAAAAASUVORK5CYII='


def main_window():
    menu_def = [['メニュー', ['初期設定', '終了']],]

    main_layout = [
        [sg.Menu(menu_def)],
        [sg.Text(
            '案件管理アプリ',
            justification='c',
            font=font.hgmaru_l_b)],
        [sg.Button(
            '新規案件登録',
            key='-Move_registration_window-',
            disabled=True,
            tooltip='案件の新規登録',
            font=font.udp_m_i),
            sg.Button(
                '案件管理',
                key='-Move_change_window-',
                disabled=True,
                tooltip='案件の変更/削除',
                font=font.udp_m_i)],
    ]

    return sg.Window('メイン画面', main_layout, icon=appicon, finalize=True)

def initialize_window():
    initialize_layout = [
        [sg.FolderBrowse('保存先'),sg.Input(key='-Folder_path-', enable_events=True)],
        [sg.Button('初期設定', key='-Initialize-', disabled=True, pad=(170, 0))]
    ]

    return sg.Window('初期設定', initialize_layout, finalize=True)

def registration_window():
    registration_layout = [
        [sg.Text('案件管理番号', size=(8, 1)), sg.Input(key='-Case_id-', enable_events=True)],
        [sg.Text('案件名', size=(8, 1)), sg.Input(key='-Case_name-')],
        [sg.Text('案件管理者', size=(8, 1)), sg.Input(key='-Owner-')],
        [sg.Button('登録', key='-Register-')],
    ]

    return sg.Window('新規登録画面', registration_layout, finalize=True)

def change_window():
    changing_layout = [
        [sg.Text(
            '案件管理番号',
            size=(8, 1)),
            sg.Listbox(
            values=[],
            key='Case_id_list',
            select_mode='LISTBOX_SELECT_MODE_SINGLE',
            size=(45, 3),
            enable_events=True)
        ],
        [sg.Text('案件名', size=(8, 1)), sg.Input(key='-Case_name-')],
        [sg.Text('案件管理者', size=(8, 1)), sg.Input(key='-Owner-')],
        [sg.Button('更新', key=('-Update-')), sg.Button('削除', key=('-Dlete-'))],
    ]

    return sg.Window('案件管理画面', changing_layout, finalize=True)