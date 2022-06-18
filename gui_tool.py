import glob
import os
import threading

import pandas as pd
import PySimpleGUI as sg


# 入力ファイルのヘッダー
HEADER = ["通し番号", "著者名", "論文名", "掲載場所", "Vol No", "ページ", "発表年", "概要", "分類番号"]


def set_dpi_awareness():
    """
    高DPIに対応する関数
    """
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass


def load_data(filenames):
    """
    指定されたフォルダやファイルからデータを読み込む関数

    Args:
        filenames: 入力に使用するファイルリスト
    
    Returns:
        df: 読み込んだデータのpandas DataFrame
    """
    df = pd.DataFrame(columns=HEADER)

    # ファイル群のデータを連結
    for filename in filenames:
        tmp_df = pd.read_excel(filename)
        df = df.append(tmp_df)

    # 改行をなくす
    df = df.replace("\n", "", regex=True)

    return df


def etl_data(df, target, filter_no, output_file):
    """
    データをconfigで設定された内容をもとに抽出・加工・出力する関数

    Args:
        df: 対象のpandas DataFrame
        target: 抽出する列名のリスト
        filter_no: フィルタリングを行う分類番号のリスト
        output_file: 出力ファイル名
    """
    etl_data = []
    
    for index, row in df.iterrows():
        concated_str = ""
        if row["分類番号"] in filter_no:
            for t in target:
                # configファイルで指定されたTARGETのデータを文字列として連結
                concated_str = concated_str + str(row[t])

            etl_data.append(concated_str)

    # テキストとして出力
    with open(output_file, "w", encoding="UTF-8") as f:
        for d in etl_data:
            f.write(d + "\n")

    print("[DEBUG] output_file = " + output_file)


def run(lock, window, filenames, target, filter_no, output_file):
    """
    引数を受け取り、処理を実行する関数

    Args:
        lock: threadのlock
        window: pysimpleguiのwindow
        filenames: 入力に使用するファイルリスト
        target: 抽出する列名のリスト
        filter_no: フィルタリングを行う分類番号のリスト
        output_file: 出力ファイル名
    """

    # 既にアプリケーションが起動してた場合は実行しない（バッチ的な処理をしない）
    if lock.locked():
        window.write_event_value('-THREAD-', "ERROR")
        return

    # threadをlock
    lock.acquire()

    # 処理実行
    df = load_data(filenames)
    etl_data(df, target, filter_no, output_file)

    window.write_event_value('-THREAD-', "OK")

    # lockの解放
    lock.release()

    return
    

if __name__ == "__main__":
    # 高DPIに対応
    set_dpi_awareness()

    # threadの競合を回避するためのlock
    lock = threading.Lock()
    
    # デザインテーマの設定
    sg.theme("SystemDefault1")

    # ウィンドウレイアウト
    layout = [
              [sg.Text('フォルダ選択', font=("Meiryo UI", 10)), 
                sg.InputText(enable_events=True, key="-INPUT_FOLDER-"), 
                sg.FolderBrowse('Browse', key='-FOLDER-', initial_folder="./")],
              [sg.Listbox([], size=(100, 10), enable_events=True, key='-LIST-', 
                select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
              [sg.Text('抽出カラム名（複数ある場合 , 区切り）',
                font=("Meiryo UI", 10), size=(28, 1)), sg.InputText(key="-TARGET-")],
              [sg.Text('分類番号（複数ある場合 , 区切り）',
                font=("Meiryo UI", 10), size=(28, 1)), sg.InputText(key="-FILTER_NO-")],
              [sg.Text('出力ファイル（*.txt）', 
                font=("Meiryo UI", 10), size=(28, 1)), 
                sg.InputText("result.txt", key="-OUTPUT_FILE-")],
              [sg.Column([[sg.Button('実行', font=("Meiryo UI", 10), size=(7, 1)),
                sg.Button('終了', font=("Meiryo UI", 10), size=(7, 1))]],
                vertical_alignment='center', justification='center')]
            ]

    # ウィンドウの生成
    window = sg.Window("GUI tool", layout)

    new_files = []
    new_file_names = []

    # イベントループ
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == '終了':
            break
        
        elif event == '実行':
            # 選択されている要素を取得
            filenames = window['-LIST-'].get()
            filenames = [os.path.join(values['-FOLDER-'], f) for f in filenames]

            # GUIからの設定値読み込みと整形
            target = values['-TARGET-'].split(",")
            target = [t.lstrip().rstrip() for t in target]
            filter_no = values['-FILTER_NO-'].split(",")
            filter_no = [int(f) for f in filter_no]
            output_file = values['-OUTPUT_FILE-']

            t = threading.Thread(target=run, args=(lock, window, filenames,
             target, filter_no, output_file), daemon=True)
            t.start()
            
        # -INPUT1-の要素に何か入力された時のイベント
        elif event == '-INPUT_FOLDER-':
            try:
                # 選択したフォルダからエクセルファイルの一覧を取得
                new_files = glob.glob("{0}/*.xlsx".format(values['-FOLDER-']))
                # basenameのみにする
                new_file_names = [os.path.basename(file_path) for file_path in new_files]

                # リストボックスに表示
                window['-LIST-'].update(values=sorted(new_file_names)) 
            except:
                pass
        
        # threadから返答が来た時
        if event == '-THREAD-':
            if values['-THREAD-'] == "OK":
                # 完了メッセージ
                sg.Window("Info", [[sg.T('実行処理が終了しました。')], [sg.Button("OK")]], 
                    disable_close=True).read(close=True)  
            # threadのlockを取得できなかった場合
            else:
                # エラーメッセージ
                sg.Window("Error", [[sg.T('実行中の処理があります。')], [sg.Button("OK")]], 
                    disable_close=True).read(close=True)
    
    window.close()
