import configparser
import glob
import json
import os

import pandas as pd


# 入力ファイルのヘッダー
HEADER = ["通し番号", "著者名", "論文名", "掲載場所", "Vol No", "ページ", "発表年", "概要", "分類番号"]


def get_filenames():
    """
    configから入力ファイル群を読み込む関数

    Returns:
        filename: 入力に使用するのファイル群
    
    """
    # ディレクトリのパスをconfigから抽出 
    input_dir = config["DEFAULT"]["INPUT_DIR"]

    # ファイル群をconfigから抽出
    if config["DEFAULT"]["INPUT_FILE"]:
        filenames = json.loads(config["DEFAULT"]["INPUT_FILE"])
        filenames = [os.path.join(input_dir, filename) for filename in filenames]
    else:
        # もしファイル名の指定がなければ全件抽出
        filenames = glob.glob("{0}/*.xlsx".format(input_dir))

    return filenames


def load_data():
    """
    指定されたフォルダやファイルからデータを読み込む関数
    
    Returns:
        df: 読み込んだデータのpandas DataFrame
    """
    filenames = get_filenames()

    df = pd.DataFrame(columns=HEADER)

    # ファイル群のデータを連結
    for filename in filenames:
        tmp_df = pd.read_excel(filename)
        df = df.append(tmp_df)

    # 改行をなくす
    df = df.replace("\n", "", regex=True)

    return df


def main():
     df = load_data()
     


if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("setting.txt", encoding="utf-8")
    main()
