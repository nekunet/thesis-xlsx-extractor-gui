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


def etl_data(df):
    """
    データをconfigで設定された内容をもとに抽出・加工・出力する関数

    Args:
        df: 対象のpandas DataFrame    
    """
    etl_data = []
    
    target = json.loads(config["DEFAULT"]["TARGET"])
    filter_no = json.loads(config["DEFAULT"]["FILTER_NO"])
    output_file = config["DEFAULT"]["OUTPUT_FILE"]
    
    for index, row in df.iterrows():
        concated_str = ""
        if row["分類番号"] in filter_no:
            for t in target:
                # configファイルで指定されたTARGETのデータを文字列として連結
                concated_str = concated_str + str(row[t])
                
            etl_data.append(concated_str)

    # テキストとして出力
    with open(output_file, "w") as f:
        for d in etl_data:
            f.write(d + "\n")

    print("[INFO] {0} に結果を出力しました。".format(output_file))


def main():
     df = load_data()
     etl_data(df)
     

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read("setting.txt", encoding="utf-8")
    main()
