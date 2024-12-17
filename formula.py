import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
import csv
import argparse
import sys

# デバッグフラグ
DEBUG = False

# デバッグメッセージ出力
def debug_print(message):
    if DEBUG:
        print(f"[DEBUG] {message}")

# Tableauファイルの解凍
def unzip_file(source_file, target_folder):
    try:
        debug_print(f"Unzipping {source_file} to {target_folder}")
        with zipfile.ZipFile(source_file, 'r') as zip_ref:
            zip_ref.extractall(target_folder)
        return True
    except Exception as e:
        print(f"Error during unzip: {e}")
        return False

# データをCSV形式で出力
def write_to_csv(file_path, data):
    try:
        debug_print(f"Writing data to {file_path}")
        with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file)
            writer.writerows(data)
    except Exception as e:
        print(f"Error writing CSV: {e}")

# データを標準出力に出力
def write_to_stdout(data):
    debug_print("Writing data to stdout")
    if sys.stdout.isatty():
        writer = csv.writer(sys.stdout)
    else:
        writer = csv.writer(open(sys.stdout.fileno(), mode='w', newline='', encoding='utf-8-sig', closefd=False))
    writer.writerows(data)

# XMLから数式を解析
def parse_formula(xml_node):

    data = [["DataSource", "Column", "Formula"]]

    # XML解析ロジック
    for datasource in xml_node.findall(".//datasource"):
        datasource_name = datasource.get("name", "")
        datasource_caption = datasource.get("caption", datasource_name)
        column_mapping = {}
        columns = {}
        for column in datasource.findall(".//column"):
            column_name = column.get("name", "")
            column_caption = column.get("caption", column_name)
            column_mapping[column_name] = column_caption
            calculation = column.find("./calculation") 
            formula = calculation.get("formula", "") if calculation is not None else ""
            formula = formula.replace(datasource_name, datasource_caption)  # データソースが含まれていたら name を caption に置換
            formula = formula.replace("[Parameters].", "")  # "[Parameters]." を削除
            columns[column_name] = {"caption": column_caption, "formula": formula}

        # データを整理して配列に格納
        for column_name, column_data in columns.items():
            column_caption = column_data["caption"]
            formula = column_data["formula"]
            # column_mapping を使用して name を caption に置換
            for col_name, col_caption in column_mapping.items():
                if not col_name.startswith("["):
                    col_name = f"[{col_name}]"
                if not col_caption.startswith("["):
                    col_caption = f"[{col_caption}]"
                formula = formula.replace(col_name, col_caption)
            data.append([datasource_caption, column_caption, formula])

    # datasource_caption と column_caption の昇順でソート
    data_sorted = sorted(data[1:], key=lambda x: (x[0], x[1]))  # 先頭のヘッダーはソートから除外
    data_sorted.insert(0, data[0])  # ヘッダーを先頭に戻す

    return data_sorted

# 主処理
def open_tableau(input_file, output_file=None):
    debug_print(f"Processing input file: {input_file}")
    # 拡張子の判定
    file_ext = os.path.splitext(input_file)[-1].lower()
    temp_folder = None
    twb_file_path = input_file

    if file_ext == ".twbx":
        debug_print(f"Input file is a TWBX package")
        temp_folder = os.path.join(os.environ["TEMP"], "tableau_temp")
        if os.path.exists(temp_folder):
            debug_print(f"Cleaning up existing temp folder: {temp_folder}")
            shutil.rmtree(temp_folder)
        os.makedirs(temp_folder)

        if unzip_file(input_file, temp_folder):
            found_twb_file = False
            for root, _, files in os.walk(temp_folder):
                for file in files:
                    if file.endswith(".twb"):
                        twb_file_path = os.path.join(root, file)
                        debug_print(f"Found TWB file: {twb_file_path}")
                        found_twb_file = True
                        break
                if found_twb_file:
                    break
            else:
                print("TWB file not found in TWBX archive.")
                return
        else:
            print("Failed to unzip TWBX file.")
            return

    # XMLを読み込む
    try:
        debug_print(f"Parsing XML from {twb_file_path}")
        tree = ET.parse(twb_file_path)
        root = tree.getroot()
        datasources_node = root.find(".//datasources")

        if datasources_node is not None:
            data = parse_formula(datasources_node)

            # 出力先に応じて処理
            if output_file:
                write_to_csv(output_file, data)
                print(f"Data saved to {output_file}")
            else:
                write_to_stdout(data)
        else:
            print("No datasources found in the Tableau workbook.")
    except Exception as e:
        print(f"Error parsing XML: {e}")
    finally:
        # 解凍フォルダのクリーンアップ
        if temp_folder and os.path.exists(temp_folder):
            debug_print(f"Cleaning up temp folder: {temp_folder}")
            shutil.rmtree(temp_folder)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process a Tableau workbook and export formulas to CSV or stdout.")
    parser.add_argument("input_file", type=str, help="Path to the Tableau workbook (.twb or .twbx).")
    parser.add_argument(
        "-o", "--output_file", type=str, default=None,
        help="Path to the output CSV file. If not specified, output is written to stdout."
    )

    args = parser.parse_args()

    # デバッグモードを設定
    debug_print("Debug mode is enabled")

    # 引数からファイルパスを取得して処理を実行
    open_tableau(args.input_file, args.output_file)
