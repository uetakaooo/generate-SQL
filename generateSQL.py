import pandas as pd
import glob
import os
import configparser

# テーブル名 (適宜変更してください)
table_name = 'KYF0010'

# ConfigParserを使用して設定ファイルからディレクトリのパスを読み取る
config = configparser.ConfigParser()
config.read('config.ini', 'UTF-8')
directory = config['DEFAULT']['Directory']

# ディレクトリ内のすべてのエクセルファイルを取得します
excel_files = glob.glob(os.path.join(directory, '*.xlsx'))

def get_column_range(sheet, max_col):
    """シートの列数に基づいて列範囲を取得する"""
    columns = sheet.columns[:max_col]
    return columns

for excel_file_path in excel_files:
    # ファイル名を取得してSQLファイル名に使用します
    file_name = os.path.basename(excel_file_path).split('.')[0]

    # エクセルファイル内のシート名を取得
    xls = pd.ExcelFile(excel_file_path)
    sheet_names = xls.sheet_names

    # シートの存在を確認してデータを読み込みます
    if len(sheet_names) > 1:
        kyf0010data = pd.read_excel(excel_file_path, sheet_name=1, dtype=str)
        kyf0010columns = get_column_range(kyf0010data, 115)  # 存在する列数に基づいて列範囲を取得
        kyf0010data = kyf0010data[kyf0010columns]
    else:
        kyf0010data = pd.read_excel(excel_file_path, sheet_name=0, dtype=str)
        kyf0010columns = get_column_range(kyf0010data, 115)  # 存在する列数に基づいて列範囲を取得
        kyf0010data = kyf0010data[kyf0010columns]

    if len(sheet_names) > 0:
        infosheet = pd.read_excel(excel_file_path, sheet_name=0, usecols="A:G", dtype=str)
    else:
        raise ValueError("インフォシートが見つかりません")

    # INSERT文とDELETE文を保存するディレクトリ
    sql_file = os.path.join(directory, f'{file_name}.sql')

    # SQL文を作成する
    with open(sql_file, 'w', encoding='utf-8') as file:
        for index, row in kyf0010data.iterrows():
            primary_key_column = 'KYAK_CIF_C'
            primary_key_value = row[primary_key_column]
            columns = ', '.join(row.index)
            values = ', '.join([f"'{value}'" for value in row.values])
            delete_statement = f"DELETE FROM {table_name} WHERE {primary_key_column} = '{primary_key_value}';\n"
            insert_statement = f"INSERT INTO {table_name} ({columns}) VALUES ({values});\n"
            commit_statement = 'commit;\n\n'

            # コメント作成
            comment_statements = []
            for i, info_row in infosheet.iterrows():
                comment_key_column1 = 'No'
                comment_key_column2 = 'ASTAID'
                comment_key_column3 = '対象API名'
                comment_key_column4 = 'ケース番号'
                comment_key_column5 = '顧客情報'
                key_column = '口座番号'
                key_value = info_row[key_column]
                comment_key_value1 = info_row[comment_key_column1]
                comment_key_value2 = info_row[comment_key_column2]
                comment_key_value3 = info_row[comment_key_column3]
                comment_key_value4 = info_row[comment_key_column4]
                comment_key_value5 = info_row[comment_key_column5]
                if key_value == primary_key_value:
                    comment_statement = f"-- {comment_key_value1}. ASTAID_{comment_key_value2} Case{comment_key_value4} - {comment_key_value3} ({comment_key_value5})\n"
                    comment_statements.append(comment_statement)

            # コメントをSQLファイルに書き込む
            for comment in comment_statements:
                file.write(comment)

            # DELETE, INSERT, COMMIT文を書き込む
            file.write(delete_statement)
            file.write(insert_statement)
            file.write(commit_statement)

    print(f"{file_name}.sql が生成されました。")

print("すべてのSQLファイルの生成が完了しました。")