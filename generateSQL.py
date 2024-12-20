import pandas as pd
import glob
import os

# エクセルファイルが格納されているディレクトリのパス
directory_path = r'C:\workspace\generate_SQL\work'

# テーブル名 (適宜変更してください)
table_name = 'KYF0010'

# INSERT文とDELETE文を保存するディレクトリ
output_directory = 'output_sql_files'

# 出力ディレクトリが存在しない場合は作成します
os.makedirs(output_directory, exist_ok=True)

# ディレクトリ内のすべてのエクセルファイルを取得します
excel_files = glob.glob(os.path.join(directory_path, '*.xlsx'))

# ゼロサプレスを防ぎたいカラム名のリスト（例: 電話番号やIDなど）
string_columns = ['KYAK_CIF_C']

for excel_file_path in excel_files:
    # ファイル名を取得してSQLファイル名に使用します
    file_name = os.path.basename(excel_file_path).split('.')[0]

    # データを読み込みます
    df = pd.read_excel(excel_file_path, dtype=str)  # すべてのカラムを文字列として読み込む

    # INSERT文とDELETE文を保存するファイルパス
    insert_sql_file = os.path.join(output_directory, f'{file_name}_insert_statements.sql')
    delete_sql_file = os.path.join(output_directory, f'{file_name}_delete_statements.sql')

    # INSERT文を生成します
    with open(insert_sql_file, 'w', encoding='utf-8') as insert_file:
        for index, row in df.iterrows():
            columns = ', '.join(row.index)
            values = ', '.join([f"'{value}'" for value in row.values])
            insert_statement = f"INSERT INTO {table_name} ({columns}) VALUES ({values});\n"
            insert_file.write(insert_statement)

    # DELETE文を生成します (例: 主キーが 'KYAK_CIF_C' である場合)
    with open(delete_sql_file, 'w', encoding='utf-8') as delete_file:
        for index, row in df.iterrows():
            # 主キーのカラム名 (適宜変更してください)
            primary_key_column = 'KYAK_CIF_C'
            primary_key_value = row[primary_key_column]
            delete_statement = f"DELETE FROM {table_name} WHERE {primary_key_column} = '{primary_key_value}';\n"
            delete_file.write(delete_statement)

    print(f"{file_name}_insert_statements.sql と {file_name}_delete_statements.sql が生成されました。")

print("すべてのSQLファイルの生成が完了しました。")