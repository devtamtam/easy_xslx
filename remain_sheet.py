import openpyxl
import os
import re # 変更点: 正規表現を扱うreモジュールをインポート

def clean_excel_sheets():
    """
    ユーザーに対話形式でExcelファイルのパスを尋ね、
    選択されたシート以外を削除して新しいファイルとして保存する関数。
    """
    # 1. ファイルパスの入力
    while True:
        file_path = input("作業したいExcelファイル（.xlsx）のフルパスを貼り付けてください: ")
        if os.path.exists(file_path) and file_path.lower().endswith('.xlsx'):
            break
        else:
            print("エラー: 正しいファイルパスではありません。もう一度お試しください。")

    try:
        # Excelファイルを読み込み、シート名を取得
        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.sheetnames
        
        # 2. シート一覧の表示
        print(f"\n--- 「{os.path.basename(file_path)}」のシート一覧 ---")
        for i, name in enumerate(sheet_names):
            print(f"  {i}: {name}")
        print("---------------------------------")

        # 3. 残したいシートの選択
        while True:
            try:
                # 変更点: 入力形式の説明を更新
                user_input = input("残したいシートの番号をカンマやスペース区切りで入力してください (例: 0,2 3): ")
                
                # 変更点: カンマとスペースの両方で分割し、空の要素をフィルタリング
                raw_indices = re.split(r'[,\s]+', user_input.strip())
                selected_indices = {int(i) for i in raw_indices if i} # 空の文字列を除外

                if all(0 <= i < len(sheet_names) for i in selected_indices):
                    break
                else:
                    print("エラー: 存在しない番号が入力されました。")
            except ValueError:
                print("エラー: 番号を正しく入力してください。")

        # 4. シートの削除処理
        sheets_to_keep = {sheet_names[i] for i in selected_indices}
        for sheet in wb.worksheets:
            if sheet.title not in sheets_to_keep:
                wb.remove(sheet)
                print(f'シート "{sheet.title}" を削除しました。')
        
        # 5. 新しいファイルとして保存
        new_file_path = os.path.join(os.path.dirname(file_path), f"cleaned_{os.path.basename(file_path)}")
        wb.save(new_file_path)

        print("\n✨ 完了しました！")
        print(f'残したシート: {", ".join(sorted(list(sheets_to_keep)))}')
        print(f'新しいファイルが「{new_file_path}」として保存されました。')

    except Exception as e:
        print(f"エラーが発生しました: {e}")

# スクリプトの実行
if __name__ == "__main__":
    clean_excel_sheets()