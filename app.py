import streamlit as st
import pandas as pd
from io import BytesIO
import math
import re
import os

st.set_page_config(
    page_title="STORES Order Converter",
    page_icon="🍒",
    layout="centered"
)

st.title("🍒 STORES Order Converter")
st.write("STORESの注文CSVファイルをExcelファイルへ変換します。")
st.caption("※アップロードされたファイルはサーバーに保存されず、メモリ上の一時処理後に破棄されます。")

custom_filename = st.text_input(
    "保存するファイル名を指定（空欄の場合は元のファイル名＋_オーダーまとめと出力されます）",
    placeholder="例：202605_夏サーファン"
)

def make_safe_sheet_name(name, used_names):
    base = str(name)
    base = re.sub(r'[\[\]\:\*\?\/\\]', '_', base)
    base = base[:31]
    if base.strip() == "":
        base = "sheet"
    
    sheet_name = base
    counter = 1
    while sheet_name in used_names:
        suffix = f"_{counter}"
        sheet_name = (base[:31-len(suffix)]) + suffix
        counter += 1
    
    used_names.add(sheet_name)
    return sheet_name

uploaded_file = st.file_uploader(
    "STORESから出力したCSVファイルを選択してください",
    type=["csv"]
)

if uploaded_file is not None:
    try:
        try:
            df = pd.read_csv(uploaded_file, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="cp932")

        required_columns = [
            '氏(配送先)', '名(配送先)', 'アイテム名', '種類', '個数',
            '郵便番号(配送先)', '都道府県(配送先)', '住所(配送先)', '備考'
        ]
        missing = [col for col in required_columns if col not in df.columns]
        
        if missing:
            st.error(f"CSVファイルの形式が正しくありません。次の必須列が不足しています: {missing}")
        else:
            st.success("CSVファイルを正常に読み込みました。Excelファイルへの変換を行います。")
            st.dataframe(df.head(3))

            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book

                headers = required_columns
                sorted_df = df[headers].sort_values(by=['氏(配送先)', '名(配送先)']).reset_index(drop=True)
                worksheet_operate = workbook.add_worksheet("運営用一覧")

                for col_idx, header in enumerate(headers):
                    worksheet_operate.write(0, col_idx, header)

                worksheet_operate.set_column(0, 0, 12)
                worksheet_operate.set_column(1, 1, 12)
                worksheet_operate.set_column(2, 2, 40)
                worksheet_operate.set_column(3, 3, 15)
                worksheet_operate.set_column(4, 4, 6)
                worksheet_operate.set_column(5, 5, 10)
                worksheet_operate.set_column(6, 6, 12)
                worksheet_operate.set_column(7, 7, 55)
                worksheet_operate.set_column(8, 8, 35)

                row_idx = 1
                last_name = ""
                for _, row in sorted_df.iterrows():
                    full_name = f"{row['氏(配送先)']} {row['名(配送先)']}"
                    if last_name and full_name != last_name:
                        row_idx += 1  
                    last_name = full_name

                    for col_idx, col_name in enumerate(headers):
                        val = row[col_name]
                        if pd.isna(val):
                            worksheet_operate.write(row_idx, col_idx, "")
                        elif col_name == '個数':
                            try:
                                number = float(val)
                                if math.isnan(number) or math.isinf(number):
                                    worksheet_operate.write(row_idx, col_idx, "")
                                else:
                                    worksheet_operate.write_number(row_idx, col_idx, number)
                            except:
                                worksheet_operate.write(row_idx, col_idx, str(val))
                        else:
                            worksheet_operate.write(row_idx, col_idx, str(val))
                    row_idx += 1

                item_names = df['アイテム名'].dropna().unique()
                used_names = set(["運営用一覧"])

                for item in item_names:
                    df_item = df[df['アイテム名'] == item].copy()
                    sheet_name = make_safe_sheet_name(item, used_names)
                    worksheet = workbook.add_worksheet(sheet_name)

                    worksheet.set_column(0, 0, 12)
                    worksheet.set_column(1, 1, 12)
                    worksheet.set_column(2, 2, 15)
                    worksheet.set_column(3, 3, 6)
                    worksheet.set_column(4, 4, 35)

                    row_idx = 0
                    item_total = pd.to_numeric(df_item['個数'], errors='coerce').fillna(0).sum()
                    
                    worksheet.write(row_idx, 0, f"■ {item} 全体合計（計{int(item_total)}個）")
                    row_idx += 2

                    type_names = df_item['種類'].fillna("").unique()

                    for typ in type_names:
                        if typ == "":
                            df_type = df_item[df_item['種類'].isna() | (df_item['種類'] == "")]
                            title = "種類なし"
                        else:
                            df_type = df_item[df_item['種類'] == typ]
                            title = str(typ)

                        total = pd.to_numeric(df_type['個数'], errors='coerce').fillna(0).sum()

                        worksheet.write(row_idx, 0, f"■ {title}（計{int(total)}個）")
                        row_idx += 1

                        worksheet.write(row_idx, 0, '氏')
                        worksheet.write(row_idx, 1, '名')
                        worksheet.write(row_idx, 2, '種類')
                        worksheet.write(row_idx, 3, '個数')
                        worksheet.write(row_idx, 4, '備考')
                        row_idx += 1

                        for _, row in df_type.iterrows():
                            worksheet.write(row_idx, 0, str(row['氏(配送先)']) if pd.notna(row['氏(配送先)']) else "")
                            worksheet.write(row_idx, 1, str(row['名(配送先)']) if pd.notna(row['名(配送先)']) else "")
                            worksheet.write(row_idx, 2, str(row['種類']) if pd.notna(row['種類']) else "")

                            try:
                                number = float(row['個数'])
                                if math.isnan(number) or math.isinf(number):
                                    worksheet.write(row_idx, 3, "")
                                else:
                                    worksheet.write_number(row_idx, 3, number)
                            except:
                                worksheet.write(row_idx, 3, str(row['個数']))

                            worksheet.write(row_idx, 4, str(row['備考']) if pd.notna(row['備考']) else "")
                            row_idx += 1

                        row_idx += 1


            if custom_filename.strip():
                filename_base = custom_filename.strip()
                if not filename_base.endswith('.xlsx'):
                    output_filename = f"{filename_base}.xlsx"
                else:
                    output_filename = filename_base
            else:
                base_filename = os.path.splitext(uploaded_file.name)[0]
                output_filename = f"{base_filename}_オーダーまとめ.xlsx"

            st.download_button(
                label="フォーマット済みExcelファイルをダウンロード",
                data=output.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error("変換処理中に予期せぬエラーが発生しました。")
        st.exception(e)