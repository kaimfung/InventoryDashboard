import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import base64  # 添加 base64 模組

import streamlit as st
from google.oauth2.service_account import Credentials
import gspread

# 設定 Google Sheet 認證
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# 從 Streamlit Secrets 讀取認證資訊
creds_dict = st.secrets["gcp_service_account"]

# 使用字典創建認證物件
creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
client = gspread.authorize(creds)

# 連接到 Google Sheet
sheet = client.open("INVENTORY_API")

# 從 Google Sheet 讀取數據
def get_data_from_google_sheet():
    # 讀取所有工作表的數據
    weeks = ["week 1", "week 2", "week 3", "week 4", "week 5"]
    dataframes = {}
    update_dates = {}

    for week in weeks:
        try:
            worksheet = sheet.worksheet(week)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"找不到工作表 '{week}'。請確認 Google Sheet 中是否存在該工作表。")
            st.stop()
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        # 確保欄位名稱一致
        df.columns = df.columns.str.strip()
        # 將 Quantity 欄位轉為數值型，並將負數轉為 0
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").clip(lower=0).fillna(0)
        # 重命名欄位
        df = df.rename(columns={
            "G-Sub Group(Name)": "Sub Group",
            "G-Loc/Brand(Name)": "Brand",
            "Description": "Desc",
            "Location Name": "Location",
            "Unit": "Unit"
        })
        dataframes[week] = df
        # 讀取 I1 儲存格的更新日期
        update_date = worksheet.acell("I1").value
        update_dates[week] = update_date

    # 按照工作表名稱順序排列（week 1 到 week 5）
    sorted_weeks = weeks  # 直接使用原始順序
    return dataframes, update_dates, sorted_weeks

# 將 DataFrame 轉為 HTML 表格，並應用樣式
def df_to_html_table(df, update_dates, sorted_weeks, usage_threshold=None, is_low_stock=False):
    # 獲取日期欄位和變化欄位
    date_columns = [update_dates[week] for week in sorted_weeks]
    change_columns = [col for col in df.columns if "-" in col]

    # 格式化數字，保留 2 位小數
    for col in date_columns + change_columns:
        df[col] = df[col].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")

    # 開始構建 HTML 表格
    html = """
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
            overflow-x: auto;
            display: block;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
            white-space: nowrap;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        tr:nth-child(odd) {
            background-color: #ffffff;
        }
    </style>
    <table>
        <thead>
            <tr>
    """

    # 添加表頭
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"

    # 添加數據行
    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            value = row[col]
            style = ""

            # 為日期欄位設置樣式
            if col in date_columns:
                style = 'background-color: #e6f3ff; color: #000000;'
                if is_low_stock and col == update_dates["week 1"]:
                    week1_stock = float(row[update_dates["week 1"]]) if row[update_dates["week 1"]] != "N/A" else float('inf')
                    if pd.notna(usage_threshold) and week1_stock < usage_threshold:
                        style = 'background-color: #ffcccc; color: #000000;'
            # 為變化欄位設置樣式
            elif col in change_columns:
                try:
                    value_float = float(value) if value != "N/A" else 0
                    if value_float > 0:
                        style = 'background-color: #ccffcc; color: #000000;'
                    elif value_float < 0:
                        style = 'background-color: #ffcccc; color: #000000;'
                    else:
                        style = 'background-color: #ffffff; color: #000000;'
                except:
                    style = 'background-color: #ffffff; color: #000000;'
            else:
                style = 'color: #000000;'

            html += f'<td style="{style}">{value}</td>'
        html += "</tr>"

    html += "</tbody></table>"
    return html

# 將 DataFrame 轉為 CSV 並生成下載連結
def get_table_download_link(df, filename="data.csv"):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()  # 將 CSV 轉為 base64 編碼
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">下載 CSV 檔案</a>'
    return href

# 主程式
st.title("倉庫庫存查詢與缺貨提醒")

# 讀取 Google Sheet 數據
try:
    dataframes, update_dates, sorted_weeks = get_data_from_google_sheet()
except Exception as e:
    st.error(f"無法連接到 Google Sheet：{str(e)}。請檢查 Secrets 設置或 Google Sheet 存取權限。")
    st.stop()

df_week1 = dataframes["week 1"]

# 搜尋功能
st.subheader("庫存搜尋")
search_term = st.text_input("輸入搜尋關鍵字（例如產品名稱、品牌或描述）", key="inventory_search")

if search_term:
    # 在 week 1 的 Sub Group, Brand, Desc 欄中搜尋
    filtered_df = df_week1[
        df_week1["Sub Group"].str.contains(search_term, case=False, na=False) |
        df_week1["Brand"].str.contains(search_term, case=False, na=False) |
        df_week1["Desc"].str.contains(search_term, case=False, na=False)
    ]

    if not filtered_df.empty:
        # 準備結果表格
        result_df = filtered_df[["Sub Group", "Brand", "Desc", "Location", "Unit", "Quantity"]].copy()
        result_df.rename(columns={"Quantity": update_dates["week 1"]}, inplace=True)

        # 為其他週添加日期欄位
        for week in sorted_weeks:
            if week == "week 1":
                continue
            df_week = dataframes[week]
            quantities = []
            for desc in result_df["Desc"]:
                # 查找完全匹配的 Desc
                matching_row = df_week[df_week["Desc"] == desc]
                if not matching_row.empty:
                    quantities.append(matching_row["Quantity"].iloc[0])
                else:
                    quantities.append(0)  # 若無匹配，設為 0
            result_df[update_dates[week]] = quantities

        # 計算每周變化（前一週減後一週）
        for i in range(len(sorted_weeks) - 1):
            week_from = sorted_weeks[i]
            week_to = sorted_weeks[i + 1]
            date_from = update_dates[week_from]
            date_to = update_dates[week_to]
            # 將日期欄轉為數值型以計算差異
            result_df[date_from] = pd.to_numeric(result_df[date_from], errors="coerce")
            result_df[date_to] = pd.to_numeric(result_df[date_to], errors="coerce")
            # 計算變化（前一週減後一週）
            change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
            result_df[change_column_name] = result_df[date_to] - result_df[date_from]

        # 確保變化欄位也是數值型
        for col in result_df.columns:
            if "-" in col:  # 變化欄位
                result_df[col] = pd.to_numeric(result_df[col], errors="coerce")

        # 按照 Sub Group -> Brand -> Desc 排序
        result_df = result_df.sort_values(by=["Sub Group", "Brand", "Desc"])

        st.write("搜尋結果：")
        # 使用 HTML 渲染表格
        html_table = df_to_html_table(result_df, update_dates, sorted_weeks)
        st.markdown(html_table, unsafe_allow_html=True)

        # 添加交互式表格（支援放大、排序等）
        st.write("交互式表格（可排序、調整列寬）：")
        st.dataframe(result_df, use_container_width=True)

        # 添加下載按鈕
        st.markdown(get_table_download_link(result_df, "inventory_search_result.csv"), unsafe_allow_html=True)
    else:
        st.write("無符合條件的結果")
else:
    st.write("請輸入搜尋關鍵字")

# 缺貨提醒：當 week 1 的庫存量低於用量時顯示
st.subheader("缺貨提醒")

# 準備缺貨表格
low_stock_df = df_week1[["Sub Group", "Brand", "Desc", "Location", "Unit", "Quantity"]].copy()
low_stock_df.rename(columns={"Quantity": update_dates["week 1"]}, inplace=True)

# 為其他週添加日期欄位
for week in sorted_weeks:
    if week == "week 1":
        continue
    df_week = dataframes[week]
    quantities = []
    for desc in low_stock_df["Desc"]:
        matching_row = df_week[df_week["Desc"] == desc]
        if not matching_row.empty:
            quantities.append(matching_row["Quantity"].iloc[0])
        else:
            quantities.append(0)
    low_stock_df[update_dates[week]] = quantities

# 計算每周變化（前一週減後一週）
for i in range(len(sorted_weeks) - 1):
    week_from = sorted_weeks[i]
    week_to = sorted_weeks[i + 1]
    date_from = update_dates[week_from]
    date_to = update_dates[week_to]
    # 將日期欄轉為數值型以計算差異
    low_stock_df[date_from] = pd.to_numeric(low_stock_df[date_from], errors="coerce")
    low_stock_df[date_to] = pd.to_numeric(low_stock_df[date_to], errors="coerce")
    # 計算變化（前一週減後一週）
    change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
    low_stock_df[change_column_name] = low_stock_df[date_to] - low_stock_df[date_from]

# 確保變化欄位也是數值型
for col in low_stock_df.columns:
    if "-" in col:
        low_stock_df[col] = pd.to_numeric(low_stock_df[col], errors="coerce")

# 計算用量（week 2 到 week 1 的用量）
last_week = "week 2"  # 因為 sorted_weeks 現在是 ["week 1", "week 2", ..., "week 5"]
week1_date = update_dates["week 1"]
last_week_date = update_dates[last_week]
usage_column = f"{last_week_date.split('/')[0]}/{last_week_date.split('/')[1]}-{week1_date.split('/')[0]}/{week1_date.split('/')[1]}"
# 用量 = 24/3-31/3，如果 > 0 則為用量，否則為 0
low_stock_df["Last Week Usage"] = low_stock_df[usage_column].apply(lambda x: x if x > 0 else 0)

# 篩選 week 1 庫存低於用量的產品
# 處理 NaN 值，將 NaN 視為 0
low_stock_df[update_dates["week 1"]] = low_stock_df[update_dates["week 1"]].fillna(0)
low_stock_df["Last Week Usage"] = low_stock_df["Last Week Usage"].fillna(0)

# 篩選條件：week 1 庫存量 < 用量
low_stock = low_stock_df[low_stock_df[update_dates["week 1"]] < low_stock_df["Last Week Usage"]]

# 移除臨時計算欄位
low_stock = low_stock.drop(columns=["Last Week Usage"])

# 添加搜尋功能
st.write("搜尋缺貨產品：")
low_stock_search_term = st.text_input("輸入搜尋關鍵字（例如產品名稱、品牌或描述）", key="low_stock_search")

if not low_stock.empty:
    # 按照 Sub Group -> Brand -> Desc 排序
    low_stock = low_stock.sort_values(by=["Sub Group", "Brand", "Desc"])

    # 根據搜尋關鍵字過濾缺貨產品
    if low_stock_search_term:
        filtered_low_stock = low_stock[
            low_stock["Sub Group"].str.contains(low_stock_search_term, case=False, na=False) |
            low_stock["Brand"].str.contains(low_stock_search_term, case=False, na=False) |
            low_stock["Desc"].str.contains(low_stock_search_term, case=False, na=False)
        ]
        if not filtered_low_stock.empty:
            st.warning("以下產品庫存不足（低於上週用量）：")
            # 使用 HTML 渲染表格
            usage_threshold = low_stock_df["Last Week Usage"].mean()
            html_table = df_to_html_table(filtered_low_stock, update_dates, sorted_weeks, usage_threshold, is_low_stock=True)
            st.markdown(html_table, unsafe_allow_html=True)

            # 添加交互式表格（支援放大、排序等）
            st.write("交互式表格（可排序、調整列寬）：")
            st.dataframe(filtered_low_stock, use_container_width=True)

            # 添加下載按鈕
            st.markdown(get_table_download_link(filtered_low_stock, "low_stock_result.csv"), unsafe_allow_html=True)
        else:
            st.write("無符合條件的缺貨產品")
    else:
        st.warning("以下產品庫存不足（低於上週用量）：")
        # 使用 HTML 渲染表格
        usage_threshold = low_stock_df["Last Week Usage"].mean()
        html_table = df_to_html_table(low_stock, update_dates, sorted_weeks, usage_threshold, is_low_stock=True)
        st.markdown(html_table, unsafe_allow_html=True)

        # 添加交互式表格（支援放大、排序等）
        st.write("交互式表格（可排序、調整列寬）：")
        st.dataframe(low_stock, use_container_width=True)

        # 添加下載按鈕
        st.markdown(get_table_download_link(low_stock, "low_stock_result.csv"), unsafe_allow_html=True)
else:
    st.success("目前無缺貨產品！")
