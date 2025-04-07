import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import base64
import time

# 設定 Google Sheet 認證
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

try:
    # 從 Streamlit Secrets 讀取認證資訊
    creds_dict = st.secrets["gcp_service_account"]

    # 檢查必要的鍵是否存在
    required_keys = ["type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url", "client_x509_cert_url"]
    missing_keys = [key for key in required_keys if key not in creds_dict]
    if missing_keys:
        st.error(f"Streamlit Secrets 中缺少以下必要的鍵：{missing_keys}")
        st.stop()

    # 使用字典創建認證物件
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)

except Exception as e:
    st.error(f"無法初始化 Google Sheet 認證：{str(e)}。請檢查 Streamlit Secrets 設置。")
    st.stop()

# 連接到 Google Sheet
try:
    sheet = client.open("INVENTORY_API")
except gspread.exceptions.SpreadsheetNotFound:
    st.error("找不到 Google Sheet 'INVENTORY_API'。請確認名稱是否正確，並確保已與服務帳戶共用。")
    st.stop()
except Exception as e:
    st.error(f"無法連接到 Google Sheet：{str(e)}。請檢查 Google Sheet 存取權限或 Secrets 設置。")
    st.stop()

# 定義 Location 簡寫映射
LOCATION_ABBREVIATIONS = {
    "Direct Shipment Warehouse": "DSW",
    "TIN WAN": "TW",
    "KERRY 1": "K1",
    "Hong Kong Ice": "HKIce",
    "Macau": "澳",
}

# 從 Google Sheet 讀取數據
def get_data_from_google_sheet():
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
        df.columns = df.columns.str.strip()
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").clip(lower=0).fillna(0)
        df = df.rename(columns={
            "G-Sub Group(Name)": "Sub Group",
            "G-Loc/Brand(Name)": "Brand",
            "Description": "Desc",
            "Location Name": "Location",
            "Unit": "Unit"
        })
        dataframes[week] = df
        update_date = worksheet.acell("I1").value
        update_dates[week] = update_date

    sorted_weeks = weeks
    return dataframes, update_dates, sorted_weeks

# 將 DataFrame 轉為 HTML 表格，並應用樣式
def df_to_html_table(df, update_dates, sorted_weeks, last_week_usage=None, is_low_stock=False):
    date_columns = [update_dates[week] for week in sorted_weeks]
    change_columns = [col for col in df.columns if "-" in col]

    for col in date_columns + change_columns:
        df[col] = df[col].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")

    html = """
    <style>
        .table-container {
            width: 100%;
            overflow-x: scroll;
            max-height: 500px;
            overflow-y: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
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
    <div class="table-container">
    <table>
        <thead>
            <tr>
    """

    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"

    for idx, (_, row) in enumerate(df.iterrows()):
        html += "<tr>"
        for col in df.columns:
            value = row[col]
            style = ""

            if col in date_columns:
                style = 'background-color: #e6f3ff; color: #000000;'
                if is_low_stock and col == update_dates["week 1"]:
                    week1_stock = float(row[update_dates["week 1"]]) if row[update_dates["week 1"]] != "N/A" else float('inf')
                    usage_threshold = last_week_usage.iloc[idx] if last_week_usage is not None else float('inf')
                    if pd.notna(usage_threshold) and week1_stock < usage_threshold:
                        style = 'background-color: #ffcccc; color: #000000;'
            elif col in change_columns:
                try:
                    value_float = float(value) if value != "N/A" else 0
                    display_value = f"-{abs(value_float):.2f}" if value_float > 0 else f"{abs(value_float):.2f}"
                    if value_float > 0:
                        style = 'background-color: #ffcccc; color: #000000;'
                    elif value_float < 0:
                        style = 'background-color: #ccffcc; color: #000000;'
                    else:
                        style = 'background-color: #ffffff; color: #000000;'
                    value = display_value
                except:
                    style = 'background-color: #ffffff; color: #000000;'
            else:
                style = 'color: #000000;'

            html += f'<td style="{style}">{value}</td>'
        html += "</tr>"

    html += "</tbody></table></div>"
    return html

# 將 DataFrame 轉為 CSV 並生成下載連結
def get_table_download_link(df, filename="data.csv"):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">下載 CSV 檔案</a>'
    return href

# 初始化 session_state 用於儲存搜尋關鍵字和時間戳
if "inventory_search_term" not in st.session_state:
    st.session_state.inventory_search_term = ""
if "low_stock_search_term" not in st.session_state:
    st.session_state.low_stock_search_term = ""
if "inventory_search_timestamp" not in st.session_state:
    st.session_state.inventory_search_timestamp = 0
if "low_stock_search_timestamp" not in st.session_state:
    st.session_state.low_stock_search_timestamp = 0
if "inventory_search_display_term" not in st.session_state:
    st.session_state.inventory_search_display_term = ""
if "low_stock_search_display_term" not in st.session_state:
    st.session_state.low_stock_search_display_term = ""

# 定義防抖延遲時間（單位：秒）
DEBOUNCE_DELAY = 0.3  # 300 毫秒

# 定義回調函數，用於更新 session_state
def update_inventory_search():
    st.session_state.inventory_search_term = st.session_state.inventory_search_input
    st.session_state.inventory_search_timestamp = time.time()

def update_low_stock_search():
    st.session_state.low_stock_search_term = st.session_state.low_stock_search_input
    st.session_state.low_stock_search_timestamp = time.time()

# 主程式
st.title("倉庫庫存查詢與缺貨提醒")

try:
    dataframes, update_dates, sorted_weeks = get_data_from_google_sheet()
except Exception as e:
    st.error(f"無法從 Google Sheet 讀取數據：{str(e)}。請檢查工作表名稱或權限設置。")
    st.stop()

df_week1 = dataframes["week 1"]

# 搜尋功能
st.subheader("庫存搜尋")
st.text_input(
    "輸入搜尋關鍵字（例如產品名稱、品牌或描述）",
    value=st.session_state.inventory_search_term,
    key="inventory_search_input",
    on_change=update_inventory_search
)

# 檢查是否需要更新搜尋結果（防抖機制）
current_time = time.time()
if (current_time - st.session_state.inventory_search_timestamp) >= DEBOUNCE_DELAY:
    st.session_state.inventory_search_display_term = st.session_state.inventory_search_term

# 使用防抖後的搜尋關鍵字進行搜尋
search_term = st.session_state.inventory_search_display_term

if search_term:
    filtered_df = df_week1[
        df_week1["Sub Group"].str.contains(search_term, case=False, na=False) |
        df_week1["Brand"].str.contains(search_term, case=False, na=False) |
        df_week1["Desc"].str.contains(search_term, case=False, na=False)
    ]

    if not filtered_df.empty:
        result_df = filtered_df[["Sub Group", "Brand", "Desc", "Location", "Unit", "Quantity"]].copy()
        result_df.rename(columns={"Quantity": update_dates["week 1"]}, inplace=True)

        for week in sorted_weeks:
            if week == "week 1":
                continue
            df_week = dataframes[week]
            quantities = []
            for desc in result_df["Desc"]:
                matching_row = df_week[df_week["Desc"] == desc]
                if not matching_row.empty:
                    quantities.append(matching_row["Quantity"].iloc[0])
                else:
                    quantities.append(0)
            result_df[update_dates[week]] = quantities

        for i in range(len(sorted_weeks) - 1):
            week_from = sorted_weeks[i]
            week_to = sorted_weeks[i + 1]
            date_from = update_dates[week_from]
            date_to = update_dates[week_to]
            result_df[date_from] = pd.to_numeric(result_df[date_from], errors="coerce")
            result_df[date_to] = pd.to_numeric(result_df[date_to], errors="coerce")
            change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
            result_df[change_column_name] = result_df[date_to] - result_df[date_from]

        for col in result_df.columns:
            if "-" in col:
                result_df[col] = pd.to_numeric(result_df[col], errors="coerce")

        result_df = result_df.sort_values(by=["Sub Group", "Brand", "Desc"])

        st.write("搜尋結果：")
        html_table = df_to_html_table(result_df, update_dates, sorted_weeks)
        st.markdown(html_table, unsafe_allow_html=True)

        st.write("交互式表格（可排序、調整列寬）：")
        st.dataframe(result_df, use_container_width=True)

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
    locations = []
    for desc in low_stock_df["Desc"]:
        matching_rows = df_week[df_week["Desc"] == desc]
        if not matching_rows.empty:
            qty = matching_rows["Quantity"].sum()
            quantities.append(qty)
            locs = matching_rows["Location"].tolist()
            locations.append(locs)
        else:
            quantities.append(0)
            locations.append([])
    low_stock_df[update_dates[week]] = quantities
    low_stock_df[f"{week}_locations"] = locations

# 計算每周變化（前一週減後一週）
for i in range(len(sorted_weeks) - 1):
    week_from = sorted_weeks[i]
    week_to = sorted_weeks[i + 1]
    date_from = update_dates[week_from]
    date_to = update_dates[week_to]
    low_stock_df[date_from] = pd.to_numeric(low_stock_df[date_from], errors="coerce")
    low_stock_df[date_to] = pd.to_numeric(low_stock_df[date_to], errors="coerce")
    change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
    low_stock_df[change_column_name] = low_stock_df[date_to] - low_stock_df[date_from]

# 確保變化欄位也是數值型
for col in low_stock_df.columns:
    if "-" in col:
        low_stock_df[col] = pd.to_numeric(low_stock_df[col], errors="coerce")

# 計算用量（week 2 到 week 1 的用量）
last_week = "week 2"
week1_date = update_dates["week 1"]
last_week_date = update_dates[last_week]
usage_column = f"{last_week_date.split('/')[0]}/{last_week_date.split('/')[1]}-{week1_date.split('/')[0]}/{week1_date.split('/')[1]}"
low_stock_df["Last Week Usage"] = low_stock_df[usage_column].apply(lambda x: x if x > 0 else 0)

# 處理 NaN 值
low_stock_df[update_dates["week 1"]] = low_stock_df[update_dates["week 1"]].fillna(0)
low_stock_df["Last Week Usage"] = low_stock_df["Last Week Usage"].fillna(0)

# 按 Description 合併數據
grouped = low_stock_df.groupby("Desc").agg({
    "Sub Group": "first",
    "Brand": "first",
    "Unit": "first",
    update_dates["week 1"]: "sum",
    "Last Week Usage": "sum",
})

for week in sorted_weeks:
    if week == "week 1":
        continue
    grouped[update_dates[week]] = low_stock_df.groupby("Desc")[update_dates[week]].sum()

# 合併 Location 欄並轉換為簡寫
def combine_locations(group):
    all_locations = set()
    for week in sorted_weeks:
        if week == "week 1":
            locs = group["Location"]
            all_locations.update(locs)
        else:
            locs = group[f"{week}_locations"].explode()
            all_locations.update(locs)
    all_locations.discard("")
    abbreviated_locs = [LOCATION_ABBREVIATIONS.get(loc, loc) for loc in all_locations if loc]
    return ", ".join(sorted(abbreviated_locs))

grouped["Location"] = low_stock_df.groupby("Desc").apply(combine_locations)

# 重新計算變化欄位
for i in range(len(sorted_weeks) - 1):
    week_from = sorted_weeks[i]
    week_to = sorted_weeks[i + 1]
    date_from = update_dates[week_from]
    date_to = update_dates[week_to]
    change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
    grouped[change_column_name] = grouped[date_to] - grouped[date_from]

grouped = grouped.reset_index()

# 篩選缺貨產品：week 1 庫存低於用量
low_stock = grouped[grouped[update_dates["week 1"]] < grouped["Last Week Usage"]]

# 移除 Last Week Usage 欄位
low_stock = low_stock.drop(columns=["Last Week Usage"])

# 添加搜尋功能
st.write("搜尋缺貨產品：")
st.text_input(
    "輸入搜尋關鍵字（例如產品名稱、品牌或描述）",
    value=st.session_state.low_stock_search_term,
    key="low_stock_search_input",
    on_change=update_low_stock_search
)

# 檢查是否需要更新搜尋結果（防抖機制）
if (current_time - st.session_state.low_stock_search_timestamp) >= DEBOUNCE_DELAY:
    st.session_state.low_stock_search_display_term = st.session_state.low_stock_search_term

# 使用防抖後的搜尋關鍵字進行搜尋
low_stock_search_term = st.session_state.low_stock_search_display_term

if not low_stock.empty:
    low_stock = low_stock.sort_values(by=["Sub Group", "Brand", "Desc"])

    if low_stock_search_term:
        filtered_low_stock = low_stock[
            low_stock["Sub Group"].str.contains(low_stock_search_term, case=False, na=False) |
            low_stock["Brand"].str.contains(low_stock_search_term, case=False, na=False) |
            low_stock["Desc"].str.contains(low_stock_search_term, case=False, na=False)
        ]
        if not filtered_low_stock.empty:
            st.warning("以下產品庫存不足（低於上週用量）：")
            html_table = df_to_html_table(
                filtered_low_stock,
                update_dates,
                sorted_weeks,
                last_week_usage=grouped["Last Week Usage"].reindex(filtered_low_stock.index),
                is_low_stock=True
            )
            st.markdown(html_table, unsafe_allow_html=True)

            st.write("交互式表格（可排序、調整列寬）：")
            st.dataframe(filtered_low_stock, use_container_width=True)

            st.markdown(get_table_download_link(filtered_low_stock, "low_stock_result.csv"), unsafe_allow_html=True)
        else:
            st.write("無符合條件的缺貨產品")
    else:
        st.warning("以下產品庫存不足（低於上週用量）：")
        html_table = df_to_html_table(
            low_stock,
            update_dates,
            sorted_weeks,
            last_week_usage=grouped["Last Week Usage"].reindex(low_stock.index),
            is_low_stock=True
        )
        st.markdown(html_table, unsafe_allow_html=True)

        st.write("交互式表格（可排序、調整列寬）：")
        st.dataframe(low_stock, use_container_width=True)

        st.markdown(get_table_download_link(low_stock, "low_stock_result.csv"), unsafe_allow_html=True)
else:
    st.success("目前無缺貨產品！")
