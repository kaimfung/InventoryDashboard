import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import base64

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

# 從 Google Sheet 讀取數據，並使用緩存
@st.cache_data
def get_data_from_google_sheet():
    weeks = ["week 1", "week 2", "week 3", "week 4", "week 5"]
    dataframes = {}
    update_dates = {}
    inconsistencies = []

    for week in weeks:
        try:
            worksheet = sheet.worksheet(week)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"找不到工作表 '{week}'。請確認 Google Sheet 中是否存在該工作表。")
            st.stop()
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        df.columns = df.columns.str.strip()  # 去除欄位名稱中的多餘空格
        # 檢查必要的欄位是否存在
        required_columns = ["G-Sub Group(Name)", "G-Loc/Brand(Name)", "Description", "Location Name", "Unit", "Quantity"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"工作表 '{week}' 缺少以下必要欄位：{missing_columns}")
            st.stop()
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").clip(lower=0).fillna(0)
        df = df.rename(columns={
            "G-Sub Group(Name)": "Sub Group",
            "G-Loc/Brand(Name)": "Brand",
            "Description": "Desc",
            "Location Name": "Location",
            "Unit": "Unit"
        })
        # 檢查數據一致性
        for idx, row in df.iterrows():
            if pd.isna(row["Sub Group"]) or pd.isna(row["Brand"]) or pd.isna(row["Desc"]) or pd.isna(row["Location"]) or pd.isna(row["Unit"]):
                inconsistencies.append(f"工作表 '{week}' 第 {idx+2} 行：缺少必要欄位值（Sub Group, Brand, Desc, Location, Unit）")
        dataframes[week] = df
        update_date = worksheet.acell("I1").value
        update_dates[week] = update_date

    sorted_weeks = weeks
    return dataframes, update_dates, sorted_weeks, inconsistencies

# 將 DataFrame 轉為 HTML 表格，並應用樣式
def df_to_html_table(df, update_dates, sorted_weeks, last_week_usage=None, is_low_stock=False):
    date_columns = [update_dates[week] for week in sorted_weeks]
    change_columns = [col for col in df.columns if "-" in col]

    # 將數值格式化，缺失值顯示為 "N/A"
    for col in date_columns + change_columns:
        df[col] = df[col].apply(lambda x: f"{x:.2f}" if pd.notna(x) and x != "N/A" else "N/A")

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
                if value == "N/A":
                    style = 'background-color: #ffffff; color: #000000;'
                else:
                    try:
                        value_float = float(value)
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

# 初始化 session_state 用於儲存搜尋關鍵字
if "inventory_search_term" not in st.session_state:
    st.session_state.inventory_search_term = ""
if "low_stock_search_term" not in st.session_state:
    st.session_state.low_stock_search_term = ""

# 定義回調函數，用於更新 session_state
def update_inventory_search():
    st.session_state.inventory_search_term = st.session_state.inventory_search_input

def update_low_stock_search():
    st.session_state.low_stock_search_term = st.session_state.low_stock_search_input

# 主程式
st.title("倉庫庫存查詢與缺貨提醒")

# 讀取數據
try:
    dataframes, update_dates, sortedWeeks, inconsistencies = get_data_from_google_sheet()
except Exception as e:
    st.error(f"無法從 Google Sheet 讀取數據：{str(e)}。請檢查工作表名稱或權限設置。")
    st.stop()

# 顯示數據不一致的警告
if inconsistencies:
    st.warning("發現以下數據不一致問題，請檢查 Google Sheet：")
    for inconsistency in inconsistencies:
        st.write(f"- {inconsistency}")

# 為每個工作表建立索引，加速匹配
for week in sortedWeeks:
    dataframes[week] = dataframes[week].set_index(["Sub Group", "Brand", "Desc", "Location", "Unit"], drop=False)

df_week1 = dataframes["week 1"]

# 搜尋功能
st.subheader("庫存搜尋")
st.text_input(
    "輸入搜尋關鍵字（例如產品名稱、品牌或描述）",
    value=st.session_state.inventory_search_term,
    key="inventory_search_input",
    on_change=update_inventory_search
)

# 直接使用 session_state 中的值進行搜尋
search_term = st.session_state.inventory_search_term

if search_term:
    filtered_df = df_week1[
        df_week1["Sub Group"].str.contains(search_term, case=False, na=False) |
        df_week1["Brand"].str.contains(search_term, case=False, na=False) |
        df_week1["Desc"].str.contains(search_term, case=False, na=False)
    ]

    if not filtered_df.empty:
        result_df = filtered_df[["Sub Group", "Brand", "Desc", "Location", "Unit", "Quantity"]].copy()
        result_df.rename(columns={"Quantity": update_dates["week 1"]}, inplace=True)

        # 為其他週添加數據，使用索引加速匹配
        for week in sortedWeeks:
            if week == "week 1":
                continue
            df_week = dataframes[week]
            quantities = []
            for idx, row in result_df.iterrows():
                key = (row["Sub Group"], row["Brand"], row["Desc"], row["Location"], row["Unit"])
                if key in df_week.index:
                    quantities.append(df_week.loc[key, "Quantity"])
                else:
                    quantities.append("N/A")  # 缺失數據顯示為 "N/A"
            result_df[update_dates[week]] = quantities

        # 計算每周變化
        for i in range(len(sortedWeeks) - 1):
            week_from = sortedWeeks[i]
            week_to = sortedWeeks[i + 1]
            date_from = update_dates[week_from]
            date_to = update_dates[week_to]
            # 將 "N/A" 替換為 0 以進行計算，計算後再轉回 "N/A"
            result_df[date_from] = pd.to_numeric(result_df[date_from], errors="coerce").fillna(0)
            result_df[date_to] = pd.to_numeric(result_df[date_to], errors="coerce").fillna(0)
            change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
            result_df[change_column_name] = result_df[date_to] - result_df[date_from]
            # 如果任一週為 "N/A"，則變化也設為 "N/A"
            result_df[change_column_name] = result_df.apply(
                lambda row: "N/A" if row[date_from] == 0 or row[date_to] == 0 else row[change_column_name],
                axis=1
            )

        result_df = result_df.sort_values(by=["Sub Group", "Brand", "Desc"])

        # 重新排列欄位順序
        desired_columns = ["Sub Group", "Brand", "Desc", "Location", "Unit"]
        other_columns = [col for col in result_df.columns if col not in desired_columns]
        result_df = result_df[desired_columns + other_columns]

        st.write("搜尋結果：")
        html_table = df_to_html_table(result_df, update_dates, sortedWeeks)
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

# 為其他週添加數據，使用索引加速匹配
for week in sortedWeeks:
    if week == "week 1":
        continue
    df_week = dataframes[week]
    quantities = []
    for idx, row in low_stock_df.iterrows():
        key = (row["Sub Group"], row["Brand"], row["Desc"], row["Location"], row["Unit"])
        if key in df_week.index:
            quantities.append(df_week.loc[key, "Quantity"])
        else:
            quantities.append("N/A")  # 缺失數據顯示為 "N/A"
    low_stock_df[update_dates[week]] = quantities

# 計算每周變化
for i in range(len(sortedWeeks) - 1):
    week_from = sortedWeeks[i]
    week_to = sortedWeeks[i + 1]
    date_from = update_dates[week_from]
    date_to = update_dates[week_to]
    low_stock_df[date_from] = pd.to_numeric(low_stock_df[date_from], errors="coerce").fillna(0)
    low_stock_df[date_to] = pd.to_numeric(low_stock_df[date_to], errors="coerce").fillna(0)
    change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
    low_stock_df[change_column_name] = low_stock_df[date_to] - low_stock_df[date_from]
    low_stock_df[change_column_name] = low_stock_df.apply(
        lambda row: "N/A" if row[date_from] == 0 or row[date_to] == 0 else row[change_column_name],
        axis=1
    )

# 計算用量（week 2 到 week 1 的用量）
last_week = "week 2"
week1_date = update_dates["week 1"]
last_week_date = update_dates[last_week]
usage_column = f"{last_week_date.split('/')[0]}/{last_week_date.split('/')[1]}-{week1_date.split('/')[0]}/{week1_date.split('/')[1]}"
low_stock_df["Last Week Usage"] = low_stock_df[usage_column].apply(lambda x: x if x != "N/A" and x > 0 else 0)

# 處理 NaN 值
low_stock_df[update_dates["week 1"]] = low_stock_df[update_dates["week 1"]].fillna(0)
low_stock_df["Last Week Usage"] = low_stock_df["Last Week Usage"].fillna(0)

# 按 Sub Group, Brand, Desc, Unit 合併數據，計算總庫存量
group_cols = ["Sub Group", "Brand", "Desc", "Unit"]
grouped_by_desc = low_stock_df.groupby(group_cols).agg({
    update_dates["week 1"]: "sum",
    "Last Week Usage": "sum",
}).reset_index()

# 為其他週添加數據（加總所有地點）
for week in sortedWeeks:
    if week == "week 1":
        continue
    grouped_by_desc[update_dates[week]] = low_stock_df.groupby(group_cols)[update_dates[week]].sum()

# 重新計算變化欄位（基於加總後的數據）
for i in range(len(sortedWeeks) - 1):
    week_from = sortedWeeks[i]
    week_to = sortedWeeks[i + 1]
    date_from = update_dates[week_from]
    date_to = update_dates[week_to]
    grouped_by_desc[date_from] = pd.to_numeric(grouped_by_desc[date_from], errors="coerce").fillna(0)
    grouped_by_desc[date_to] = pd.to_numeric(grouped_by_desc[date_to], errors="coerce").fillna(0)
    change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
    grouped_by_desc[change_column_name] = grouped_by_desc[date_to] - grouped_by_desc[date_from]
    grouped_by_desc[change_column_name] = grouped_by_desc.apply(
        lambda row: "N/A" if row[date_from] == 0 or row[date_to] == 0 else row[change_column_name],
        axis=1
    )

# 篩選缺貨產品：week 1 總庫存低於用量
low_stock_by_desc = grouped_by_desc[grouped_by_desc[update_dates["week 1"]] < grouped_by_desc["Last Week Usage"]]

# 準備最終顯示的表格，重新加入 Location 欄
if not low_stock_by_desc.empty:
    # 合併所有地點的 Location 資訊
    location_dict = low_stock_df.groupby(group_cols)["Location"].apply(
        lambda x: ", ".join(sorted(set(x)))
    ).to_dict()

    # 將 Location 資訊映射回 low_stock_by_desc
    low_stock_by_desc["Location"] = low_stock_by_desc.apply(
        lambda row: location_dict.get((row["Sub Group"], row["Brand"], row["Desc"], row["Unit"]), ""),
        axis=1
    )

    # 移除 Last Week Usage 欄位
    low_stock_by_desc = low_stock_by_desc.drop(columns=["Last Week Usage"])

    # 重新排列欄位順序
    desired_columns = ["Sub Group", "Brand", "Desc", "Location", "Unit"]
    other_columns = [col for col in low_stock_by_desc.columns if col not in desired_columns]
    low_stock_by_desc = low_stock_by_desc[desired_columns + other_columns]

    # 添加搜尋功能
    st.write("搜尋缺貨產品：")
    st.text_input(
        "輸入搜尋關鍵字（例如產品名稱、品牌或描述）",
        value=st.session_state.low_stock_search_term,
        key="low_stock_search_input",
        on_change=update_low_stock_search
    )

    # 直接使用 session_state 中的值進行搜尋
    low_stock_search_term = st.session_state.low_stock_search_term

    low_stock_by_desc = low_stock_by_desc.sort_values(by=["Sub Group", "Brand", "Desc"])

    if low_stock_search_term:
        filtered_low_stock = low_stock_by_desc[
            low_stock_by_desc["Sub Group"].str.contains(low_stock_search_term, case=False, na=False) |
            low_stock_by_desc["Brand"].str.contains(low_stock_search_term, case=False, na=False) |
            low_stock_by_desc["Desc"].str.contains(low_stock_search_term, case=False, na=False)
        ]
        if not filtered_low_stock.empty:
            st.warning("以下產品庫存不足（低於上週用量）：")
            html_table = df_to_html_table(
                filtered_low_stock,
                update_dates,
                sortedWeeks,
                last_week_usage=grouped_by_desc["Last Week Usage"].reindex(filtered_low_stock.index),
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
            low_stock_by_desc,
            update_dates,
            sortedWeeks,
            last_week_usage=grouped_by_desc["Last Week Usage"].reindex(low_stock_by_desc.index),
            is_low_stock=True
        )
        st.markdown(html_table, unsafe_allow_html=True)

        st.write("交互式表格（可排序、調整列寬）：")
        st.dataframe(low_stock_by_desc, use_container_width=True)

        st.markdown(get_table_download_link(low_stock_by_desc, "low_stock_result.csv"), unsafe_allow_html=True)
else:
    st.success("目前無缺貨產品！")
