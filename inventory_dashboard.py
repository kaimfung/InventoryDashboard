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

# 主程式
st.title("倉庫庫存查詢與缺貨提醒")

# 讀取數據
try:
    dataframes, update_dates, sorted_weeks, inconsistencies = get_data_from_google_sheet()
except Exception as e:
    st.error(f"無法從 Google Sheet 讀取數據：{str(e)}。請檢查工作表名稱或權限設置。")
    st.stop()

# 顯示數據不一致的警告
if inconsistencies:
    st.warning("發現以下數據不一致問題，請檢查 Google Sheet：")
    for inconsistency in inconsistencies:
        st.write(f"- {inconsistency}")

df_week1 = dataframes["week 1"]

# 搜尋功能
st.subheader("庫存搜尋")
st.write("提示：輸入關鍵字後按 Enter 鍵以搜尋")
with st.form(key="inventory_search_form"):
    search_term = st.text_input(
        "輸入搜尋關鍵字（例如產品名稱、品牌或描述）",
        key="inventory_search_input"
    )
    submit_button = st.form_submit_button(label="搜尋")

if search_term or submit_button:
    filtered_df = df_week1[
        df_week1["Sub Group"].str.contains(search_term, case=False, na=False) |
        df_week1["Brand"].str.contains(search_term, case=False, na=False) |
        df_week1["Desc"].str.contains(search_term, case=False, na=False)
    ]

    if not filtered_df.empty:
        result_df = filtered_df[["Sub Group", "Brand", "Desc", "Location", "Unit", "Quantity"]].copy()
        result_df.rename(columns={"Quantity": update_dates["week 1"]}, inplace=True)

        # 為其他週添加數據，確保所有欄位匹配
        for week in sorted_weeks:
            if week == "week 1":
                continue
            df_week = dataframes[week]
            quantities = []
            for _, row in result_df.iterrows():
                matching_row = df_week[
                    (df_week["Sub Group"] == row["Sub Group"]) &
                    (df_week["Brand"] == row["Brand"]) &
                    (df_week["Desc"] == row["Desc"]) &
                    (df_week["Location"] == row["Location"]) &
                    (df_week["Unit"] == row["Unit"])
                ]
                if not matching_row.empty:
                    quantities.append(matching_row["Quantity"].iloc[0])
                else:
                    quantities.append("N/A")  # 改為 "N/A" 以保持一致
            result_df[update_dates[week]] = quantities

        # 計算每周變化
        for i in range(len(sorted_weeks) - 1):
            week_from = sorted_weeks[i]
            week_to = sorted_weeks[i + 1]
            date_from = update_dates[week_from]
            date_to = update_dates[week_to]
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
        html_table = df_to_html_table(result_df, update_dates, sorted_weeks)
        st.markdown(html_table, unsafe_allow_html=True)

        st.write("交互式表格（可排序、調整列寬）：")
        st.dataframe(result_df, use_container_width=True)

        st.markdown(get_table_download_link(result_df, "inventory_search_result.csv"), unsafe_allow_html=True)
    else:
        st.write("無符合條件的結果")
else:
    st.write("請輸入搜尋關鍵字")

# 缺貨提醒：先按 Desc 加總所有 Location 的庫存，再判斷是否缺貨
st.subheader("缺貨提醒")

# 準備缺貨表格，按 Desc 加總所有 Location
group_cols_total = ["Sub Group", "Brand", "Desc", "Unit"]  # 不包含 Location
low_stock_df = df_week1.groupby(group_cols_total).agg({
    "Quantity": "sum"
}).reset_index()
low_stock_df.rename(columns={"Quantity": update_dates["week 1"]}, inplace=True)

# 為其他週添加日期欄位
for week in sorted_weeks:
    if week == "week 1":
        continue
    df_week = dataframes[week]
    quantities = []
    for _, row in low_stock_df.iterrows():
        matching_rows = df_week[
            (df_week["Sub Group"] == row["Sub Group"]) &
            (df_week["Brand"] == row["Brand"]) &
            (df_week["Desc"] == row["Desc"]) &
            (df_week["Unit"] == row["Unit"])
        ]
        if not matching_rows.empty:
            qty = matching_rows["Quantity"].sum()
            quantities.append(qty)
        else:
            quantities.append(0)
    low_stock_df[update_dates[week]] = quantities

# 計算每周變化（前一週減後一週）
for i in range(len(sorted_weeks) - 1):
    week_from = sorted_weeks[i]
    week_to = sorted_weeks[i + 1]
    date_from = update_dates[week_from]
    date_to = update_dates[week_to]
    low_stock_df[date_from] = pd.to_numeric(low_stock_df[date_from], errors="coerce").fillna(0)
    low_stock_df[date_to] = pd.to_numeric(low_stock_df[date_to], errors="coerce").fillna(0)
    change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
    low_stock_df[change_column_name] = low_stock_df[date_to] - low_stock_df[date_from]
    # 如果任一週為 0，則變化也設為 "N/A"
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

# 處理 NaN 值並確保數值型
low_stock_df[update_dates["week 1"]] = pd.to_numeric(low_stock_df[update_dates["week 1"]], errors="coerce").fillna(0)
low_stock_df["Last Week Usage"] = pd.to_numeric(low_stock_df["Last Week Usage"], errors="coerce").fillna(0)

# 檢查必要的欄位是否存在
if update_dates["week 1"] not in low_stock_df.columns:
    st.error(f"low_stock_df 中缺少欄位：{update_dates['week 1']}。請檢查 Google Sheet 中的日期格式或數據。")
    st.stop()
if "Last Week Usage" not in low_stock_df.columns:
    st.error("low_stock_df 中缺少欄位：Last Week Usage。請檢查程式碼邏輯。")
    st.stop()

# 篩選缺貨產品：week 1 庫存低於用量（加總後）
low_stock_total = low_stock_df[low_stock_df[update_dates["week 1"]] < low_stock_df["Last Week Usage"]]

# 如果有缺貨產品，展開顯示每個 Location 的詳細庫存
if not low_stock_total.empty:
    # 準備顯示詳細庫存的表格
    detailed_low_stock = pd.DataFrame()
    for _, row in low_stock_total.iterrows():
        # 從 week 1 數據中獲取該產品的所有 Location 詳細資料
        matching_rows = df_week1[
            (df_week1["Sub Group"] == row["Sub Group"]) &
            (df_week1["Brand"] == row["Brand"]) &
            (df_week1["Desc"] == row["Desc"]) &
            (df_week1["Unit"] == row["Unit"])
        ][["Sub Group", "Brand", "Desc", "Location", "Unit", "Quantity"]]
        matching_rows.rename(columns={"Quantity": update_dates["week 1"]}, inplace=True)

        # 為其他週添加數據
        for week in sorted_weeks:
            if week == "week 1":
                continue
            df_week = dataframes[week]
            quantities = []
            for _, match_row in matching_rows.iterrows():
                match = df_week[
                    (df_week["Sub Group"] == match_row["Sub Group"]) &
                    (df_week["Brand"] == match_row["Brand"]) &
                    (df_week["Desc"] == match_row["Desc"]) &
                    (df_week["Location"] == match_row["Location"]) &
                    (df_week["Unit"] == match_row["Unit"])
                ]
                if not match.empty:
                    quantities.append(match["Quantity"].iloc[0])
                else:
                    quantities.append("N/A")  # 改為 "N/A" 以保持一致
            matching_rows[update_dates[week]] = quantities

        # 計算每周變化
        for i in range(len(sorted_weeks) - 1):
            week_from = sorted_weeks[i]
            week_to = sorted_weeks[i + 1]
            date_from = update_dates[week_from]
            date_to = update_dates[week_to]
            matching_rows[date_from] = pd.to_numeric(matching_rows[date_from], errors="coerce").fillna(0)
            matching_rows[date_to] = pd.to_numeric(matching_rows[date_to], errors="coerce").fillna(0)
            change_column_name = f"{date_to.split('/')[0]}/{date_to.split('/')[1]}-{date_from.split('/')[0]}/{date_from.split('/')[1]}"
            matching_rows[change_column_name] = matching_rows[date_to] - matching_rows[date_from]
            # 如果任一週為 "N/A"，則變化也設為 "N/A"
            matching_rows[change_column_name] = matching_rows.apply(
                lambda row: "N/A" if row[date_from] == 0 or row[date_to] == 0 else row[change_column_name],
                axis=1
            )

        detailed_low_stock = pd.concat([detailed_low_stock, matching_rows], ignore_index=True)

    # 重新排列欄位順序
    desired_columns = ["Sub Group", "Brand", "Desc", "Location", "Unit"]
    other_columns = [col for col in detailed_low_stock.columns if col not in desired_columns]
    detailed_low_stock = detailed_low_stock[desired_columns + other_columns]

    # 添加搜尋功能
    st.write("搜尋缺貨產品：")
    st.write("提示：輸入關鍵字後按 Enter 鍵以搜尋")
    with st.form(key="low_stock_search_form"):
        low_stock_search_term = st.text_input(
            "輸入搜尋關鍵字（例如產品名稱、品牌或描述）",
            key="low_stock_search_input"
        )
        low_stock_submit_button = st.form_submit_button(label="搜尋")

    detailed_low_stock = detailed_low_stock.sort_values(by=["Sub Group", "Brand", "Desc", "Location"])

    if low_stock_search_term or low_stock_submit_button:
        filtered_low_stock = detailed_low_stock[
            detailed_low_stock["Sub Group"].str.contains(low_stock_search_term, case=False, na=False) |
            detailed_low_stock["Brand"].str.contains(low_stock_search_term, case=False, na=False) |
            detailed_low_stock["Desc"].str.contains(low_stock_search_term, case=False, na=False)
        ]
        if not filtered_low_stock.empty:
            st.warning("以下產品庫存不足（低於上週用量，加總所有地點後）：")
            html_table = df_to_html_table(
                filtered_low_stock,
                update_dates,
                sorted_weeks,
                last_week_usage=None,  # 因為已經按加總後判斷，這裡不再需要單獨的用量比較
                is_low_stock=True
            )
            st.markdown(html_table, unsafe_allow_html=True)

            st.write("交互式表格（可排序、調整列寬）：")
            st.dataframe(filtered_low_stock, use_container_width=True)

            st.markdown(get_table_download_link(filtered_low_stock, "low_stock_result.csv"), unsafe_allow_html=True)
        else:
            st.write("無符合條件的缺貨產品")
    else:
        st.warning("以下產品庫存不足（低於上週用量，加總所有地點後）：")
        html_table = df_to_html_table(
            detailed_low_stock,
            update_dates,
            sorted_weeks,
            last_week_usage=None,
            is_low_stock=True
        )
        st.markdown(html_table, unsafe_allow_html=True)

        st.write("交互式表格（可排序、調整列寬）：")
        st.dataframe(detailed_low_stock, use_container_width=True)

        st.markdown(get_table_download_link(detailed_low_stock, "low_stock_result.csv"), unsafe_allow_html=True)
else:
    st.success("目前無缺貨產品（加總所有地點後）！")
