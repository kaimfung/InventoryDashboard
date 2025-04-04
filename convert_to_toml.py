import json
import toml

# 讀取 credentials.json 文件
with open('credentials.json', 'r') as f:
    credentials = json.load(f)

# 創建 TOML 結構
toml_data = {
    "gcp_service_account": credentials
}

# 將數據轉換為 TOML 格式並保存
with open('secrets.toml', 'w') as f:
    toml.dump(toml_data, f)

print("轉換完成！請查看 secrets.toml 文件。")

