# Streamlit Cloud + Windows Word 服務

這個方案分成兩台機器：

1. Streamlit Community Cloud
   `監造附錄六合一產出系統app.py` 跑前端，讓同事用網址上傳 Word、選工項、下載結果。
2. Windows 電腦或 Windows Server
   `word_refresh_service.py` 跑 Word 校正服務，負責最後的頁碼、目錄頁碼區間、頁尾欄位更新。

## Cloud 端設定

在 Streamlit Community Cloud 的 app 設定頁，打開 `Secrets`，填入：

```toml
APP_PASSWORD = "replace-with-your-app-password"
WORD_REFRESH_SERVICE_URL = "https://your-word-service.example.com"
WORD_REFRESH_SERVICE_TOKEN = "replace-with-a-long-random-token"
```

也可以參考 `.streamlit/secrets.toml.example`。

如果沒設定 `WORD_REFRESH_SERVICE_URL`，前端會退回原本的本機 Word 邏輯。
如果有設定 `APP_PASSWORD`，同事開網址時會先看到登入畫面。

## Windows 端需求

- Windows
- Python
- Microsoft Word

## Windows 端啟動

先安裝依賴：

```powershell
pip install -r requirements.txt
```

設定服務環境變數：

```powershell
$env:WORD_REFRESH_SERVICE_TOKEN="replace-with-a-long-random-token"
$env:WORD_REFRESH_SERVICE_HOST="127.0.0.1"
$env:WORD_REFRESH_SERVICE_PORT="8765"
```

啟動服務：

```powershell
python word_refresh_service.py
```

健康檢查網址：

```text
http://127.0.0.1:8765/health
```

## 對外提供給 Streamlit Cloud

因為 Streamlit Cloud 要能連到 Windows 服務，所以 Windows 端還需要一個外部可存取網址。最穩的做法是：

- 用 Cloudflare Tunnel 把 `http://127.0.0.1:8765` 對外轉成 `https://...`
- 或用公司既有反向代理 / VPN / 內網穿透方案

建議：

- `WORD_REFRESH_SERVICE_TOKEN` 一定要設
- 優先走 HTTPS
- 服務盡量不要直接裸露在公網 IP

## 目前程式行為

- 有 `WORD_REFRESH_SERVICE_URL`：
  先送到 Windows 服務做 Word 校正，成功就用遠端回傳檔案。
- 遠端失敗：
  退回本機 Word 校正。
- 本機也沒有 Word：
  直接回傳未校正的 `.docx`。
