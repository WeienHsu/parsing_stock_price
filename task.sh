#!/bin/bash
SETUP_PATH="/home/ubuntu/weien/parsing_stock_price"
LOG_FILE="/home/ubuntu/weien/parsing_stock_price/cron.log"

# 1. 進入路徑 (加上引號更安全)
cd "$SETUP_PATH"

# 2. 定義 Python 執行檔的路徑 (直接指向虛擬環境內部)
PYTHON_BIN="$SETUP_PATH/.venv/bin/python"

echo "------------------------------------------" >> "$LOG_FILE"
echo "[$(date '+%Y-%m-%d %H:%M:%S')] 任務啟動" >> "$LOG_FILE"

# 3. 執行程式 (直接用 PYTHON_BIN，不用 source activate)
$PYTHON_BIN parsing_and_upload_to_drive.py >> "$LOG_FILE" 2>&1

# 4. 檢查執行結果
if [ $? -eq 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] 任務完成：成功上傳 ✅" >> "$LOG_FILE"
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] 任務完成：執行失敗 ❌ (請檢查上方錯誤訊息)" >> "$LOG_FILE"
fi

# 修正了這裡的變數名稱錯誤
echo "------------------------------------------" >> "$LOG_FILE"
