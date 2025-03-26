# 智慧居家溫度彙整器

## 功能
處理智慧居家感測器數據，彙整房間溫度報告，支援多房間、多天數據分析。

## 技術
- Python
- pandas（數據處理）
- tkinter（GUI）
- openpyxl（Excel 輸出）

## 檔案結構
- `智慧居家溫度彙整器.py`：主程式
- `generate_50_room_data.py`：測試數據生成器，可產生 50 個檔案
- `TestData/`：5 個範例數據（完整版可用生成器產生 50 個）

## 使用方法
1. 執行 `generate_50_room_data.py` 生成 50 個測試檔案。
2. 執行 `智慧居家溫度彙整器.py`，選擇 `TestData` 資料夾，生成「溫度彙整報告.xlsx」。