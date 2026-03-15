# 冷卻水塔水蒸發量計算與分析系統

## 功能概述

本系統提供完整的冷卻水塔蒸發量計算、分析與報告匯出功能。

### 計算功能
- **簡易經驗公式**：E = C × ΔT × Q / 600，快速估算
- **Merkel 焓差法**：考慮氣水比、填料特性，精確計算
- **濕空氣性質計算**：飽和壓力、比濕度、焓值、相對濕度
- **水損失明細**：蒸發、飛濺、排污、總補給水量

### 圖表分析
- 水損失佔比圓餅圖
- 冷卻曲線（焓值 vs 溫度）
- COC vs 補給水量趨勢圖
- 溫度-焓值關係圖

### 匯出功能
- **PDF 分析報告**：封面、表格、圖表、結論建議（reportlab）
- **Word 報告文件**：目錄結構、公式說明、附錄（python-docx）
- **Excel 計算表**：可編輯公式、敏感度分析矩陣（openpyxl）

## 安裝與執行

```bash
# 1. 安裝相依套件
pip install -r requirements.txt

# 2. 啟動應用
streamlit run app.py
```

瀏覽器會自動開啟 http://localhost:8501

## 檔案結構

```
cooling_tower/
├── app.py              # 主程式（Streamlit 介面）
├── calculations.py     # 核心計算模組
├── export_pdf.py       # PDF 匯出模組
├── export_docx.py      # Word 匯出模組
├── export_xlsx.py      # Excel 匯出模組
├── requirements.txt    # Python 相依套件
└── README.md           # 說明文件
```

## 輸入參數說明

| 參數 | 說明 | 預設值 | 單位 |
|------|------|--------|------|
| Q | 循環水量 | 500 | m³/h |
| T_in | 進水溫度 | 37 | °C |
| T_out | 出水溫度 | 32 | °C |
| Tdb | 環境乾球溫度 | 35 | °C |
| Twb | 環境濕球溫度 | 28 | °C |
| G | 空氣流量 | 400,000 | m³/h |
| KaV/L | 填料特性係數 | 1.5 | - |
| COC | 濃縮倍數 | 3.0 | - |

## 計算公式參考

### 簡易經驗公式
```
E = C × ΔT × Q / 600
```
- C：蒸發係數（夏季 1.0，冬季 0.8）
- ΔT：進出水溫差（°C）
- Q：循環水量（m³/h）

### Merkel 焓差法
```
NTU = ∫(Cp·dT) / (h_s - h_a)
```
使用 Simpson 數值積分計算。

### 補給水量
```
Total = Evaporation + Splash + Blowdown
Blowdown = Evaporation / (COC - 1)
```
