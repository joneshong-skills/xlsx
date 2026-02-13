[English](README.md) | [繁體中文](README.zh.md)

# XLSX 技能

一個用於建立、讀取、編輯和分析 Excel 試算表（.xlsx、.xlsm、.csv、.tsv）的 Claude Code 技能，支援專業格式、公式和財務模型。

## 功能特色

- **讀取與分析**: 使用 pandas 載入試算表進行資料分析、統計和視覺化
- **建立**: 使用 openpyxl 建立新的 Excel 檔案，包含公式、格式和圖表
- **編輯**: 修改現有活頁簿，同時保留範本、樣式和慣例
- **財務模型**: 業界標準色彩編碼、數字格式和公式建構
- **公式重算**: 透過 LibreOffice 自動重算，附帶錯誤偵測與報告
- **資料清理**: 將雜亂的表格資料（格式錯誤的列、錯位的標題）整理成正確的試算表
- **格式轉換**: 在 .xlsx、.xlsm、.csv 和 .tsv 格式之間轉換

## 安裝

```bash
# 複製到你的技能目錄
cp -r xlsx/ ~/.claude/skills/xlsx/
```

### Python 相依套件

```bash
pip install openpyxl pandas
```

### 系統相依套件 (macOS)

```bash
brew install --cask libreoffice
```

## 使用方式

直接請 Claude 處理試算表檔案：

- 「建立試算表，做 Q4 財務預測」
- 「幫這個 Excel 加一個摘要工作表和公式」
- 「清理這個雜亂的 CSV 並匯出為 xlsx」
- 「讀取並分析 sales.xlsx 的資料」
- 「幫這個試算表設定正確的數字格式」

## 檔案結構

```
xlsx/
  SKILL.md          -- 主要技能指引
  LICENSE.txt       -- 授權條款
  references/       -- （保留供參考資料使用）
  scripts/
    recalc.py       -- 透過 LibreOffice 重算公式
    office/         -- Office 文件工具（打包、解包、驗證）
      helpers/      -- XML 處理輔助程式
      schemas/      -- OOXML XSD 結構描述驗證
      validators/   -- 格式專用驗證器
  assets/           -- （保留供未來資源使用）
```

## 授權

專有授權。完整條款請見 LICENSE.txt。

## 來源

改編自 [anthropics/skills](https://github.com/anthropics/skills/tree/main/skills/xlsx)。
