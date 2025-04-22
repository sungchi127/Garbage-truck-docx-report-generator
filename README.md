# Garbage-truck-docx-report-generator

一款基於 Python、Tkinter 與 PaddleOCR 的桌面應用，  
可即時對圖片進行 OCR 辨識，並依照車牌／輪胎規格對照表生成格式化的 Word 報告（.docx）。

---

## 功能

- 圖片文字辨識（OCR）  
- 實時預覽所選圖片  
- 可滾動檢視辨識結果  
- 根據「車牌對照表」與「輪胎規格表」，自動填入 Word 模板  
- 支援多種模板：`template_yellow.docx`、`template_white.docx`  
- 輸出報告至 `output/` 資料夾  

---

## 環境需求

- Python 3.7 以上  
- Windows / macOS / Linux  

---

## 安裝步驟

1. 準備外部檔案  
   - `license_mapping/車牌對照表、輪胎規格表114.03.03.xlsx`  
   - `templates/template_yellow.docx`  
   - `templates/template_white.docx`  

   請確保上述檔案路徑與專案結構相符，否則無法正確載入。
   
---

## 執行方式

- **開發版**  
  ```bash
  python main.py
  ```

- **已打包執行檔**  
  ```bash
  python main-pack.py
  ```
![image](https://github.com/user-attachments/assets/52d81163-42e8-4412-a9c2-77794036df61)

1. 點擊「選擇圖片」按鈕並挑選欲辨識之圖片  
2. 左側顯示圖片預覽，並自動觸發 OCR  
3. 右側文字框可滾動檢視辨識結果  
4. 點擊「產生報告」按鈕，.docx 檔案會輸出至 `output/`  

---

## 相依套件
text
paddleocr
pillow
python-docx
pandas
