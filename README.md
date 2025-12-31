# 📂 AI PPT Architect Pro — 智能簡報重構師

本專案為 **HW5 — Advanced Topic (Q3)** 的實作成果。這是一個結合 **Gemini AI** 語義理解與 **Python 自動化技術** 的簡報工具，旨在解決傳統 PPT 換版型時「美感不足」與「內容冗餘」的痛點。

## 🌟 專案介紹
不同於單純修改底色的工具，本專案採用 **「解耦內容與形式 (Decoupling Content and Style)」** 的核心邏輯：
* **🧠 AI 內容重構**：利用 Google Gemini API 深度閱讀原始 PPT，將混亂的內容重新摘要成邏輯嚴謹的 3-5 頁大綱。
* **🎨 模板樣式映射**：將 AI 重寫後的內容，精準注入使用者上傳的專業設計模板（Template），完美繼承母片中的字體、顏色與排版配置。

---

## 🚀 使用方法

### 1. 環境準備
請確保你的電腦已安裝 Python 3.9+，並安裝必要套件：
```bash
pip install streamlit python-pptx google-generativeai
```
### 2. 取得 Gemini API Key
前往 Google AI Studio 免費申請 API Key。

### 3. 執行程式
在終端機輸入以下指令啟動 Streamlit 服務：

```Bash
streamlit run app.py
```
### 4. 操作步驟
1. 輸入金鑰：在側邊欄輸入你的 Gemini API Key。

2. 上傳來源：上傳一份你想「改寫」的原始 PPT 檔案。

3. 上傳外觀：上傳一份「空的」專業 PPT 模板檔案（建議包含精美母片設計）。

4. 開始轉換：點擊「🚀 開始 AI 重構」，等待 AI 處理完成後，下載重塑後的 PPT。

### 🛠 技術細節與挑戰克服
* JSON 容錯機制：針對 LLM 回傳格式不穩定的問題，開發了 Markdown 清理邏輯，確保 JSON 解析 100% 成功。

* 佔位符自動匹配 (Placeholder Mapping)：透過 slide.placeholders 定位技術，讓內容自動對齊模板中的標題框與內容框，解決了 python-pptx 直接寫入導致排版跑掉的問題。

* 內容二次創作：程式不只是複製文字，而是透過 Prompt Engineering 讓 AI 扮演簡報設計師的角色，提升內容的含金量。