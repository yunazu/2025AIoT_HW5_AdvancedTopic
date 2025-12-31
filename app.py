import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from google import genai
import io
import json

# --- 介面設定 ---
st.set_page_config(page_title="AI PPT Architect", layout="wide")
st.title("🧠 AI 簡報重構師 (NotebookLM 風格)")

# --- 側邊欄設定 ---
with st.sidebar:
    api_key = st.text_input("請輸入 Gemini API Key", type="password")
    st.info("本工具會提取原始 PPT 內容，由 AI 重新編排大綱並套用新設計。")

# --- 核心邏輯：AI 內容重寫 ---
def rewrite_content_with_ai(original_text, api_key):
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    你是一個專業的簡報設計師。以下是從一份舊簡報中提取的原始內容：
    ---
    {original_text}
    ---
    請幫我執行以下任務：
    1. 重新梳理內容，精簡為 3 頁最具代表性的投影片。
    2. 每頁內容包含：標題 (Title)、內文重點 (Bullet Points, 3條)。
    3. 為整份簡報選擇一個專業配色，並提供一個主題色的 RGB 數值 (例如: [0, 51, 102])。
    
    請嚴格按照以下 JSON 格式回傳，不要有額外文字：
    {{
      "theme_rgb": [0, 51, 102],
      "slides": [
        {{"title": "標題1", "content": ["重點1", "重點2", "重點3"]}},
        {{"title": "標題2", "content": ["重點1", "重點2", "重點3"]}},
        {{"title": "標題3", "content": ["重點1", "重點2", "重點3"]}}
      ]
    }}
    """
    response = client.models.generate_content(
                model='gemini-2.5-flash-lite', # Flash 是免費版最穩定的
                contents=prompt
            )
    return json.loads(response.text)

# --- 核心邏輯：從零生成全新 PPT ---
def create_new_ppt(ai_data):
    prs = Presentation()
    theme_rgb = RGBColor(*ai_data["theme_rgb"])

    for slide_data in ai_data["slides"]:
        # 使用標題+內容版面
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        # 設定標題
        title = slide.shapes.title
        title.text = slide_data["title"]
        title.text_frame.paragraphs[0].font.color.rgb = theme_rgb
        title.text_frame.paragraphs[0].font.bold = True

        # 設定內容
        content_box = slide.placeholders[1]
        content_box.text = "\n".join(slide_data["content"])
        
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- UI 流程 ---
uploaded_file = st.file_uploader("1. 上傳原始 PPT", type=["pptx"])

if uploaded_file and api_key:
    if st.button("🚀 開始 AI 重構並更換版型"):
        with st.spinner("AI 正在深度閱讀並重新設計中..."):
            # 1. 提取文字
            old_prs = Presentation(uploaded_file)
            full_text = ""
            for slide in old_prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        full_text += shape.text + "\n"

            # 2. AI 重新創作
            try:
                ai_result = rewrite_content_with_ai(full_text, api_key)
                
                # 3. 生成新檔案
                new_ppt = create_new_ppt(ai_result)
                
                st.success("✅ 重構完成！AI 已根據內容重新設計了版型與文案。")
                
                # 預覽 AI 的建議
                st.subheader("AI 設計大綱預覽")
                for i, s in enumerate(ai_result["slides"]):
                    st.write(f"**Slide {i+1}: {s['title']}**")

                st.download_button(
                    label="📥 下載 AI 設計的新簡報",
                    data=new_ppt,
                    file_name="AI_Redesigned_PPT.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"AI 處理過程中發生錯誤: {e}")
                st.info("請檢查 API Key 是否正確，或原始 PPT 文字是否過多。")
elif not api_key:
    st.warning("👈 請在左側輸入 Gemini API Key 以啟動 AI 功能。")