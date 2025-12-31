import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io

st.set_page_config(page_title="PPT AI Redesigner", page_icon="🎨")
st.title("🔄 PPT 智能換版型工具")
st.write("上傳一份原始 PPT，由 AI 自動轉換為兩種不同設計風格。")

# --- 核心功能：讀取原始 PPT 內容 ---
def extract_text_from_ppt(uploaded_file):
    prs = Presentation(uploaded_file)
    content_list = []
    for slide in prs.slides:
        slide_data = {"title": "", "text": ""}
        if slide.shapes.title:
            slide_data["title"] = slide.shapes.title.text
        
        # 抓取非標題的文字方塊內容
        other_texts = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape != slide.shapes.title:
                other_texts.append(shape.text)
        slide_data["text"] = "\n".join(other_texts)
        content_list.append(slide_data)
    return content_list

# --- 核心功能：生成新風格 PPT ---
def redesign_ppt(original_content, style="business"):
    new_prs = Presentation()
    
    # 設定風格參數
    bg_color = RGBColor(255, 255, 255) if style == "business" else RGBColor(30, 30, 30)
    title_color = RGBColor(0, 80, 150) if style == "business" else RGBColor(0, 255, 200)
    text_color = RGBColor(50, 50, 50) if style == "business" else RGBColor(220, 220, 220)
    alignment = PP_ALIGN.LEFT if style == "business" else PP_ALIGN.CENTER

    for data in original_content:
        slide_layout = new_prs.slide_layouts[1] # 標題+內容
        slide = new_prs.slides.add_slide(slide_layout)
        
        # 1. 背景設定
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = bg_color

        # 2. 標題重新設計
        if slide.shapes.title:
            title_shape = slide.shapes.title
            title_shape.text = data["title"]
            para = title_shape.text_frame.paragraphs[0]
            para.font.bold = True
            para.font.color.rgb = title_color
            para.alignment = alignment

        # 3. 內文重新設計
        content_shape = slide.placeholders[1]
        content_shape.text = data["text"]
        for p in content_shape.text_frame.paragraphs:
            p.font.size = Pt(18)
            p.font.color.rgb = text_color
            p.alignment = alignment

    output = io.BytesIO()
    new_prs.save(output)
    output.seek(0)
    return output

# --- UI 介面 ---
uploaded_file = st.file_uploader("請上傳原始 PPT 檔案 (.pptx)", type=["pptx"])

if uploaded_file:
    # 1. 執行提取
    with st.spinner("正在解析原始投影片內容..."):
        extracted_data = extract_text_from_ppt(uploaded_file)
    
    st.success(f"成功讀取 {len(extracted_data)} 頁投影片！")

    # 2. 提供風格選項
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("風格 A：專業商務 (Blue)")
        st.caption("特點：左對齊、商務藍、高對比白背景")
        ppt_a = redesign_ppt(extracted_data, style="business")
        st.download_button("下載商務版型", ppt_a, "business_redesign.pptx")

    with col2:
        st.subheader("風格 B：未來科技 (Cyber)")
        st.caption("特點：置中對齊、螢光綠標題、深色背景")
        ppt_b = redesign_ppt(extracted_data, style="cyber")
        st.download_button("下載科技版型", ppt_b, "cyber_redesign.pptx")