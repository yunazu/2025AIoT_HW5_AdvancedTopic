import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
import io

st.set_page_config(page_title="PPT AI Style Transformer", page_icon="🪄")
st.title("🪄 PPT 視覺風格強行轉換器")

# --- 風格定義字典 ---
STYLES = {
    "科技深邃藍": {
        "bg_color": RGBColor(10, 20, 50),
        "title_color": RGBColor(0, 255, 255), # 螢光青
        "text_color": RGBColor(200, 230, 255),
        "font_name": "Arial"
    },
    "極簡商務白": {
        "bg_color": RGBColor(255, 255, 255),
        "title_color": RGBColor(0, 51, 102),  # 深藍
        "text_color": RGBColor(60, 60, 60),
        "font_name": "Microsoft JhengHei"
    },
    "時尚活力橘": {
        "bg_color": RGBColor(40, 40, 40),
        "title_color": RGBColor(255, 102, 0), # 亮橘
        "text_color": RGBColor(240, 240, 240),
        "font_name": "Verdana"
    }
}

def transform_ppt(uploaded_file, selected_style):
    prs = Presentation(uploaded_file)
    style_config = STYLES[selected_style]

    for slide in prs.slides:
        # 1. 強制設定背景顏色
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = style_config["bg_color"]
        
        # 2. 遍歷所有形狀 (包含圖片以外的所有物件)
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # 強制修改字體與顏色
                    run.font.color.rgb = style_config["title_color"] if shape == slide.shapes.title else style_config["text_color"]
                    run.font.name = style_config["font_name"]
                    run.font.bold = True if shape == slide.shapes.title else False

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- UI 介面 ---
src_file = st.file_uploader("1. 上傳原始 PPT", type=["pptx"])
style_choice = st.selectbox("2. 選擇 AI 重新設計的風格", list(STYLES.keys()))

if src_file:
    if st.button("立即套用 AI 風格並更換版型"):
        with st.spinner("正在重新計算版型配色..."):
            result_ppt = transform_ppt(src_file, style_choice)
            st.success(f"成功將簡報轉換為【{style_choice}】風格！")
            st.download_button(
                label="📥 下載新版簡報",
                data=result_ppt,
                file_name=f"redesigned_{style_choice}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )