import streamlit as st
from pptx import Presentation
import io

st.set_page_config(page_title="PPT Style Transfer", page_icon="🎨")
st.title("🪄 PPT 模板風格轉換器")
st.write("上傳原始簡報與目標模板，AI 將自動完成內容移植。")

def transfer_style(source_ppt, template_ppt):
    source = Presentation(source_ppt)
    template = Presentation(template_ppt)
    
    # 建立一個新的簡報，起始於模板的母片架構
    # 這裡我們直接在 template 後面新增投影片，避免遺失模板的背景
    
    for slide in source.slides:
        # 從模板中選擇一個版型 (通常索引 1 是「標題+內容」)
        try:
            layout = template.slide_layouts[1] 
        except:
            layout = template.slide_layouts[0]
            
        new_slide = template.slides.add_slide(layout)
        
        # 1. 移植標題
        if slide.shapes.title and new_slide.shapes.title:
            new_slide.shapes.title.text = slide.shapes.title.text
            
        # 2. 移植主要內容文字
        source_placeholders = [sp for sp in slide.placeholders if sp != slide.shapes.title]
        target_placeholders = [tp for tp in new_slide.placeholders if tp != new_slide.shapes.title]
        
        if source_placeholders and target_placeholders:
            # 簡單的一對一移植
            target_placeholders[0].text = source_placeholders[0].text

    output = io.BytesIO()
    template.save(output)
    output.seek(0)
    return output

# --- UI 介面 ---
col1, col2 = st.columns(2)

with col1:
    src_file = st.file_uploader("1. 上傳【原始檔案】(內容來源)", type=["pptx"])
with col2:
    tpl_file = st.file_uploader("2. 上傳【空的模板】(風格來源)", type=["pptx"])

if src_file and tpl_file:
    if st.button("開始轉換風格"):
        with st.spinner("正在將內容移植至新模板..."):
            result_ppt = transfer_style(src_file, tpl_file)
            
            st.success("轉換完成！")
            st.download_button(
                label="📥 下載轉換後的簡報",
                data=result_ppt,
                file_name="styled_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

st.divider()
st.info("💡 提示：模板檔案建議包含您想要的背景、Logo 與字體設定。本工具會將原始文字填入模板的『標題與內容』框中。")