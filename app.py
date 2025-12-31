import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io

# --- 頁面設定 ---
st.set_page_config(page_title="AI PPT Stylist", page_icon="📊")
st.title("🎨 AI 簡報風格設計器")
st.write("輸入主題，一鍵生成兩種不同風格的 PPT 範本！")

# --- 核心功能：生成 PPT ---
def create_ppt(theme_name, style="business"):
    prs = Presentation()
    
    # 定義風格參數
    if style == "business":
        bg_color = RGBColor(255, 255, 255) # 白色背景
        title_color = RGBColor(0, 51, 102) # 深藍色標題
        align = PP_ALIGN.LEFT
        font_name = "Arial"
    else:
        bg_color = RGBColor(43, 43, 43)    # 深灰色背景
        title_color = RGBColor(255, 102, 0) # 亮橘色標題
        align = PP_ALIGN.CENTER
        font_name = "Verdana"

    # 建立三頁投影片
    slides_content = [
        ["標題頁", f"關於 {theme_name} 的分析報告", "報告人：AI 助手"],
        ["重點摘要", "核心技術探討", "1. 自動化流程\n2. AI 視覺設計\n3. 使用者體驗優化"],
        ["結論", "未來展望", "持續進化，創造更多 AI 應用的可能性。"]
    ]

    for slide_data in slides_content:
        slide_layout = prs.slide_layouts[1] # 使用標題+內容版面
        slide = prs.slides.add_slide(slide_layout)
        
        # 設定背景顏色 (僅示範，進階可加圖案)
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bg_color

        # 設定標題風格
        title = slide.shapes.title
        title.text = slide_data[1]
        title_text_frame = title.text_frame.paragraphs[0]
        title_text_frame.font.bold = True
        title_text_frame.font.size = Pt(36)
        title_text_frame.font.color.rgb = title_color
        title_text_frame.alignment = align
        
        # 設定內容風格
        content = slide.placeholders[1]
        content.text = slide_data[2]
        for para in content.text_frame.paragraphs:
            para.font.size = Pt(18)
            if style == "modern":
                para.font.color.rgb = RGBColor(200, 200, 200) # 淺灰文字

    # 將 PPT 存入記憶體體中回傳
    binary_output = io.BytesIO()
    prs.save(binary_output)
    binary_output.seek(0)
    return binary_output

# --- UI 介面 ---
topic = st.text_input("請輸入簡報主題：", placeholder="例如：2025 AI 發展趨勢")

if topic:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("風格 A：專業商務藍")
        st.info("特點：白色背景、深藍標題、靠左對齊。適合正式會議。")
        ppt_a = create_ppt(topic, style="business")
        st.download_button(
            label="下載商務風格 PPT",
            data=ppt_a,
            file_name=f"{topic}_business.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    with col2:
        st.subheader("風格 B：極簡現代黑")
        st.warning("特點：深色背景、亮橘標題、置中對齊。適合技術分享。")
        ppt_b = create_ppt(topic, style="modern")
        st.download_button(
            label="下載現代風格 PPT",
            data=ppt_b,
            file_name=f"{topic}_modern.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )