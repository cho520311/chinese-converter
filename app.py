import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- 設定網頁標題與風格 ---
st.set_page_config(page_title="雅致漢字轉換器", page_icon="📜", layout="wide")

# 加入 CSS 讓介面更雅致
st.markdown("""
    <style>
    .main {
        background-color: #fdfaf5; /* 輕微的米白色背景 */
    }
    h1 {
        color: #4a4a4a;
        font-family: "Microsoft JhengHei", sans-serif;
        font-weight: 300;
        text-align: center;
    }
    .stMarkdown {
        font-size: 1.2rem !important;
        color: #555;
    }
    /* 放大上傳框文字 */
    div[data-testid="stFileUploader"] section {
        padding: 2rem;
        border: 1px dashed #d3c4a8;
        background-color: #fffcf9;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 保留原本的核心邏輯 (get_tone_color, create_row_table 等) ---
def get_tone_color(py_text):
    py = py_text.lower().strip()
    if re.search(r'5$', py) or any(c in py for c in ['â', 'ê', 'î', 'ô', 'û', '̂', 'ˆ', '^']):
        return RGBColor(0, 0, 255)
    if py.endswith(('p', 't', 'k')):
        return RGBColor(255, 0, 0)
    marks = ['á', 'à', 'ā', 'ǎ', 'í', 'ì', 'ī', 'ǐ', 'ú', 'ù', 'ū', 'ǔ', 'é', 'è', 'ē', 'ě', 'ó', 'ò', 'ō', 'ǒ', '̍', '́', '̀', '̌', '̄']
    if any(c in py for c in marks) or re.search(r'[234678]$', py):
        return RGBColor(255, 0, 0)
    return RGBColor(0, 0, 255)

def set_cell_margins_zero(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for m in ['top', 'left', 'bottom', 'right']:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), '0')
        node.set(qn('w:type'), 'dxa')
        mar.append(node)
    tcPr.append(mar)

def create_row_table(doc, row_data):
    if not row_data: return
    table = doc.add_table(rows=2, cols=len(row_data))
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in table.rows:
        row.allow_break_across_pages = False
    for idx, (hanzi, pinyin) in enumerate(row_data):
        c1 = table.cell(0, idx)
        set_cell_margins_zero(c1)
        c1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p1.add_run(pinyin)
        run1.font.size = Pt(11)
        run1.font.name = 'Times New Roman'
        run1.font.color.rgb = get_tone_color(pinyin)
        run1.bold = True
        c2 = table.cell(1, idx)
        set_cell_margins_zero(c2)
        c2.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = p2.add_run(hanzi)
        run2.font.size = Pt(20)
        run2.font.name = '標楷體'
        run2._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    spacer = doc.add_paragraph()
    spacer.paragraph_format.line_spacing = Pt(12)

# --- 介面排版 ---
st.title("📜 漢字音標雅致轉換工具")
st.write("---")

# 範例預覽區
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 💡 格式範例")
    st.info("請確保您的 TXT 檔案內容格式如下：")
    st.code("學(xué)而(ér)時(shí)習(xí)之(zhī)\n不(bù)亦(yì)說(yuè)乎(hū)", language="text")

with col2:
    st.markdown("### 📝 溫馨提示")
    st.write("1. 系統會自動根據聲調標示顏色。")
    st.write("2. 轉換完成後請下載 Word 檔。")
    st.write("3. 下載後建議使用標楷體查看。")

st.write("---")

# 上傳區
uploaded_file = st.file_uploader("📂 選擇檔案 (請上傳您的 .txt 檔)", type="txt")

if uploaded_file:
    # 讀取檔案
    content = uploaded_file.read().decode("utf-8")
    lines = content.splitlines()

    # 建立 Word
    doc = Document()
    doc.styles['Normal'].font.name = '標楷體'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

    for line in lines:
        matches = re.findall(r'([\u4e00-\u9fff])\(([^)]+)\)', line)
        if matches:
            create_row_table(doc, matches)
        elif line.strip():
            p = doc.add_paragraph(line)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            doc.add_paragraph()

    # 下載
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.balloons() # 撒花特效
    st.success("✨ 轉換成功！請點擊下方按鈕，檔案將儲存至您的下載資料夾。")
    st.download_button(
        label="📥 下載轉換後的 Word 檔案",
        data=buffer,
        file_name=f"轉換結果_{uploaded_file.name.replace('.txt', '')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True # 讓按鈕變大
    )
