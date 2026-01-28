import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- 保留你原本的核心邏輯 ---
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
        # 第一列：音標
        c1 = table.cell(0, idx)
        set_cell_margins_zero(c1)
        c1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        run1 = p1.add_run(pinyin)
        run1.font.size = Pt(11)
        run1.font.name = 'Times New Roman'
        run1.font.color.rgb = get_tone_color(pinyin)
        run1.bold = True

        # 第二列：漢字
        c2 = table.cell(1, idx)
        set_cell_margins_zero(c2)
        c2.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        run2 = p2.add_run(hanzi)
        run2.font.size = Pt(20)
        run2.font.name = '標楷體'
        run2._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

    spacer = doc.add_paragraph()
    spacer.paragraph_format.line_spacing = Pt(12)

# --- Streamlit 網頁介面 ---
st.set_page_config(page_title="漢字音標轉換器", page_icon="📝")

st.title("📝 漢字音標轉 Word 工具")
st.markdown("""
將格式為 `漢字(音標)` 的文字檔轉換為漂亮的 Word 表格。
1. 上傳你的 **.txt** 檔案。
2. 系統會自動處理轉換。
3. 點擊按鈕下載產出的 **.docx** 檔。
""")

uploaded_file = st.file_uploader("選擇 TXT 檔案", type="txt")

if uploaded_file is not None:
    # 讀取檔案內容
    stringio = uploaded_file.getvalue().decode("utf-8")
    lines = stringio.splitlines()

    # 建立 Word 文件
    doc = Document()
    doc.styles['Normal'].font.name = '標楷體'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

    progress_bar = st.progress(0)
    
    for i, line in enumerate(lines):
        matches = re.findall(r'([\u4e00-\u9fff])\(([^)]+)\)', line)
        if matches:
            create_row_table(doc, matches)
        elif line.strip():
            p = doc.add_paragraph(line)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            doc.add_paragraph()
        progress_bar.progress((i + 1) / len(lines))

    # 將檔案儲存在記憶體中供下載
    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    st.success("✅ 轉換完成！")
    st.download_button(
        label="📥 下載 Word 檔案",
        data=file_stream,
        file_name=f"轉換結果_{uploaded_file.name.replace('.txt', '')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )