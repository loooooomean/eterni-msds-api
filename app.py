from fastapi import FastAPI, Body
from fastapi.responses import JSONResponse, FileResponse
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tempfile
import os
import re
import json
from datetime import datetime

app = FastAPI()
TEMP_DIR = tempfile.gettempdir()

def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_three_line_borders(table):
    """应用科学三线表样式"""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name, sz in [('top', '12'), ('bottom', '12')]: # 1.5pt
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:color'), 'auto')
        tblBorders.append(b)
    for border_name in ['left', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'none')
        tblBorders.append(b)
    tblPr.append(tblBorders)

def set_header_border(cell):
    """设置表头下方的细线 (1.0pt)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:color'), 'auto')
    tcBorders.append(bottom)
    tcPr.append(tcBorders)

def add_separator_line(paragraph):
    """在段落下发添加一条横向分割线"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6') # 0.75pt
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)

def parse_content(paragraph, text, product_name=""):
    """解析内容，处理换行乱码，并自动加粗产品名"""
    content = str(text).replace('\\n', '\n').replace('<br>', '\n').replace('<br/>', '\n')
    if product_name and product_name in content:
        content = content.replace(product_name, f"**{product_name}**")
    
    chunks = content.split('**')
    for i, chunk in enumerate(chunks):
        lines = chunk.split('\n')
        for j, line in enumerate(lines):
            if line:
                run = paragraph.add_run(line)
                if i % 2 == 1: run.bold = True
            if j < len(lines) - 1:
                paragraph.add_run().add_break()

@app.post("/generate_document")
async def generate_document(
    content: str = Body(...),
    product_name: str = Body(""),
    product_model: str = Body("")
):
    # JSON 兼容清洗
    content = content.strip().replace('```json', '').replace('```', '')
    try:
        data = json.loads(content)
        content = data.get('docx') or data.get('output') or list(data.values())[0]
    except: pass
    content = str(content).replace('\\n', '\n')

    doc = Document()
    section = doc.sections[0]
    section.header_distance = Cm(0) # 页眉上边距 0cm
    
    # 全局样式：Times New Roman, 1.5倍行距，0段间距
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    fmt = style.paragraph_format
    fmt.line_spacing = 1.5
    fmt.space_before = fmt.space_after = Pt(0)

    # 1. 页眉排版 (Logo大小减半，靠下对齐)
    header = section.header
    htable = header.add_table(1, 2, Inches(6.5))
    htable.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
    htable.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
    
    if os.path.exists("logo.png"):
        htable.rows[0].cells[0].paragraphs[0].add_run().add_picture("logo.png", width=Inches(0.6))
    
    p_meta = htable.rows[0].cells[1].paragraphs[0]
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    meta_run = p_meta.add_run(f"{product_name}  {product_model}")
    meta_run.font.size = Pt(11)
    meta_run.font.color.rgb = RGBColor(120, 120, 120)
    add_separator_line(header.add_paragraph())

    # 2. 核心逻辑解析
    lines = content.split('\n')
    in_table = False
    table_data = []

    for line in lines:
        stripped = line.strip()
        if not stripped: continue

        # 表格处理 (三线表)
        if stripped.startswith('|'):
            in_table = True
            if '---' not in stripped:
                table_data.append([c.strip() for c in stripped.strip('|').split('|')])
            continue
        
        if in_table and not stripped.startswith('|'):
            if table_data:
                table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                set_three_line_borders(table)
                for r_idx, row in enumerate(table_data):
                    for c_idx, val in enumerate(row):
                        cell = table.rows[r_idx].cells[c_idx]
                        parse_content(cell.paragraphs[0], val, product_name)
                        if r_idx == 0: set_header_border(cell)
            in_table = False; table_data = []

        # 标题与正文样式
        if stripped.startswith('# '): # 一级标题 (小二，左对齐，组合信息)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            title = stripped[2:]
            today = datetime.now().strftime("%B %d, %Y")
            full_title = f"{title}\nETERNI {product_name}\nIssue Date: {today}"
            run = p.add_run(full_title)
            run.font.size = Pt(18)
            run.bold = True
            add_separator_line(doc.add_paragraph()) # 标题下方分割线

        elif stripped.startswith('## '): # 二级标题 (蓝底白字，前有空行和分割线)
            doc.add_paragraph() # 前空行
            add_separator_line(doc.add_paragraph()) # 前分割线
            t_h2 = doc.add_table(1, 1, Inches(6.5))
            cell = t_h2.rows[0].cells[0]
            set_cell_background(cell, '0033CC')
            run = cell.paragraphs[0].add_run(stripped[3:])
            run.font.size = Pt(14)
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.bold = True
            doc.add_paragraph() # 后空行

        elif stripped.startswith('### '): # 三级标题 (小四加粗)
            run = doc.add_paragraph().add_run(stripped[4:])
            run.font.size = Pt(12); run.bold = True

        else:
            p = doc.add_paragraph()
            if stripped.startswith(('- ', '* ')):
                p.style = 'List Bullet'
                parse_content(p, stripped[2:], product_name)
            else:
                parse_content(p, stripped, product_name)

    # 3. 保存与下载
    file_name = f"ETERNI_{product_model or 'Doc'}.docx"
    path = os.path.join(TEMP_DIR, file_name)
    doc.save(path)
    return {"url": f"https://eterni-msds-api-1.onrender.com/download/{file_name}"}

@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(TEMP_DIR, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename=filename, headers={"Content-Disposition": f'attachment; filename="{filename}"'})
    return JSONResponse(status_code=404, content={"error": "File not found"})
