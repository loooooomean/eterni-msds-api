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
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'bottom']:
        elem = OxmlElement(f'w:{side}')
        elem.set(qn('w:val'), 'single')
        elem.set(qn('w:sz'), '12') # 1.5 pt
        elem.set(qn('w:color'), 'auto')
        tblBorders.append(elem)
    for side in ['left', 'right', 'insideH', 'insideV']:
        elem = OxmlElement(f'w:{side}')
        elem.set(qn('w:val'), 'none')
        tblBorders.append(elem)
    table._tbl.tblPr.append(tblBorders)

def parse_inline_bold(paragraph, text, product_name=""):
    content = str(text) if text else ""
    # 核心修复：处理各种乱码换行符
    content = content.replace('\\n', '\n').replace('<br>', '\n').replace('<br/>', '\n')
    
    if product_name and product_name in content:
        if f"**{product_name}**" not in content:
            content = content.replace(product_name, f"**{product_name}**")
            
    chunks = content.split('**')
    for i, chunk in enumerate(chunks):
        sub_chunks = chunk.split('\n')
        for j, sub in enumerate(sub_chunks):
            if sub:
                run = paragraph.add_run(sub)
                if i % 2 == 1: run.bold = True
            if j < len(sub_chunks) - 1:
                paragraph.add_run().add_break()

@app.post("/generate_document")
async def generate_document(
    content: str = Body(..., description="文本内容"),
    product_name: str = Body(""),
    product_model: str = Body("")
):
    # 自动兼容逻辑：如果收到的还是 JSON 字符串，尝试自动解包
    try:
        data = json.loads(content)
        if isinstance(data, dict):
            content = data.get('docx') or data.get('output') or list(data.values())[0]
    except: pass

    doc = Document()
    section = doc.sections[0]
    section.header_distance = Cm(0) # 页眉顶满
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_before = style.paragraph_format.space_after = Pt(0)

    # 1. 页眉 - 垂直靠下对齐
    header = section.header
    htable = header.add_table(1, 2, Inches(6.5))
    for cell in htable.rows[0].cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM # 靠下显示
    
    # Logo 宽度减半 (0.6 inch)
    if os.path.exists("logo.png"):
        htable.rows[0].cells[0].paragraphs[0].add_run().add_picture("logo.png", width=Inches(0.6))
    
    p_right = htable.rows[0].cells[1].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_h = p_right.add_run(f"{product_name}  {product_model}")
    run_h.font.size = Pt(11); run_h.font.color.rgb = RGBColor(120, 120, 120)

    # 2. 正文解析
    lines = content.split('\n')
    in_table = False
    table_data = []

    for line in lines:
        stripped = line.strip()
        if not stripped: continue
        
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
                        parse_inline_bold(cell.paragraphs[0], val, product_name)
                        if r_idx == 0: # 表头
                            set_cell_background(cell, 'E0E0E0')
            in_table = False; table_data = []

        if stripped.startswith('# '): # 一级标题：靠左，小二
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(stripped[2:])
            run.font.size = Pt(18); run.bold = True
        elif stripped.startswith('## '): # 二级标题：蓝底白字，前后空行
            doc.add_paragraph() # 前空行
            t = doc.add_table(1, 1); cell = t.rows[0].cells[0]
            set_cell_background(cell, '0033CC')
            run = cell.paragraphs[0].add_run(stripped[3:])
            run.font.size = Pt(14); run.font.color.rgb = RGBColor(255, 255, 255); run.bold = True
            doc.add_paragraph() # 后空行
        elif stripped.startswith('### '): # 三级标题：小四加粗
            run = doc.add_paragraph().add_run(stripped[4:])
            run.font.size = Pt(12); run.bold = True
        else:
            p = doc.add_paragraph()
            if stripped.startswith('- ') or stripped.startswith('* '):
                p.style = 'List Bullet'
                parse_inline_bold(p, stripped[2:], product_name)
            else:
                parse_inline_bold(p, stripped, product_name)

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
