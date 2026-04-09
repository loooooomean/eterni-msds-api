from fastapi import FastAPI, Body
from fastapi.responses import JSONResponse
import requests
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tempfile
import os

app = FastAPI()

def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def parse_inline_bold(paragraph, text):
    chunks = str(text).split('**')
    for i, chunk in enumerate(chunks):
        run = paragraph.add_run(chunk)
        if i % 2 == 1: 
            run.bold = True

@app.post("/generate_document")
async def generate_document(
    content: str = Body(..., description="Markdown文本"),
    product_name: str = Body("Product Name"),
    product_model: str = Body("Model")
):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    # --- 页眉 Logo 与信息 ---
    header = doc.sections[0].header
    htable = header.add_table(1, 2, Inches(6.5))
    cell_left = htable.rows[0].cells[0]
    
    # 尝试加载仓库中的 logo.png
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        cell_left.paragraphs[0].add_run().add_picture(logo_path, width=Inches(1.2))
    else:
        cell_left.paragraphs[0].add_run("ETERNI").bold = True 

    cell_right = htable.rows[0].cells[1]
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_text = p_right.add_run(f"{product_name}   {product_model}")
    run_text.font.size = Pt(11)
    run_text.font.color.rgb = RGBColor(120, 120, 120) 
    
    # 页眉横线
    header_para = header.add_paragraph()
    p_pr = header_para._p.get_or_add_pPr()
    p_pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    p_pbdr.append(bottom)
    p_pr.append(p_pbdr)

    # --- 内容解析 ---
    lines = content.split('\n')
    in_table = False
    table_data = []

    for line in lines:
        stripped = line.strip()
        if not stripped: continue
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            if '---' in stripped: continue
            row = [c.strip() for c in stripped.strip('|').split('|')]
            table_data.append(row)
            continue
        if in_table and not stripped.startswith('|'):
            if table_data:
                cols = len(max(table_data, key=len))
                table = doc.add_table(rows=1, cols=cols)
                table.style = 'Table Grid'
                for i, heading in enumerate(table_data[0]):
                    if i < len(table.rows[0].cells):
                        cell = table.rows[0].cells[i]
                        parse_inline_bold(cell.paragraphs[0], heading)
                        set_cell_background(cell, 'E0E0E0') 
                for row_data in table_data[1:]:
                    row_cells = table.add_row().cells
                    for i, item in enumerate(row_data):
                        if i < len(row_cells): parse_inline_bold(row_cells[i].paragraphs[0], item)
            in_table = False
            table_data = []

        if stripped.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(stripped[2:])
            run.font.size = Pt(18); run.bold = True
        elif stripped.startswith('## '):
            doc.add_paragraph()
            table_sect = doc.add_table(rows=1, cols=1)
            cell = table_sect.rows[0].cells[0]
            set_cell_background(cell, '0033CC') 
            run = cell.paragraphs[0].add_run(stripped[3:])
            run.font.color.rgb = RGBColor(255, 255, 255); run.bold = True
        else:
            p = doc.add_paragraph()
            parse_inline_bold(p, stripped)

    # --- 保存并上传获取 URL ---
    temp_path = os.path.join(tempfile.gettempdir(), f"ETERNI_{product_model}.docx")
    doc.save(temp_path)
    
    # 将文件上传到临时托管服务器 (file.io 是免费的)
    download_url = ""
    try:
        with open(temp_path, 'rb') as f:
            # 上传文件，设置为 14 天有效
            r = requests.post('https://file.io/?expires=14d', files={'file': f})
            download_url = r.json().get('link', '')
    except:
        download_url = "Upload Failed"

    # 返回 Coze 能读懂的 JSON 格式
    return {"url": download_url}
