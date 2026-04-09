from fastapi import FastAPI, Body
from fastapi.responses import JSONResponse, FileResponse
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tempfile
import os
import re

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
    """设置标准的学术三线表边框（上下粗线）"""
    tblBorders = OxmlElement('w:tblBorders')
    
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '12')  # 1.5 pt
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), 'auto')
    
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12') # 1.5 pt
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), 'auto')

    insideH = OxmlElement('w:insideH')
    insideH.set(qn('w:val'), 'none')
    insideV = OxmlElement('w:insideV')
    insideV.set(qn('w:val'), 'none')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'none')
    right = OxmlElement('w:right')
    right.set(qn('w:val'), 'none')
    
    tblBorders.append(top)
    tblBorders.append(left)
    tblBorders.append(bottom)
    tblBorders.append(right)
    tblBorders.append(insideH)
    tblBorders.append(insideV)
    
    table._tbl.tblPr.append(tblBorders)

def set_header_bottom_border(cell):
    """设置三线表表头的下划线（细线）"""
    tcBorders = OxmlElement('w:tcBorders')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8') # 1.0 pt
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), 'auto')
    tcBorders.append(bottom)
    cell._tc.get_or_add_tcPr().append(tcBorders)

def parse_inline_bold(paragraph, text, product_name=""):
    """解析文本，自动加粗产品名，并处理 <br> 乱码换行"""
    content = str(text) if text else ""
    
    # 处理乱码：将 <br> 或 <br/> 替换为真实的换行符
    content = re.sub(r'<br\s*/?>', '\n', content)
    
    # 自动对产品名称进行加粗匹配
    if product_name and product_name in content:
        if f"**{product_name}**" not in content:
            content = content.replace(product_name, f"**{product_name}**")
            
    chunks = content.split('**')
    for i, chunk in enumerate(chunks):
        sub_chunks = chunk.split('\n')
        for j, sub in enumerate(sub_chunks):
            if sub:
                run = paragraph.add_run(sub)
                if i % 2 == 1: 
                    run.bold = True
            # Word 里的换行必须用 add_break() 才能生效
            if j < len(sub_chunks) - 1:
                paragraph.add_run().add_break()

@app.post("/generate_document")
async def generate_document(
    content: str = Body(..., description="Markdown文本"),
    product_name: str = Body("Product Name"),
    product_model: str = Body("Model")
):
    doc = Document()
    
    # --- 页面与全局样式设置 ---
    section = doc.sections[0]
    section.header_distance = Cm(0)  # 页眉上边距设为 0 cm
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    
    # 行间距 1.5 倍，段前段后 0 距离
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    p_name = product_name if product_name else "ETERNI Product"
    p_model = product_model if product_model else ""

    # --- 1. 页眉排版 ---
    header = section.header
    htable = header.add_table(1, 2, Inches(6.5))
    cell_left = htable.rows[0].cells[0]
    
    # Logo 变为一半大小：宽 0.6 英寸
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        try:
            cell_left.paragraphs[0].add_run().add_picture(logo_path, width=Inches(0.6))
        except:
            cell_left.paragraphs[0].add_run("ETERNI").bold = True 
    else:
        cell_left.paragraphs[0].add_run("ETERNI").bold = True 

    cell_right = htable.rows[0].cells[1]
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_text = p_right.add_run(f"{p_name}   {p_model}")
    run_text.font.size = Pt(11)
    run_text.font.color.rgb = RGBColor(120, 120, 120) 
    
    header_para = header.add_paragraph()
    p_pr = header_para._p.get_or_add_pPr()
    p_pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    p_pbdr.append(bottom)
    p_pr.append(p_pbdr)

    # --- 2. 动态内容解析 ---
    lines = content.split('\n')
    in_table = False
    table_data = []

    for line in lines:
        stripped = line.strip()
        if not stripped: continue
        
        # 表格数据收集
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            if '---' in stripped: continue
            row = [c.strip() for c in stripped.strip('|').split('|')]
            table_data.append(row)
            continue
        
        # 绘制三线表
        if in_table and not stripped.startswith('|'):
            if table_data:
                cols = len(max(table_data, key=len))
                table = doc.add_table(rows=1, cols=cols)
                set_three_line_borders(table) # 应用三线表样式
                
                for i, heading in enumerate(table_data[0]):
                    if i < len(table.rows[0].cells):
                        cell = table.rows[0].cells[i]
                        parse_inline_bold(cell.paragraphs[0], heading, p_name)
                        set_header_bottom_border(cell) # 表头底部加线
                        
                for row_data in table_data[1:]:
                    row_cells = table.add_row().cells
                    for i, item in enumerate(row_data):
                        if i < len(row_cells): 
                            parse_inline_bold(row_cells[i].paragraphs[0], item, p_name)
            in_table = False
            table_data = []

        # 解析 Markdown 语法
        if stripped.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT # 靠左对齐
            
            # 自动补全品牌和产品名 (如果 LLM 没有写全)
            title_text = stripped[2:]
            if p_name.lower() not in title_text.lower():
                title_text = f"{title_text}\nETERNI {p_name} {p_model}".strip()
                
            run = p.add_run(title_text)
            run.font.size = Pt(18) # 小二字号
            run.bold = True
            
        elif stripped.startswith('## '):
            doc.add_paragraph() # 二级标题前空行
            table_sect = doc.add_table(rows=1, cols=1)
            cell = table_sect.rows[0].cells[0]
            set_cell_background(cell, '0033CC') # 蓝底
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(stripped[3:])
            run.font.size = Pt(14) # 四号字号
            run.font.color.rgb = RGBColor(255, 255, 255) # 白字
            run.bold = True
            
        elif stripped.startswith('### '):
            p = doc.add_paragraph()
            run = p.add_run(stripped[4:])
            run.font.size = Pt(12) # 小四字号
            run.bold = True
            
        elif stripped.startswith('- ') or stripped.startswith('* '):
            # 圆点分点列表
            p = doc.add_paragraph(style='List Bullet')
            parse_inline_bold(p, stripped[2:], p_name)
            
        else:
            p = doc.add_paragraph()
            parse_inline_bold(p, stripped, p_name)

    # --- 3. 生成与下载 ---
    clean_model = re.sub(r'[^a-zA-Z0-9]', '_', str(p_model))
    if not clean_model:
        clean_model = "Document"
    file_name = f"ETERNI_{clean_model}.docx"
    
    temp_path = os.path.join(TEMP_DIR, file_name)
    doc.save(temp_path)
    
    download_url = f"https://eterni-msds-api-1.onrender.com/download/{file_name}"
    return {"url": download_url}

@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(TEMP_DIR, filename)
    if os.path.exists(file_path):
        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"'
        }
        return FileResponse(
            path=file_path, 
            filename=filename, 
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers=headers
        )
    return JSONResponse(status_code=404, content={"error": f"File {filename} not found on server."})
