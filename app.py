import io
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(page_title="木谷さんW8LY専用エキスパート整理ツール", page_icon="📋", layout="wide")

# 指定カラーテーマ（背景: #355E3B, ボックス: #F5EFD6, 文字: #895129）とスタイル設定
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Zen+Kaku+Gothic+New:wght@700;900&display=swap');

    .stApp {
        background-color: #355E3B !important;
        background-image: none !important;
    }
    
    .main-card {
        background-color: #F5EFD6 !important;
        border: 2px solid #895129 !important;
        border-radius: 24px;
        padding: 40px 35px;
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.25);
        color: #895129 !important;
        text-align: center;
        max-width: 750px;
        margin: 10px auto 25px auto;
    }

    .badge-theme {
        background-color: #895129 !important;
        color: #F5EFD6 !important;
        padding: 6px 20px;
        border-radius: 30px;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 2.5px;
        text-transform: uppercase;
        display: inline-block;
        margin-bottom: 14px;
    }

    h1.main-title {
        font-family: 'Zen Kaku Gothic New', 'Hiragino Sans', sans-serif;
        color: #895129 !important;
        font-size: 28px;
        font-weight: 900;
        margin: 0 0 12px 0;
        letter-spacing: -0.5px;
        -webkit-text-fill-color: #895129 !important;
    }

    p.sub-desc {
        color: #895129 !important;
        font-size: 13px;
        font-weight: 700;
        line-height: 1.6;
        margin-bottom: 0;
    }

    .stButton > button {
        background-color: #895129 !important;
        color: #F5EFD6 !important;
        border: none !important;
        padding: 14px 32px !important;
        border-radius: 50px !important;
        font-family: 'Zen Kaku Gothic New', sans-serif !important;
        font-size: 16px !important;
        font-weight: 900 !important;
        letter-spacing: 1.5px !important;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.2) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
    }

    .stButton > button:hover {
        transform: translateY(-2px) !important;
        background-color: #704020 !important;
    }

    div[data-testid="stTextArea"] textarea {
        background-color: #F5EFD6 !important;
        color: #222222 !important;
        border: 2px solid #895129 !important;
        border-radius: 16px !important;
        font-size: 13.5px !important;
    }
    
    div[data-testid="stTextArea"] label {
        color: #F5EFD6 !important;
        font-weight: 800 !important;
    }

    /* メールコピー表示エリアのフォント＆スタイル */
    .email-preview-box {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 2px solid #895129 !important;
        border-radius: 16px;
        padding: 25px;
        font-family: Arial, Helvetica, 'Segoe UI', sans-serif !important;
        font-size: 13.5px !important;
        line-height: 1.6 !important;
        user-select: text !important;
        -webkit-user-select: text !important;
    }

    mark.yellow-hl {
        background-color: #FFF2CC !important;
        color: #111111 !important;
        padding: 0 2px;
        border-radius: 2px;
    }

    /* アングル名タイトル（青文字＋下線） */
    .angle-category-title {
        color: #2B78A0 !important;
        font-weight: bold !important;
        font-size: 14.5px !important;
        border-bottom: 1.5px solid #2B78A0 !important;
        padding-bottom: 3px !important;
        margin: 20px 0 12px 0 !important;
        display: block !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-card">
    <span class="badge-theme">W8LY SPECIAL TOOL</span>
    <h1 class="main-title">木谷さんW8LY専用エキスパート整理ツール</h1>
    <p class="sub-desc">メール本文をコピペするだけで、アングルタイトル・名前・企業名・Q&A回答の自動装飾メール文面とエクセルを同時に生成します。</p>
</div>
""", unsafe_allow_html=True)

input_text = st.text_area("▼ 送信されてきたメール文面をここに貼り付けてください", height=260, placeholder="Hi Yui, ... から始まるメールテキストをそのままペースト")

def parse_single_expert(chunk_str, current_scope=""):
    if not chunk_str or not re.search(r'#\d+(?:\.\d+)?', chunk_str):
        return None
        
    header_match = re.search(r'#(\d+(?:\.\d+)?)\s*-\s*([^-]+?)\s*-\s*(.+?)(?=\n|$)', chunk_str)
    if not header_match:
        return None
        
    number = header_match.group(1).strip()
    name = header_match.group(2).strip()
    title = header_match.group(3).strip()
    
    lines = chunk_str.split('\n')
    summary_lines = []
    screened_lines = []
    avail_lines = []
    
    in_screened = False
    in_emp = False
    in_avail = False
    
    for line in lines[1:]:
        l = line.strip()
        if not l:
            continue
            
        if re.match(r'\[\s*(?:Screened|Re-screened).*?\]', l, re.IGNORECASE) or re.match(r'Screened\s+\d', l, re.IGNORECASE):
            in_screened = True
            in_emp = False
            in_avail = False
            screened_lines.append(l)
            continue
        elif "Employment History:" in l:
            in_screened = False
            in_emp = True
            in_avail = False
            continue
        elif "Availability:" in l:
            in_screened = False
            in_emp = False
            in_avail = True
            continue
        elif any(k in l for k in ["This specialist has not yet provided any availability", "Request Availability", "Book Now", "This specialist is based in", "Hourly Fee:"]):
            continue
            
        if in_screened:
            screened_lines.append(l)
        elif in_avail:
            if "Time Zone:" in l:
                continue
            avail_lines.append(l)
        elif not in_emp and not in_screened and not in_avail:
            summary_lines.append(l)
            
    summary_text = "\n".join(summary_lines)
    qa_text = "\n".join(screened_lines)
    
    if avail_lines:
        avail_text = "\n".join(avail_lines)
    else:
        avail_text = "回収中"
        
    return {
        "scope": current_scope,
        "number": number,
        "name": name,
        "title": title,
        "summary": summary_text,
        "qa": qa_text,
        "availability": avail_text
    }

def parse_full_email(raw_text):
    clean_raw = raw_text
    footer_keywords = ["Natsuko Shiga", "Research - Japan", "Disclaimer:", "Important:", "Compliance Reminder", "Third Bridge (Hong Kong)"]
    for kw in footer_keywords:
        if kw in clean_raw:
            clean_raw = clean_raw.split(kw)[0]
            
    lines = clean_raw.split('\n')
    
    structured_items = [] # list of {"type": "category", "text": ...} or {"type": "expert", "data": ...}
    parsed_experts = []
    
    current_category = ""
    current_expert_chunk = []
    
    for line in lines:
        l_str = line.strip()
        
        # アングル名タイトル判定 (例: [0916] Sanofi: や [0916] RBQM/Real time monitoring 海外:)
        if re.match(r'\[\d{4}\].+', l_str):
            if current_expert_chunk:
                exp_data = parse_single_expert("\n".join(current_expert_chunk), current_category)
                if exp_data:
                    structured_items.append({"type": "expert", "data": exp_data})
                    parsed_experts.append(exp_data)
                current_expert_chunk = []
                
            current_category = l_str
            structured_items.append({"type": "category", "text": l_str})
        elif re.match(r'#\d+(?:\.\d+)?\s*-', l_str):
            if current_expert_chunk:
                exp_data = parse_single_expert("\n".join(current_expert_chunk), current_category)
                if exp_data:
                    structured_items.append({"type": "expert", "data": exp_data})
                    parsed_experts.append(exp_data)
                current_expert_chunk = []
            current_expert_chunk.append(line)
        else:
            if current_expert_chunk:
                current_expert_chunk.append(line)
                
    if current_expert_chunk:
        exp_data = parse_single_expert("\n".join(current_expert_chunk), current_category)
        if exp_data:
            structured_items.append({"type": "expert", "data": exp_data})
            parsed_experts.append(exp_data)
            
    return structured_items, parsed_experts

def highlight_title_company(title):
    match = re.search(r'(\bat\s+)(.+?)(\s*\(\d{2}/\d{4}|\s*\[|\s*$)', title, re.IGNORECASE)
    if match:
        before_at = title[:match.start(2)]
        company = match.group(2)
        after_company = title[match.end(2):]
        return f'{before_at}<mark class="yellow-hl">{company}</mark>{after_company}'
    return title

def generate_formatted_email_html(structured_items):
    html_parts = []
    for item in structured_items:
        if item["type"] == "category":
            html_parts.append(f'<div class="angle-category-title">{item["text"]}</div>')
        elif item["type"] == "expert":
            exp = item["data"]
            highlighted_name = f'<mark class="yellow-hl">{exp["name"]}</mark>'
            highlighted_title = highlight_title_company(exp["title"])
            
            header_line = f'<span style="color: #2B78A0; font-weight: bold;">#{exp["number"]}</span> - <strong>{highlighted_name}</strong> - <strong>{highlighted_title}</strong>'
            
            block = f'<div style="margin-bottom: 15px;">{header_line}</div>'
            
            if exp['summary']:
                summary_formatted = exp['summary'].replace('\n', '<br>')
                block += f'<div style="margin-top: 8px;">{summary_formatted}</div>'
                
            if exp['qa']:
                qa_formatted_lines = []
                for line in exp['qa'].split('\n'):
                    line_str = line.strip()
                    if line_str.startswith('A:') or line_str.startswith('A：'):
                        qa_formatted_lines.append(f'<mark class="yellow-hl">{line_str}</mark>')
                    else:
                        qa_formatted_lines.append(line_str)
                qa_html = "<br>".join(qa_formatted_lines)
                block += f'<div style="margin-top: 8px;">{qa_html}</div>'
                
            avail_formatted = exp['availability'].replace('\n', '<br>')
            block += f'<div style="margin-top: 10px; margin-bottom: 20px;"><strong>Availability:</strong><br>{avail_formatted}</div>'
            
            html_parts.append(block)
            
    return "".join(html_parts)

def generate_excel_bytes(experts):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "全員一覧"
    
    headers = ["Scope", "Number", "Name", "Relevant Titles", "Relevant experience", "Availability"]
    
    font_bold = Font(name="Meiryo UI", size=9, bold=True)
    font_regular = Font(name="Meiryo UI", size=9)
    fill_header = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )
    
    ws["A1"] = "ThirdBridge"
    ws["A1"].font = font_bold
    
    for c_idx, h_text in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c_idx, value=h_text)
        cell.font = font_bold
        cell.fill = fill_header
        cell.alignment = Alignment(vertical="top")
        
    for r_idx, exp in enumerate(experts, 4):
        exp_full_text = exp['summary']
        if exp['qa']:
            exp_full_text += f"\n\n{exp['qa']}"
            
        row_vals = [
            exp['scope'],
            exp['number'],
            exp['name'],
            exp['title'],
            exp_full_text,
            exp['availability']
        ]
        
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            
            if c_idx == 2:
                try:
                    cell.value = float(val) if "." in str(val) else int(val)
                    cell.number_format = '0.0'
                except ValueError:
                    pass
                cell.alignment = Alignment(horizontal="center", vertical="top")
                
    col_widths = [18, 10, 18, 38, 60, 30]
    for c_i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(c_i)].width = w
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

if input_text:
    structured_items, parsed_experts = parse_full_email(input_text)
    
    if parsed_experts:
        st.success(f"✨ {len(parsed_experts)}名のエキスパート情報を正常に整理しました！")
        
        st.markdown("### 📧 メール送信用整形テキスト（ドラッグ選択してそのままコピーしてください）")
        email_html = generate_formatted_email_html(structured_items)
        st.markdown(f'<div class="email-preview-box">{email_html}</div>', unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        excel_data = generate_excel_bytes(parsed_experts)
        st.download_button(
            label="📥 エクセルファイルをダウンロード",
            data=excel_data,
            file_name="W8LY_Expert_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("エキスパート情報（#番号 で始まる記述）が検出されませんでした。入力内容をご確認ください。")
