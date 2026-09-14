import os, shutil
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date
from typing import Optional
from docx import Document
from docx.oxml.ns import qn

_TEMPLATES = {
    '0038': os.path.join(os.path.dirname(__file__), 'Final demand letter Template - CN 0038.docx'),
    '0073': os.path.join(os.path.dirname(__file__), 'Final demand letter Template - TW 0073.docx'),
    '0074': os.path.join(os.path.dirname(__file__), 'Final demand letter Template - TW 0074.docx'),
}


@dataclass
class FDLInput:
    company_code: str
    customer_name_cn: str
    customer_name_en: str
    sold_to_party: str
    currency: str
    total_amount: float
    deadline: date
    letter_date: date
    collector_name: str
    collector_phone: str
    collector_email: str
    attention: str
    customer_address: str
    invoice_rows: list = field(default_factory=list)


def _fmt_date_cn(d) -> str:
    return d.strftime('%Y年%m月%d日') if d else ''


def _fmt_date_slash(d) -> str:
    return d.strftime('%-m/%-d/%Y') if d else ''


def _fmt_date_en(d) -> str:
    return d.strftime('%-d %B %Y') if d else ''


def _fmt_amount(amount: float) -> str:
    return f'{amount:,.2f}'


def _replace_in_para(para, placeholder: str, value: str):
    full = ''.join(r.text for r in para.runs)
    if placeholder not in full:
        return
    new = full.replace(placeholder, value)
    if para.runs:
        para.runs[0].text = new
        for r in para.runs[1:]:
            r.text = ''
    else:
        para.text = new


def _replace_all_paras(doc: Document, replacements: dict):
    all_paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paras.extend(cell.paragraphs)
    for para in all_paras:
        for ph, val in replacements.items():
            _replace_in_para(para, ph, str(val))


def _replace_context_aware_0074(doc: Document, inp: FDLInput):
    """0074: [Date] appears twice and [    ] appears twice — replace by context."""
    date_letter = _fmt_date_en(inp.letter_date)
    date_deadline = _fmt_date_en(inp.deadline)
    for para in doc.paragraphs:
        txt = ''.join(r.text for r in para.runs)
        if txt.strip() == '[Date]':
            _replace_in_para(para, '[Date]', date_letter)
        elif '[Date]' in txt and 'look into this matter' in txt.lower():
            _replace_in_para(para, '[Date]', date_deadline)
        elif '[    ]' in txt and 'ERP' in txt:
            _replace_in_para(para, '[    ]', inp.sold_to_party)
        elif '[    ]' in txt and 'Attention' in txt:
            _replace_in_para(para, '[    ]', inp.attention)


def _set_cell_text(cell, value: str):
    paras = cell.paragraphs
    if paras and paras[0].runs:
        paras[0].runs[0].text = value
        for r in paras[0].runs[1:]:
            r.text = ''
    elif paras:
        paras[0].add_run(value)


def _rebuild_table(doc: Document, inp: FDLInput):
    """Replace sample data rows with aggregated invoice rows."""
    table = doc.tables[0]
    # Save format reference from first data row before removing it
    sample_tr = deepcopy(table.rows[1]._tr)
    # Remove all rows between header (0) and total (last)
    while len(table.rows) > 2:
        tr = table.rows[1]._tr
        tr.getparent().remove(tr)
    total_tr = table.rows[-1]._tr

    for row_data in inp.invoice_rows:
        new_tr = deepcopy(sample_tr)
        # Clear all run text in the new row
        for tc in new_tr.findall('.//' + qn('w:tc')):
            for r_elem in tc.findall('.//' + qn('w:r')):
                for t_elem in r_elem.findall('.//' + qn('w:t')):
                    t_elem.text = ''
        total_tr.getparent().insert(
            list(total_tr.getparent()).index(total_tr), new_tr
        )
        new_row = table.rows[-2]
        cells = new_row.cells
        inv_date = row_data.get('invoice_date')
        due_date = row_data.get('due_date')
        values = [
            str(row_data.get('invoice_no', '')),
            _fmt_date_slash(inv_date) if inv_date else '',
            _fmt_date_slash(due_date) if due_date else '',
            str(row_data.get('overdue_days', '')),
            _fmt_amount(row_data.get('amount', 0)),
            str(row_data.get('description', '')),
        ]
        for i, val in enumerate(values):
            if i < len(cells):
                _set_cell_text(cells[i], val)

    # Update total row amount cell (col 4)
    total_cells = table.rows[-1].cells
    if len(total_cells) > 4:
        _set_cell_text(total_cells[4], _fmt_amount(inp.total_amount))


def fill_template(fdl_input: FDLInput, output_path: str) -> str:
    """Fill the appropriate DOCX template with fdl_input data. Returns output_path."""
    cc = fdl_input.company_code
    if cc not in _TEMPLATES:
        raise ValueError(f"Unknown company code '{cc}'. Must be 0038, 0073, or 0074.")

    shutil.copy2(_TEMPLATES[cc], output_path)
    doc = Document(output_path)

    if cc == '0038':
        replacements = {
            '[客户名称]': fdl_input.customer_name_cn,
            '[Currency]': fdl_input.currency,
            '[合计金额]': _fmt_amount(fdl_input.total_amount),
            '[截止日期]': _fmt_date_cn(fdl_input.deadline),
            '[Collector姓名]': fdl_input.collector_name,
            '[Collector电话]': fdl_input.collector_phone,
            '[Collector Email]': fdl_input.collector_email,
            '[日期]': _fmt_date_cn(fdl_input.letter_date),
        }
        _replace_all_paras(doc, replacements)
        for para in doc.paragraphs:
            txt = ''.join(r.text for r in para.runs)
            if '[    ]' in txt and 'ERP' in txt:
                _replace_in_para(para, '[    ]', fdl_input.sold_to_party)

    elif cc == '0073':
        replacements = {
            '[日期]': _fmt_date_cn(fdl_input.letter_date),
            '[客戶名稱]': fdl_input.customer_name_cn,
            '[ERP編號]': fdl_input.sold_to_party,
            '[Currency]': fdl_input.currency,
            '[合計金額]': _fmt_amount(fdl_input.total_amount),
            '[截止日期]': _fmt_date_cn(fdl_input.deadline),
            '[Collector姓名]': fdl_input.collector_name,
            '[Collector電話]': fdl_input.collector_phone,
            '[Collector Email]': fdl_input.collector_email,
        }
        _replace_all_paras(doc, replacements)

    elif cc == '0074':
        _replace_context_aware_0074(doc, fdl_input)
        replacements = {
            '[Customer Name]': fdl_input.customer_name_en,
            '[Customer Address]': fdl_input.customer_address,
            '[Currency]': fdl_input.currency,
            '[Total Outstanding Amount]': _fmt_amount(fdl_input.total_amount),
            '[Collector Name]': fdl_input.collector_name,
            '[Collector Phone]': fdl_input.collector_phone,
            '[Collector Email]': fdl_input.collector_email,
        }
        _replace_all_paras(doc, replacements)

    _rebuild_table(doc, fdl_input)
    doc.save(output_path)
    return output_path
