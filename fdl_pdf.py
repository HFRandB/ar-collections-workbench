import os
from docx2pdf import convert

_CC_LABELS = {'0038': 'CN 0038', '0073': 'TW 0073', '0074': 'HK 0074'}


def build_pdf_filename(company_code: str, sold_to: str, name_en: str, name_cn: str) -> str:
    """Build the output PDF filename matching the naming convention."""
    label = _CC_LABELS.get(company_code, company_code)
    cn_part = f' {name_cn}' if name_cn else ''
    if company_code in ('0038', '0073'):
        return f'最终催款通知 - {sold_to} {name_en}{cn_part} - {label}.pdf'
    else:
        return f'Final Demand Notice - {sold_to} {name_en}{cn_part} - {label}.pdf'


def generate_pdf(docx_path: str, output_dir: str = '/tmp/fdl_output') -> str:
    """Convert a filled DOCX to PDF using docx2pdf (Word on macOS). Returns PDF path."""
    os.makedirs(output_dir, exist_ok=True)
    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    pdf_path = os.path.join(output_dir, pdf_name)
    convert(docx_path, pdf_path)
    return pdf_path
