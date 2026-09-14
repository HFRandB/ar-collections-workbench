import os, shutil, subprocess, tempfile

_CC_LABELS = {'0038': 'CN 0038', '0073': 'TW 0073', '0074': 'HK 0074'}


def build_pdf_filename(company_code: str, sold_to: str, name_en: str, name_cn: str) -> str:
    """Build the output PDF filename matching the naming convention."""
    label = _CC_LABELS.get(company_code, company_code)
    cn_part = f' {name_cn}' if name_cn else ''
    if company_code in ('0038', '0073'):
        return f'最终催款通知 - {sold_to} {name_en}{cn_part} - {label}.pdf'
    else:
        return f'Final Demand Notice - {sold_to} {name_en}{cn_part} - {label}.pdf'


def _convert_via_applescript(docx_path: str, pdf_path: str):
    """Convert DOCX to PDF using macOS AppleScript + Microsoft Word."""
    script = f'''
tell application "Microsoft Word"
    activate
    open "{docx_path}"
    delay 2
    set theDoc to active document
    save as theDoc file name "{pdf_path}" file format format PDF
    close theDoc saving no
end tell
'''
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript PDF conversion failed: {result.stderr.strip()}")


def generate_pdf(docx_path: str, output_dir: str = '/tmp/fdl_output') -> str:
    """Convert a filled DOCX to PDF using Microsoft Word on macOS. Returns PDF path.

    Uses AppleScript directly to avoid docx2pdf JXA compatibility issues.
    Converts through a temp ASCII-named copy to handle non-ASCII paths.
    """
    os.makedirs(output_dir, exist_ok=True)
    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    final_pdf = os.path.join(output_dir, pdf_name)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_docx = os.path.join(tmp, 'fdl_tmp.docx')
        tmp_pdf = os.path.join(tmp, 'fdl_tmp.pdf')
        shutil.copy2(docx_path, tmp_docx)
        _convert_via_applescript(tmp_docx, tmp_pdf)
        shutil.move(tmp_pdf, final_pdf)

    return final_pdf
