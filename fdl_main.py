#!/usr/bin/env python3
"""
FDL Generator CLI — called by Claude Code to generate Final Demand Letters.

Phase 1: Extract data from master report
  python3 fdl_main.py --phase data --sold-to 1302690 --lob CL \\
    --master-report path/to/Master_Report_GC_*.xlsx

Phase 2: Generate DOCX + PDF with complete data
  python3 fdl_main.py --phase generate \\
    --sold-to 1302690 --lob CL \\
    --collector-name "Li, Lisa" --collector-phone "021-22213000" \\
    --collector-email "lisa.li02@sap.com" \\
    --customer-name-cn "天能电池集团股份有限公司" \\
    --fapiao-numbers "26312000004548751336;26312000004548751337" \\
    --fapiao-dates "2026-07-01;2026-07-02" \\
    [--attention "Accounts Payable"] \\
    [--customer-address "Suite 100, HK"] \\
    [--deadline 2026-09-30] \\
    [--master-report path/to/report.xlsx]

Claude Code asks the user to provide the Master Report file at the start of each session.
"""
import argparse, glob, json, os, sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(__file__))
from fdl_data import load_and_filter, aggregate_rows
from fdl_holidays import add_business_days
from fdl_contacts import get_customer_emails
from fdl_template import FDLInput, fill_template
from fdl_pdf import build_pdf_filename, generate_pdf


def find_master_report() -> str:
    # Check ~/Downloads as fallback (user may have placed it there)
    dl_files = sorted(
        glob.glob(os.path.expanduser('~/Downloads/Master Report_GC_*.xlsx')),
        key=os.path.getmtime, reverse=True
    )
    if dl_files:
        return dl_files[0]
    raise FileNotFoundError(
        "Master Report not found. Please share the Master Report_GC_*.xlsx file "
        "(drag it into the chat or provide the file path)."
    )


def _region_for_cc(cc: str) -> str:
    return {'0038': 'CN', '0073': 'TW', '0074': 'HK'}.get(cc, 'CN')


def phase_data(args) -> dict:
    report_path = args.master_report or find_master_report()
    rows = load_and_filter(report_path, args.sold_to, args.lob)
    if not rows:
        return {
            'status': 'no_data',
            'message': f"No rows found for sold-to={args.sold_to}, lob={args.lob}",
        }
    agg = aggregate_rows(rows)
    cc = agg.company_code
    region = _region_for_cc(cc)
    today = date.today()
    deadline = add_business_days(today, 14, region)
    od_days = (today - agg.earliest_net_due_date).days if agg.earliest_net_due_date else 0
    per_row = [
        {
            'doc_number': r.doc_number,
            'ref_doc_number': r.ref_doc_number,
            'amount': r.local_currency_amt,
            'posting_date': r.posting_date.isoformat() if r.posting_date else None,
            'net_due_date': r.net_due_date.isoformat() if r.net_due_date else None,
            'fapiao_date': r.fapiao_date.isoformat() if r.fapiao_date else None,
            'text': r.text,
        }
        for r in agg.rows
    ]
    return {
        'status': 'ok',
        'company_code': cc,
        'region': region,
        'sold_to': args.sold_to,
        'lob': args.lob,
        'customer_name_en': agg.sold_to_name,
        'collector_name_from_report': agg.collector_name,
        'currency': agg.local_currency,
        'total_amount': agg.total_amount,
        'doc_numbers': agg.doc_numbers,
        'ref_doc_numbers': agg.ref_doc_numbers,
        'per_row_data': per_row,
        'earliest_posting_date': agg.earliest_posting_date.isoformat() if agg.earliest_posting_date else None,
        'earliest_net_due_date': agg.earliest_net_due_date.isoformat() if agg.earliest_net_due_date else None,
        'earliest_fapiao_date': agg.earliest_fapiao_date.isoformat() if agg.earliest_fapiao_date else None,
        'overdue_days': od_days,
        'deadline': deadline.isoformat(),
        'customer_emails': get_customer_emails(args.sold_to),
        'next_step': {
            '0038': 'Search SharePoint VAT folder for fapiao files matching doc_numbers; filenames are dzfp_{number}_{ISSUER_name}_{ts}.xml — the middle name is SAP issuer entity, NOT customer. Use web search on customer_name_en to find Chinese customer name. Pass one --fapiao-numbers, --fapiao-dates, --fapiao-amounts entry per row using per_row_data amounts.',
            '0073': 'Web-search customer_name_en to find Traditional Chinese company name; per_row_data has one entry per ref_doc_number with individual amounts.',
            '0074': 'Web-search customer_name_en to find Chinese company name; per_row_data has one entry per doc_number with individual amounts.',
        }.get(cc, ''),
    }


def phase_generate(args) -> dict:
    report_path = args.master_report or find_master_report()
    rows = load_and_filter(report_path, args.sold_to, args.lob)
    if not rows:
        return {'status': 'no_data', 'message': f"No rows found for sold-to={args.sold_to}, lob={args.lob}"}
    agg = aggregate_rows(rows)
    cc = agg.company_code
    region = _region_for_cc(cc)
    today = date.today()
    deadline_str = args.deadline or add_business_days(today, 14, region).isoformat()
    deadline = date.fromisoformat(deadline_str)

    fapiao_numbers = [x.strip() for x in (args.fapiao_numbers or '').split(';') if x.strip()]
    fapiao_dates = [x.strip() for x in (args.fapiao_dates or '').split(';') if x.strip()]
    fapiao_amounts = [x.strip() for x in (args.fapiao_amounts or '').split(';') if x.strip()]

    if cc == '0038':
        if fapiao_numbers:
            # Build one row per fapiao number
            # Amounts: use --fapiao-amounts if provided, else distribute from master report rows 1:1,
            # else fall back to equal share of total
            row_amounts = agg.rows  # individual master report rows
            invoice_rows = []
            for i, fap_no in enumerate(fapiao_numbers):
                inv_date = date.fromisoformat(fapiao_dates[i]) if i < len(fapiao_dates) else (agg.earliest_fapiao_date or today)
                due = inv_date + timedelta(days=30)
                od = max(0, (today - due).days)
                if fapiao_amounts and i < len(fapiao_amounts):
                    amt = float(fapiao_amounts[i].replace(',', ''))
                elif len(row_amounts) == len(fapiao_numbers):
                    amt = row_amounts[i].local_currency_amt
                else:
                    amt = agg.total_amount / len(fapiao_numbers)
                invoice_rows.append({
                    'invoice_no': fap_no,
                    'invoice_date': inv_date,
                    'due_date': due,
                    'overdue_days': od,
                    'amount': amt,
                    'description': agg.texts[i] if i < len(agg.texts) else '',
                })
        else:
            inv_date = agg.earliest_fapiao_date or today
            due = agg.earliest_net_due_date or today
            invoice_rows = [{
                'invoice_no': 'N/A',
                'invoice_date': inv_date,
                'due_date': due,
                'overdue_days': max(0, (today - due).days),
                'amount': agg.total_amount,
                'description': agg.texts[0] if agg.texts else '',
            }]
    elif cc == '0073':
        invoice_rows = []
        for r in agg.rows:
            due = r.net_due_date or today
            invoice_rows.append({
                'invoice_no': r.ref_doc_number,
                'invoice_date': r.posting_date or today,
                'due_date': due,
                'overdue_days': max(0, (today - due).days),
                'amount': r.local_currency_amt,
                'description': r.text,
            })
    else:  # 0074
        invoice_rows = []
        for r in agg.rows:
            due = r.net_due_date or today
            invoice_rows.append({
                'invoice_no': r.doc_number,
                'invoice_date': r.posting_date or today,
                'due_date': due,
                'overdue_days': max(0, (today - due).days),
                'amount': r.local_currency_amt,
                'description': r.text,
            })

    os.makedirs('/tmp/fdl_output', exist_ok=True)
    name_cn = args.customer_name_cn or ''
    pdf_filename = build_pdf_filename(cc, args.sold_to, agg.sold_to_name, name_cn)
    docx_path = f'/tmp/fdl_output/{os.path.splitext(pdf_filename)[0]}.docx'

    fdl_input = FDLInput(
        company_code=cc,
        customer_name_cn=name_cn,
        customer_name_en=agg.sold_to_name,
        sold_to_party=args.sold_to,
        currency=agg.local_currency,
        total_amount=agg.total_amount,
        deadline=deadline,
        letter_date=today,
        collector_name=args.collector_name,
        collector_phone=args.collector_phone,
        collector_email=args.collector_email,
        attention=args.attention or '',
        customer_address=args.customer_address or '',
        invoice_rows=invoice_rows,
    )

    fill_template(fdl_input, docx_path)
    pdf_path = generate_pdf(docx_path)

    _BODIES = {
        '0038': (
            "尊敬的客户：\n您好！\n截至目前，我司仍未收到贵司应付账款。请查收附件《最终催款通知》。\n"
            "请贵司尽快安排付款，并于收到本通知后及时与我们确认付款时间。"
            "如贵司已完成付款，请忽略此通知，并将付款凭证发送给我们以便核对。\n"
            "感谢贵司的配合与支持，期待尽快收到您的回复。\n此致\n敬礼！\n\n"
            "思爱普（中国）有限公司\n全球财务共享服务中心（Cash Collections,Global Finance Shared Services）"
        ),
        '0073': (
            "尊敬的客戶：\n您好！\n截至目前，本公司仍未收到貴公司應付帳款。請查閱附件《最終催款通知》。\n"
            "請貴公司盡快安排付款，並於收到本通知後及時與我們確認付款時間。"
            "如貴公司已完成付款，請忽略此通知，並將付款憑證發送給我們以便核對。\n"
            "感謝貴公司的配合與支持，期待盡快收到您的回覆。\n此致\n敬禮！\n\n"
            "SAP Taiwan Co., Ltd. 思愛普軟體系統股份有限公司\n"
            "Cash Collections, Global Finance Shared Services"
        ),
        '0074': (
            "Dear Customer,\n\nPlease find attached the Final Demand Notice for outstanding invoices.\n\n"
            "We kindly request that you arrange payment at your earliest convenience and confirm the "
            "payment date upon receipt of this notice. If payment has already been made, please disregard "
            "this notice and send us the payment confirmation for our records.\n\n"
            "Thank you for your cooperation. We look forward to hearing from you soon.\n\n"
            "Yours faithfully,\n\nSAP Hong Kong Co. Ltd.\nCash Collections, Global Finance Shared Services"
        ),
    }
    cn_part = f' {name_cn}' if name_cn else ''
    if cc in ('0038', '0073'):
        subject = f'最終催款通知 - {args.sold_to} {agg.sold_to_name}{cn_part}' if cc == '0073' else f'最终催款通知 - {args.sold_to} {agg.sold_to_name}{cn_part}'
    else:
        subject = f'Final Demand Notice - {args.sold_to} {agg.sold_to_name}{cn_part}'

    return {
        'status': 'ok',
        'company_code': cc,
        'sold_to': args.sold_to,
        'docx_path': docx_path,
        'pdf_path': pdf_path,
        'email_subject': subject,
        'email_to': get_customer_emails(args.sold_to),
        'email_cc': [args.collector_email],
        'email_body': _BODIES[cc],
    }


def main():
    parser = argparse.ArgumentParser(description='FDL Generator — Phase 1 (data) or Phase 2 (generate)')
    parser.add_argument('--phase', choices=['data', 'generate'], required=True)
    parser.add_argument('--sold-to', required=True)
    parser.add_argument('--lob', required=True)
    parser.add_argument('--master-report', help='Path to Master Report_GC_*.xlsx (required if not in ~/Downloads/)')
    parser.add_argument('--collector-name', default='')
    parser.add_argument('--collector-phone', default='')
    parser.add_argument('--collector-email', default='')
    parser.add_argument('--customer-name-cn', default='')
    parser.add_argument('--fapiao-numbers', default='')
    parser.add_argument('--fapiao-dates', default='')
    parser.add_argument('--fapiao-amounts', default='', help='Per-fapiao amounts (semicolon-separated, 0038 only)')
    parser.add_argument('--attention', default='')
    parser.add_argument('--customer-address', default='')
    parser.add_argument('--deadline')
    args = parser.parse_args()

    if args.phase == 'data':
        result = phase_data(args)
    else:
        result = phase_generate(args)

    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == '__main__':
    main()
