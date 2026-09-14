import os, sys
from datetime import date
import pytest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from fdl_template import FDLInput, fill_template
from docx import Document


@pytest.fixture
def base_input_0038():
    return FDLInput(
        company_code='0038',
        customer_name_cn='天能电池集团股份有限公司',
        customer_name_en='Tianneng Battery Group Co., Ltd.',
        sold_to_party='1302690',
        currency='CNY',
        total_amount=4161964124.0,
        deadline=date(2026, 9, 30),
        letter_date=date(2026, 9, 10),
        collector_name='Li, Lisa',
        collector_phone='021-22213000',
        collector_email='lisa.li02@sap.com',
        attention='',
        customer_address='',
        invoice_rows=[
            {'invoice_no': '26312000004548751336', 'invoice_date': date(2026, 7, 1),
             'due_date': date(2026, 7, 31), 'overdue_days': 41,
             'amount': 4161964124.0, 'description': 'CL_26.07.2026'},
        ],
    )


def test_fill_0038_replaces_customer_name(base_input_0038, tmp_path):
    out = fill_template(base_input_0038, str(tmp_path / 'out.docx'))
    doc = Document(out)
    full_text = '\n'.join(p.text for p in doc.paragraphs)
    assert '天能电池集团股份有限公司' in full_text
    assert '[客户名称]' not in full_text


def test_fill_0038_replaces_erp(base_input_0038, tmp_path):
    out = fill_template(base_input_0038, str(tmp_path / 'out.docx'))
    doc = Document(out)
    full_text = '\n'.join(p.text for p in doc.paragraphs)
    assert '1302690' in full_text


def test_fill_0038_no_remaining_placeholders(base_input_0038, tmp_path):
    import re
    out = fill_template(base_input_0038, str(tmp_path / 'out.docx'))
    doc = Document(out)
    full_text = '\n'.join(p.text for p in doc.paragraphs)
    remaining = re.findall(r'\[[^\]]{1,40}\]', full_text)
    remaining = [r for r in remaining if r not in ('[人民币/元]', '[Currency]')]
    assert remaining == [], f"Unreplaced placeholders: {remaining}"


def test_fill_0038_table_has_data_row(base_input_0038, tmp_path):
    out = fill_template(base_input_0038, str(tmp_path / 'out.docx'))
    doc = Document(out)
    table = doc.tables[0]
    assert len(table.rows) == 3
    data_row = table.rows[1]
    assert '26312000004548751336' in data_row.cells[0].text


def test_fill_0074_context_aware_date(tmp_path):
    inp = FDLInput(
        company_code='0074',
        customer_name_cn='',
        customer_name_en='Cheerplan (HK) Limited',
        sold_to_party='3087470',
        currency='HKD',
        total_amount=500000.0,
        deadline=date(2026, 9, 30),
        letter_date=date(2026, 9, 10),
        collector_name='Fan, Huang',
        collector_phone='+852-12345678',
        collector_email='fan.huang01@sap.com',
        attention='Accounts Payable',
        customer_address='Suite 100, Hong Kong',
        invoice_rows=[
            {'invoice_no': '10074260001', 'invoice_date': date(2026, 8, 1),
             'due_date': date(2026, 8, 31), 'overdue_days': 10,
             'amount': 500000.0, 'description': 'Cloud Services'},
        ],
    )
    out = fill_template(inp, str(tmp_path / 'out0074.docx'))
    doc = Document(out)
    full_text = '\n'.join(p.text for p in doc.paragraphs)
    assert '10 September 2026' in full_text
    assert '30 September 2026' in full_text
    assert '[Date]' not in full_text
