import os, sys
from datetime import date, datetime
import openpyxl
import pytest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from fdl_data import load_and_filter, aggregate_rows, MasterReportRow

FIXTURE_PATH = os.path.join(os.path.dirname(__file__), 'fixtures', 'sample_master_report.xlsx')


def _make_fixture():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Master Report'
    ws.append([None] * 83)
    headers = [''] * 83
    headers[2] = 'Company Code'
    headers[3] = 'Line of Business'
    headers[17] = 'Sold-to Number'
    headers[18] = 'Sold-To Party Name'
    headers[19] = 'Collector\xa0Name'
    headers[20] = 'Document Number'
    headers[21] = 'Ref. document number'
    headers[22] = 'Text'
    headers[23] = 'Local Currency Amt'
    headers[24] = 'Local Currency'
    headers[27] = 'Posting Date'
    headers[28] = 'Net Due Date'
    headers[53] = 'Fapiao/Tax Invoice Date'
    ws.append(headers)
    # 0038 CL doc 1
    r1 = [''] * 83
    r1[2], r1[3], r1[17], r1[18] = '0038', 'CL', 1302690, 'Tianneng Battery Group'
    r1[19], r1[20], r1[21] = 'Gao, Camille', 990100001111, 0
    r1[22], r1[23], r1[24] = 'Desc1', 100000.0, 'CNY'
    r1[27] = datetime(2026, 7, 1)
    r1[28] = datetime(2026, 7, 31)
    r1[53] = datetime(2026, 7, 1)
    ws.append(r1)
    # 0038 CL doc 2 (same LOB — should aggregate)
    r2 = list(r1)
    r2[20], r2[23] = 990100002222, 50000.0
    r2[27] = datetime(2026, 6, 1)
    r2[28] = datetime(2026, 6, 30)
    r2[53] = datetime(2026, 6, 1)
    ws.append(r2)
    # different LOB SW
    r3 = list(r1)
    r3[3], r3[20], r3[23] = 'SW', 990100003333, 200000.0
    ws.append(r3)
    # different sold-to
    r4 = list(r1)
    r4[17] = 9999999
    ws.append(r4)
    wb.save(FIXTURE_PATH)


@pytest.fixture(autouse=True)
def ensure_fixture():
    if not os.path.exists(FIXTURE_PATH):
        _make_fixture()


def test_load_returns_matching_rows():
    rows = load_and_filter(FIXTURE_PATH, '1302690', 'CL')
    assert len(rows) == 2
    assert all(r.company_code == '0038' for r in rows)
    assert all(r.lob == 'CL' for r in rows)


def test_load_filters_out_other_lob():
    rows = load_and_filter(FIXTURE_PATH, '1302690', 'SW')
    assert len(rows) == 1
    assert rows[0].doc_number == '990100003333'


def test_load_returns_empty_for_unknown_sold_to():
    rows = load_and_filter(FIXTURE_PATH, '0000000', 'CL')
    assert rows == []


def test_aggregate_sums_amounts():
    rows = load_and_filter(FIXTURE_PATH, '1302690', 'CL')
    agg = aggregate_rows(rows)
    assert agg.total_amount == pytest.approx(150000.0)


def test_aggregate_uses_earliest_dates():
    rows = load_and_filter(FIXTURE_PATH, '1302690', 'CL')
    agg = aggregate_rows(rows)
    assert agg.earliest_posting_date == date(2026, 6, 1)
    assert agg.earliest_net_due_date == date(2026, 6, 30)


def test_aggregate_joins_doc_numbers():
    rows = load_and_filter(FIXTURE_PATH, '1302690', 'CL')
    agg = aggregate_rows(rows)
    assert '990100001111' in agg.doc_numbers
    assert '990100002222' in agg.doc_numbers


def test_aggregate_single_row():
    rows = load_and_filter(FIXTURE_PATH, '1302690', 'SW')
    agg = aggregate_rows(rows)
    assert agg.total_amount == pytest.approx(200000.0)
    assert agg.ref_doc_numbers == ['0']
