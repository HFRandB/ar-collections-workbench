from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional
import openpyxl

# Column indices (0-based) in Master Report sheet, header on row index 1
_C = {
    'company_code': 2, 'lob': 3, 'sold_to_number': 17, 'sold_to_name': 18,
    'collector_name': 19, 'doc_number': 20, 'ref_doc_number': 21, 'text': 22,
    'local_currency_amt': 23, 'local_currency': 24,
    'posting_date': 27, 'net_due_date': 28, 'fapiao_date': 53,
}
_HEADER_ROW_IDX = 1  # 0-indexed


@dataclass
class MasterReportRow:
    company_code: str
    lob: str
    sold_to_number: str
    sold_to_name: str
    collector_name: str
    doc_number: str
    ref_doc_number: str
    text: str
    local_currency_amt: float
    local_currency: str
    posting_date: Optional[date]
    net_due_date: Optional[date]
    fapiao_date: Optional[date]


@dataclass
class AggregatedLOB:
    company_code: str
    lob: str
    sold_to_number: str
    sold_to_name: str
    collector_name: str
    total_amount: float
    local_currency: str
    earliest_posting_date: Optional[date]
    earliest_net_due_date: Optional[date]
    earliest_fapiao_date: Optional[date]
    doc_numbers: list = field(default_factory=list)
    ref_doc_numbers: list = field(default_factory=list)
    texts: list = field(default_factory=list)
    rows: list = field(default_factory=list)  # original MasterReportRow list


def _to_date(val) -> Optional[date]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    return None


def _str(val) -> str:
    if val is None:
        return ''
    return str(val).strip()


def load_and_filter(path: str, sold_to: str, lob: str) -> list:
    """Load master report and return rows matching sold_to + lob."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Master Report']
    sold_to_norm = sold_to.lstrip('0') or '0'
    rows = []
    for idx, row in enumerate(ws.iter_rows(values_only=True)):
        if idx <= _HEADER_ROW_IDX:
            continue
        if row[_C['company_code']] is None:
            continue
        row_sold_to = _str(row[_C['sold_to_number']]).lstrip('0') or '0'
        row_lob = _str(row[_C['lob']])
        if row_sold_to != sold_to_norm:
            continue
        if row_lob.upper() != lob.upper():
            continue
        amt = row[_C['local_currency_amt']]
        if amt is None:
            continue
        rows.append(MasterReportRow(
            company_code=_str(row[_C['company_code']]),
            lob=row_lob,
            sold_to_number=_str(row[_C['sold_to_number']]),
            sold_to_name=_str(row[_C['sold_to_name']]),
            collector_name=_str(row[_C['collector_name']]),
            doc_number=_str(row[_C['doc_number']]),
            ref_doc_number=_str(row[_C['ref_doc_number']]),
            text=_str(row[_C['text']]),
            local_currency_amt=float(amt),
            local_currency=_str(row[_C['local_currency']]),
            posting_date=_to_date(row[_C['posting_date']]),
            net_due_date=_to_date(row[_C['net_due_date']]),
            fapiao_date=_to_date(row[_C['fapiao_date']]),
        ))
    return rows


def aggregate_rows(rows: list) -> 'AggregatedLOB':
    """Aggregate multiple document rows for the same Sold-to + LOB into one."""
    if not rows:
        raise ValueError("Cannot aggregate empty rows list")
    base = rows[0]
    total = sum(r.local_currency_amt for r in rows)
    posting_dates = [r.posting_date for r in rows if r.posting_date]
    earliest_posting = min(posting_dates) if posting_dates else None
    due_dates = [r.net_due_date for r in rows if r.net_due_date]
    earliest_due = min(due_dates) if due_dates else None
    fap_dates = [r.fapiao_date for r in rows if r.fapiao_date]
    earliest_fap = min(fap_dates) if fap_dates else None
    return AggregatedLOB(
        company_code=base.company_code,
        lob=base.lob,
        sold_to_number=base.sold_to_number,
        sold_to_name=base.sold_to_name,
        collector_name=base.collector_name,
        total_amount=total,
        local_currency=base.local_currency,
        earliest_posting_date=earliest_posting,
        earliest_net_due_date=earliest_due,
        earliest_fapiao_date=earliest_fap,
        doc_numbers=[r.doc_number for r in rows],
        ref_doc_numbers=[r.ref_doc_number for r in rows],
        texts=[r.text for r in rows],
        rows=rows,
    )
