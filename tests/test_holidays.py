from datetime import date
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from fdl_holidays import add_business_days

def test_add_business_days_skips_weekend():
    # 2026-09-10 is Thursday; +2 business days skips Saturday → Monday 09-14
    assert add_business_days(date(2026, 9, 10), 2, 'CN') == date(2026, 9, 14)

def test_add_business_days_14_no_holidays():
    # 2026-09-10 (Thu) + 14 business days = 2026-09-30 (Wed), no CN holidays in range
    assert add_business_days(date(2026, 9, 10), 14, 'CN') == date(2026, 9, 30)

def test_add_business_days_skips_cn_national_day():
    # 2026-09-25 (Fri) + 5 business days: skip Oct 1-7 (National Day) → Oct 9 (Fri)
    assert add_business_days(date(2026, 9, 25), 5, 'CN') == date(2026, 10, 9)

def test_add_business_days_tw():
    # 2026-10-08 (Thu) + 2 business days: Oct 10 is TW National Day → skip → Oct 13 (Tue)
    assert add_business_days(date(2026, 10, 8), 2, 'TW') == date(2026, 10, 13)

def test_add_business_days_hk():
    # 2026-09-29 (Tue) + 2 business days: Oct 1 is HK National Day holiday → Oct 5 (Mon)
    assert add_business_days(date(2026, 9, 29), 2, 'HK') == date(2026, 10, 5)

def test_unknown_region_raises():
    import pytest
    with pytest.raises(ValueError, match="Unknown region"):
        add_business_days(date(2026, 9, 10), 5, 'XX')
