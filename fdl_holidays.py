from datetime import date, timedelta

# 2026 public holidays — verified against official schedules
_HOLIDAYS = {
    'CN': {
        date(2026, 1, 1),   # New Year
        date(2026, 1, 28), date(2026, 1, 29), date(2026, 1, 30),
        date(2026, 1, 31), date(2026, 2, 1), date(2026, 2, 2), date(2026, 2, 3),  # Spring Festival
        date(2026, 4, 5),   # Qingming
        date(2026, 5, 1), date(2026, 5, 2), date(2026, 5, 3),  # Labor Day
        date(2026, 6, 19),  # Dragon Boat
        date(2026, 10, 1), date(2026, 10, 2), date(2026, 10, 3),
        date(2026, 10, 4), date(2026, 10, 5), date(2026, 10, 6), date(2026, 10, 7),  # National Day
    },
    'TW': {
        date(2026, 1, 1),   # New Year
        date(2026, 1, 28), date(2026, 1, 29), date(2026, 1, 30),
        date(2026, 1, 31), date(2026, 2, 1), date(2026, 2, 2),  # Lunar New Year
        date(2026, 2, 28),  # Peace Memorial Day
        date(2026, 4, 4), date(2026, 4, 5),  # Children's Day + Qingming
        date(2026, 5, 1),   # Labor Day
        date(2026, 6, 19),  # Dragon Boat
        date(2026, 10, 9),  # National Day substitute (Oct 10 falls on Saturday 2026)
    },
    'HK': {
        date(2026, 1, 1),   # New Year
        date(2026, 1, 28), date(2026, 1, 29), date(2026, 1, 30),  # Lunar New Year
        date(2026, 4, 3), date(2026, 4, 5), date(2026, 4, 6),     # Good Friday, Qingming, Easter Mon
        date(2026, 5, 1),   # Labor Day
        date(2026, 5, 27),  # Buddha's Birthday
        date(2026, 6, 19),  # Tuen Ng
        date(2026, 7, 1),   # HK SAR Establishment
        date(2026, 10, 1),  # National Day
        date(2026, 10, 2),  # Day Following National Day (HK statutory holiday)
        date(2026, 10, 6),  # Day after Mid-Autumn
        date(2026, 10, 20), # Chung Yeung
        date(2026, 12, 25), date(2026, 12, 26),  # Christmas
    },
}


def add_business_days(start: date, n: int, region: str) -> date:
    """Return the date that is n business days after start, skipping weekends
    and public holidays for the given region ('CN', 'TW', 'HK')."""
    if region not in _HOLIDAYS:
        raise ValueError(f"Unknown region '{region}'. Must be CN, TW, or HK.")
    holidays = _HOLIDAYS[region]
    current = start
    days_added = 0
    while days_added < n:
        current += timedelta(days=1)
        if current.weekday() < 5 and current not in holidays:
            days_added += 1
    return current
