import json, os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from fdl_contacts import get_customer_emails


def _make_contacts_file(data: dict) -> str:
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.js', delete=False)
    tmp.write(f'const CONTACTS = {json.dumps(data)};\n')
    tmp.close()
    return tmp.name


def test_get_emails_returns_list_for_known_sold_to():
    path = _make_contacts_file({
        "1302690": {"emails": ["a@example.com", "b@example.com"], "names": ["A"]},
    })
    result = get_customer_emails("1302690", path)
    assert result == ["a@example.com", "b@example.com"]


def test_get_emails_returns_empty_for_unknown():
    path = _make_contacts_file({"999": {"emails": ["x@y.com"], "names": []}})
    assert get_customer_emails("1234567", path) == []


def test_get_emails_strips_leading_zeros_on_lookup():
    path = _make_contacts_file({"195218": {"emails": ["j@tsmc.com"], "names": []}})
    assert get_customer_emails("0195218", path) == ["j@tsmc.com"]
    assert get_customer_emails("195218", path) == ["j@tsmc.com"]
