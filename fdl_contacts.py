import json, re, os

_DEFAULT_CONTACTS = os.path.join(os.path.dirname(__file__), 'ar_contacts.js')


def get_customer_emails(sold_to: str, contacts_path: str = None) -> list:
    """Return email addresses for the given sold-to party number."""
    path = contacts_path or _DEFAULT_CONTACTS
    with open(path, encoding='utf-8') as f:
        content = f.read()
    m = re.search(r'const CONTACTS\s*=\s*(\{.*?\});', content, re.DOTALL)
    if not m:
        return []
    data = json.loads(m.group(1))
    key = sold_to.lstrip('0') or '0'
    entry = data.get(key) or data.get(sold_to)
    if not entry:
        return []
    return entry.get('emails', [])
