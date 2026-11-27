"""Creator attribution for live file and folder observations during deep scan."""
import csv
from collections import Counter
from pathlib import Path
from uuid import UUID
from functools import lru_cache

COLUMNS = ('User Email', 'Files Created', 'Folders Created', 'Total Items Created')
USERS_URL = 'https://graph.microsoft.com/v1.0/users?$select=id,mail,userPrincipalName&$top=999'


def creator_counts_enabled(config, data):
    return (getattr(config, 'user_source', None) == 'tenant'
            and not getattr(config, 'shallow_scan', False)
            and not any(data.get(key) for key in ('emailIds', 'siteUrls', 'drives')))


@lru_cache(maxsize=8192)
def _guid(value):
    try:
        return str(UUID(value))
    except ValueError:
        return ''


def creator_key(item):
    identity = item.get('createdBy')
    if not isinstance(identity, dict):
        return ('unknown', '')
    user = identity.get('user')
    if isinstance(user, dict) and user:
        user_id = user.get('id')
        if isinstance(user_id, str):
            normalized = _guid(user_id)
            if normalized:
                return ('user', normalized)
        return ('unresolved_user', '')
    if identity.get('siteUser'):
        return ('unresolved_user', '')
    app = identity.get('application')
    if isinstance(app, dict) and app:
        app_id = app.get('id')
        return ('application', _guid(app_id) if isinstance(app_id, str) else '')
    if any(identity.get(key) for key in ('device', 'group', 'sharePointGroup')):
        return ('other', '')
    return ('unknown', '')


def _email(value):
    # Reject guest sign-in aliases and values that could be interpreted as
    # spreadsheet formulas; never manufacture an email from a display name.
    if not isinstance(value, str):
        return None
    value = value.strip().lower()
    if (not value or value[0] in '=+-@' or '#ext#' in value
            or any(ch.isspace() or ord(ch) < 32 for ch in value)
            or value.count('@') != 1):
        return None
    local, domain = value.split('@')
    return value if local and domain else None


def resolve_creator_emails(counts, fetch_page, should_stop, logger):
    """Resolve observed directory IDs once, using paginated user roster reads.

    Keep only matching addresses in memory; any permission, paging or service
    failure is non-fatal and leaves remaining creators unresolved.
    """
    pending = {creator for kind, creator, _ in counts if kind == 'user'}
    resolved = {}
    url = USERS_URL
    seen_links = set()
    try:
        while pending and url and not should_stop():
            # Only follow Graph user-roster links and never repeat a page.
            if (not isinstance(url, str)
                    or not url.startswith('https://graph.microsoft.com/v1.0/users?')
                    or url in seen_links):
                raise ValueError('Invalid or repeated directory pagination link')
            seen_links.add(url)
            page = fetch_page(url)
            if not isinstance(page, dict) or not isinstance(page.get('value'), list):
                raise ValueError('Invalid directory response')
            for user in page['value']:
                if not isinstance(user, dict):
                    continue
                user_id = user.get('id')
                user_id = _guid(user_id) if isinstance(user_id, str) else ''
                if user_id not in pending:
                    continue
                email = _email(user.get('mail')) or _email(user.get('userPrincipalName'))
                if email:
                    resolved[user_id] = email
                pending.remove(user_id)
            url = page.get('@odata.nextLink')
    except Exception as exc:
        logger(f'[WARNING] Creator email resolution incomplete: {type(exc).__name__}. '
               'Check the documented User.Read.All permission and scan logs.')
    return resolved


def creator_report(counts, expected_drives, completed_drives, partial, emails=None):
    emails = emails or {}
    by_email = Counter()
    total = sum(counts.values())
    attributed = 0
    with_guid = 0
    for (kind, creator, facet), count in counts.items():
        if kind != 'user':
            continue
        with_guid += count
        email = emails.get(creator)
        if email:
            by_email[email, facet] += count
            attributed += count
    rows = []
    for email in sorted({email for email, _ in by_email}):
        files = by_email[email, 'file']
        folders = by_email[email, 'folder']
        rows.append(dict(zip(COLUMNS, (email, files, folders, files + folders))))
    rows.sort(key=lambda row: -row['Total Items Created'])
    return {'rows': rows, 'coverage': {
        'status': 'partial' if partial else 'completed_no_recorded_failures',
        'discovered_drives': expected_drives,
        'completed_drives': completed_drives,
        'observed_items': total,
        'items_with_user_guid': with_guid,
        'items_with_user_email': attributed,
        'unreported_items': total - attributed,
        'unresolved_user_items': with_guid - attributed,
        'non_directory_user_items': total - with_guid,
    }}


def write_creator_report(output_dir, report, timestamp):
    directory = Path(output_dir)
    path = directory / f'creator_counts_{timestamp}.csv'
    with path.open('w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(report['rows'])
    return path


def is_terminal_page(response):
    status = response.get('status')
    body = response.get('body')
    return (isinstance(status, int) and 200 <= status < 300
            and isinstance(body, dict) and isinstance(body.get('value'), list)
            and '@odata.nextLink' not in body)
