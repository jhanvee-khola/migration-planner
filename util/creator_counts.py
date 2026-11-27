"""Creator attribution for live file and folder observations during deep scan."""
import csv
from collections import Counter
from pathlib import Path
from uuid import UUID
from functools import lru_cache

COLUMNS = ('User Email', 'Files Created', 'Folders Created', 'Total Items Created')
USERS_URL = 'https://graph.microsoft.com/v1.0/users?$select=id,mail,userPrincipalName&$top=999'
DELETED_USERS_URL = ('https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user'
                     '?$select=id,mail,userPrincipalName&$top=999')


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


def _email(value):
    # Membership claims can expose the historical SharePoint sign-in address.
    # Other claim types and guest sign-in aliases are not email addresses.
    if not isinstance(value, str):
        return None
    value = value.strip().lower()
    prefix = 'i:0#.f|membership|'
    if value.startswith(prefix):
        value = value[len(prefix):]
    if (not value or value[0] in '=+-@' or '#ext#' in value or '|' in value
            or any(ch.isspace() or ord(ch) < 32 for ch in value)
            or value.count('@') != 1):
        return None
    local, domain = value.split('@')
    return value if local and domain else None


def _identity_email(user, site_user):
    for identity in (user, site_user):
        for field in ('mail', 'email', 'userPrincipalName', 'loginName'):
            email = _email(identity.get(field))
            if email:
                return email
    return None


def creator_key(item, metadata=None):
    """Use stable IDs, not display names, for aggregation.

    Site-local IDs must not be merged across unrelated site collections.
    Fall back to drive scope when Graph does not return a site ID.
    Metadata stores only one address and name per distinct creator.
    """
    identity = item.get('createdBy')
    if not isinstance(identity, dict):
        return ('unknown', '')
    user = identity.get('user')
    site_user = identity.get('siteUser')
    user = user if isinstance(user, dict) else {}
    site_user = site_user if isinstance(site_user, dict) else {}
    if user or site_user:
        user_id = user.get('id')
        guid = _guid(user_id) if isinstance(user_id, str) else ''
        if guid:
            key = ('user', guid)
        else:
            email = _identity_email(user, site_user)
            if email:
                key = ('email', email)
            else:
                parent = item.get('parentReference') or {}
                scope = parent.get('siteId') or parent.get('driveId')
                local_id = site_user.get('id') or user_id
                if scope and isinstance(local_id, str) and local_id:
                    key = ('site_user', (scope, local_id))
                else:
                    key = ('unresolved_user', '')
        if metadata is not None and key[0] in ('user', 'site_user', 'email'):
            entry = metadata.get(key)
            if entry is None:
                entry = {}
                metadata[key] = entry
            if 'email' not in entry:
                email = _identity_email(user, site_user)
                if email:
                    entry['email'] = email
            if 'name' not in entry:
                name = user.get('displayName') or site_user.get('displayName')
                if isinstance(name, str) and name.strip() and not _guid(name.strip()):
                    entry['name'] = ' '.join(name.split())[:200]
        return key
    if any(identity.get(field) for field in
           ('application', 'device', 'group', 'sharePointGroup', 'siteGroup')):
        return ('non_user', '')
    return ('unknown', '')


def resolve_creator_emails(counts, fetch_page, should_stop, logger, metadata=None):
    """Use observed addresses first, then at most one pass per directory roster."""
    metadata = metadata or {}
    ids = {creator for kind, creator, _ in counts if kind == 'user'}
    resolved = {creator: metadata['user', creator]['email'] for creator in ids
                if metadata.get(('user', creator), {}).get('email')}
    pending = ids - resolved.keys()

    def lookup(initial_url, deleted=False):
        url = initial_url
        prefix = initial_url.split('?', 1)[0] + '?'
        seen_links = set()
        while pending and url and not should_stop():
            if (not isinstance(url, str) or not url.startswith(prefix)
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
                email = _email(user.get('mail'))
                # Entra may rewrite a deleted account's UPN, so do not use it
                # as a historical email unless the original mail is available.
                if not email and not deleted:
                    email = _email(user.get('userPrincipalName'))
                if email:
                    resolved[user_id] = email
                pending.remove(user_id)
            url = page.get('@odata.nextLink')

    if pending and not should_stop():
        try:
            lookup(USERS_URL)
        except Exception as exc:
            logger(f'[WARNING] Creator live-directory resolution incomplete: {type(exc).__name__}. '
                   'Check the documented User.Read.All permission and scan logs.')
            # Avoid repeating a likely permission/service failure on another
            # endpoint; retain all unresolved counts in the report.
            return resolved
    if pending and not should_stop():
        try:
            lookup(DELETED_USERS_URL, deleted=True)
        except Exception as exc:
            logger(f'[WARNING] Creator deleted-directory resolution incomplete: {type(exc).__name__}.')
    return resolved


def creator_report(counts, expected_drives, completed_drives, partial, emails=None, metadata=None):
    emails = emails or {}
    metadata = metadata or {}
    by_label = Counter()
    total = sum(counts.values())
    attributed = 0
    with_guid = 0
    unresolved = 0
    non_user_or_unknown = 0
    labels = {}
    unresolved_index = 0
    # Deterministic labels within a report, never expose internal IDs.
    for kind, creator in sorted({(kind, creator) for kind, creator, _ in counts}):
        key = (kind, creator)
        email = (emails.get(creator) if kind == 'user' else None)
        email = email or metadata.get(key, {}).get('email')
        if kind == 'email':
            email = creator
        if email:
            label = email
        elif kind in ('user', 'site_user'):
            unresolved_index += 1
            name = metadata.get(key, {}).get('name')
            label = f'Unresolved user {unresolved_index}'
            if name:
                label += f' - {name}'
        elif kind == 'unresolved_user':
            label = 'Unattributed users (combined)'
        elif kind == 'non_user':
            label = 'Non-user creators (combined)'
        else:
            label = 'Unknown creators (combined)'
        labels[key] = (label, bool(email))
    for (kind, creator, facet), count in counts.items():
        label, has_email = labels[kind, creator]
        by_label[label, facet] += count
        if kind == 'user':
            with_guid += count
        if has_email:
            attributed += count
        elif kind in ('user', 'site_user', 'unresolved_user'):
            unresolved += count
        else:
            non_user_or_unknown += count
    rows = []
    for label in sorted({label for label, _ in by_label}):
        files = by_label[label, 'file']
        folders = by_label[label, 'folder']
        rows.append(dict(zip(COLUMNS, (label, files, folders, files + folders))))
    rows.sort(key=lambda row: -row['Total Items Created'])
    return {'rows': rows, 'coverage': {
        'status': 'partial' if partial else 'completed_no_recorded_failures',
        'discovered_drives': expected_drives,
        'completed_drives': completed_drives,
        'observed_items': total,
        'items_with_user_guid': with_guid,
        'items_with_user_email': attributed,
        'unresolved_user_items': unresolved,
        'non_user_or_unknown_items': non_user_or_unknown,
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
