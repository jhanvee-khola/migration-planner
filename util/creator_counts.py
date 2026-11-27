"""Creator attribution for live file and folder observations during deep scan."""
import csv
import json
from pathlib import Path
from uuid import UUID
from functools import lru_cache

COLUMNS = ('Creator Type', 'Creator ID', 'Files Created', 'Folders Created', 'Total Items Created')


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


def creator_report(counts, expected_drives, completed_drives, partial):
    keys = {(kind, creator) for kind, creator, _ in counts}
    rows = []
    for kind, creator in sorted(keys):
        files = counts[(kind, creator, 'file')]
        folders = counts[(kind, creator, 'folder')]
        rows.append(dict(zip(COLUMNS, (kind, creator, files, folders, files + folders))))
    rows.sort(key=lambda row: -row['Total Items Created'])
    total = sum(row['Total Items Created'] for row in rows)
    attributed = sum(row['Total Items Created'] for row in rows if row['Creator Type'] == 'user')
    return {'rows': rows, 'coverage': {
        'status': 'partial' if partial else 'completed_no_recorded_failures',
        'discovered_drives': expected_drives,
        'completed_drives': completed_drives,
        'observed_items': total,
        'items_with_user_guid': attributed,
        'tenant_membership_verified': False,
        'count_semantics': 'first observed live non-root files/folders; no historical versions or deleted items; existing facet precedence',
    }}


def write_creator_report(output_dir, report):
    directory = Path(output_dir)
    path = directory / 'creator_counts.csv'
    with path.open('w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(report['rows'])
    with (directory / 'creator_counts_coverage.json').open('w', encoding='utf-8') as f:
        json.dump(report['coverage'], f, indent=2)
    return path


def is_terminal_page(response):
    status = response.get('status')
    body = response.get('body')
    return (isinstance(status, int) and 200 <= status < 300
            and isinstance(body, dict) and isinstance(body.get('value'), list)
            and '@odata.nextLink' not in body)
