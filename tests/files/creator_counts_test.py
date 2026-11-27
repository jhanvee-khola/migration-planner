"""Offline creator attribution tests, including the actual delta traversal."""
import csv
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

from estimators.file_estimator import FileEstimator
from util.creator_counts import creator_key, creator_report, creator_counts_enabled, write_creator_report, is_terminal_page, resolve_creator_emails, USERS_URL, DELETED_USERS_URL
from util.utils import ScanConfig
from util.thread_safe_ds import ThreadSafeMap

A = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
B = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'


class CreatorCountsTest(unittest.TestCase):
    def test_mode(self):
        config = SimpleNamespace(user_source='tenant', shallow_scan=False)
        self.assertTrue(creator_counts_enabled(config, {}))
        for key in ('emailIds', 'siteUrls', 'drives'):
            self.assertFalse(creator_counts_enabled(config, {key: ['x']}))
        self.assertFalse(creator_counts_enabled(SimpleNamespace(user_source='csv'), {}))
        self.assertFalse(creator_counts_enabled(SimpleNamespace(user_source='tenant', shallow_scan=True), {}))

    def test_identity(self):
        def item(user, name):
            return {'createdBy': {'user': {'id': user, 'displayName': name}}}
        self.assertEqual(creator_key(item(A.upper(), 'renamed')), ('user', A))
        self.assertNotEqual(creator_key(item(A, 'same')), creator_key(item(B, 'same')))
        self.assertEqual(creator_key(item('42', 'local')), ('unresolved_user', ''))
        self.assertEqual(creator_key({'createdBy': {'user': {'id': A}, 'application': {'id': B}}}), ('user', A))
        self.assertEqual(creator_key({'createdBy': {'application': {'id': B}}}), ('non_user', ''))
        self.assertEqual(creator_key({'createdBy': {'siteUser': {'id': '42'}}}), ('unresolved_user', ''))
        for bad in ({}, {'createdBy': None}, {'createdBy': []}, {'createdBy': {'user': None}}):
            self.assertEqual(creator_key(bad), ('unknown', ''))

    def test_merge_and_coverage(self):
        first = Counter({('user', A, 'file'): 3, ('unknown', '', 'folder'): 1})
        first.update(Counter({('user', A, 'folder'): 2, ('user', B, 'file'): 1}))
        report = creator_report(first, 2, 1, True, {A: 'a@example.com', B: 'b@example.com'})
        self.assertEqual(report['rows'][0]['Total Items Created'], 5)
        self.assertEqual(report['coverage']['observed_items'], 7)
        self.assertEqual(report['coverage']['items_with_user_guid'], 6)
        self.assertEqual(report['coverage']['status'], 'partial')
        self.assertEqual(report['coverage']['items_with_user_email'], 6)
        self.assertEqual(report['coverage']['non_user_or_unknown_items'], 1)
        self.assertEqual(sum(row['Total Items Created'] for row in report['rows']), 7)

    def test_terminal_page(self):
        self.assertTrue(is_terminal_page({'status': 200, 'body': {'value': []}}))
        for response in ({}, {'status': 200, 'body': None},
                         {'status': 403, 'body': {'value': []}},
                         {'status': 200, 'body': {'value': [], '@odata.nextLink': 'next'}}):
            self.assertFalse(is_terminal_page(response))

    def test_empty_report(self):
        report = creator_report(Counter(), 0, 0, False)
        with tempfile.TemporaryDirectory() as directory:
            path = write_creator_report(directory, report, '20261127_123456')
            with path.open() as f:
                self.assertEqual(list(csv.DictReader(f)), [])
            self.assertEqual({p.name for p in Path(directory).iterdir()}, {'creator_counts_20261127_123456.csv'})

    def test_nonempty_report_csv_only(self):
        report = creator_report(Counter({('user', A, 'file'): 2,
                                         ('user', A, 'folder'): 1}), 2, 1, True, {A: 'a@example.com'})
        with tempfile.TemporaryDirectory() as directory:
            path = write_creator_report(directory, report, '20261127_123456')
            with path.open() as f:
                rows = list(csv.DictReader(f))
            self.assertEqual(rows, [{'User Email': 'a@example.com',
                                     'Files Created': '2', 'Folders Created': '1',
                                     'Total Items Created': '3'}])
            self.assertEqual({p.name for p in Path(directory).iterdir()}, {'creator_counts_20261127_123456.csv'})

    def test_resolve_emails_paginated_and_stop_early(self):
        counts = Counter({('user', A, 'file'): 10, ('user', B, 'folder'): 2})
        next_url = 'https://graph.microsoft.com/v1.0/users?$skiptoken=1'
        fetch = MagicMock(side_effect=[
            {'value': [{'id': A, 'mail': 'A@Example.com', 'userPrincipalName': 'other@example.com'}],
             '@odata.nextLink': next_url},
            {'value': [{'id': B, 'mail': None, 'userPrincipalName': 'B@example.com'}],
             '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=2'},
        ])
        self.assertEqual(resolve_creator_emails(counts, fetch, lambda: False, MagicMock()),
                         {A: 'a@example.com', B: 'b@example.com'})
        self.assertEqual([call.args[0] for call in fetch.call_args_list], [USERS_URL, next_url])

    def test_unresolved_and_non_user_items_retained_without_raw_ids(self):
        counts = Counter({('user', A, 'file'): 10, ('user', B, 'folder'): 2,
                          ('non_user', '', 'file'): 3, ('unknown', '', 'file'): 4})
        fetch = MagicMock(return_value={'value': [
            {'id': A, 'mail': 'a@example.com'},
            {'id': B, 'mail': None, 'userPrincipalName': 'guest#EXT#@tenant.example.com'},
        ]})
        emails = resolve_creator_emails(counts, fetch, lambda: False, MagicMock())
        report = creator_report(counts, 1, 1, False, emails, {('user', B): {'name': 'Alex Smith'}})
        rows = {row['User Email']: row['Total Items Created'] for row in report['rows']}
        self.assertEqual(rows, {'a@example.com': 10, 'Unresolved user 1 - Alex Smith': 2,
                                'Non-user creators (combined)': 3, 'Unknown creators (combined)': 4})
        self.assertEqual(sum(rows.values()), 19)
        self.assertEqual(report['coverage']['unresolved_user_items'], 2)
        self.assertEqual(report['coverage']['non_user_or_unknown_items'], 7)

    def test_item_metadata_avoids_directory_requests(self):
        metadata = {}
        item = {'createdBy': {'user': {'id': A}, 'siteUser': {
            'loginName': 'i:0#.f|membership|Former@Example.com', 'displayName': 'Former User'}}}
        key = creator_key(item, metadata)
        counts = Counter({(*key, 'file'): 10})
        fetch = MagicMock()
        emails = resolve_creator_emails(counts, fetch, lambda: False, MagicMock(), metadata)
        self.assertEqual(emails, {A: 'former@example.com'})
        fetch.assert_not_called()
        self.assertEqual(creator_report(counts, 1, 1, False, emails, metadata)['rows'][0]['User Email'],
                         'former@example.com')

    def test_deleted_directory_fallback_only_for_remaining_ids(self):
        counts = Counter({('user', A, 'file'): 1, ('user', B, 'folder'): 2})
        fetch = MagicMock(side_effect=[
            {'value': [{'id': A, 'mail': 'a@example.com'}]},
            {'value': [{'id': B, 'mail': 'former@example.com',
                        'userPrincipalName': 'rewritten@example.com'}]},
        ])
        self.assertEqual(resolve_creator_emails(counts, fetch, lambda: False, MagicMock()),
                         {A: 'a@example.com', B: 'former@example.com'})
        self.assertEqual([call.args[0] for call in fetch.call_args_list], [USERS_URL, DELETED_USERS_URL])

    def test_deleted_upn_without_mail_not_assumed_to_be_original(self):
        counts = Counter({('user', A, 'file'): 1})
        fetch = MagicMock(side_effect=[{'value': []},
            {'value': [{'id': A, 'userPrincipalName': 'rewritten@example.com'}]}])
        self.assertEqual(resolve_creator_emails(counts, fetch, lambda: False, MagicMock()), {})

    def test_site_local_ids_and_duplicate_names_do_not_merge(self):
        metadata = {}
        def item(site):
            return {'parentReference': {'siteId': site}, 'createdBy': {
                'siteUser': {'id': '42', 'displayName': 'Alex Smith'}}}
        key1 = creator_key(item('site1'), metadata)
        key2 = creator_key(item('site2'), metadata)
        self.assertNotEqual(key1, key2)
        counts = Counter({(*key1, 'file'): 2, (*key2, 'folder'): 3})
        rows = creator_report(counts, 2, 2, False, metadata=metadata)['rows']
        self.assertEqual({row['User Email'] for row in rows},
                         {'Unresolved user 1 - Alex Smith', 'Unresolved user 2 - Alex Smith'})
        self.assertEqual(sum(row['Total Items Created'] for row in rows), 5)

    def test_missing_identity_preserved_in_combined_bucket(self):
        counts = Counter({('unresolved_user', '', 'file'): 600000})
        report = creator_report(counts, 1, 1, False)
        self.assertEqual(report['rows'][0]['User Email'], 'Unattributed users (combined)')
        self.assertEqual(report['rows'][0]['Total Items Created'], 600000)

    def test_directory_failure_does_not_raise_and_preserves_partial_results(self):
        counts = Counter({('user', A, 'file'): 1, ('user', B, 'folder'): 2})
        fetch = MagicMock(side_effect=[
            {'value': [{'id': A, 'mail': 'a@example.com'}],
             '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=1'},
            RuntimeError('Forbidden'),
        ])
        logger = MagicMock()
        self.assertEqual(resolve_creator_emails(counts, fetch, lambda: False, logger),
                         {A: 'a@example.com'})
        logger.assert_called_once()

    def test_no_directory_request_without_users_or_after_stop(self):
        fetch = MagicMock()
        self.assertEqual(resolve_creator_emails(Counter({('unknown', '', 'file'): 1}),
                                                fetch, lambda: False, MagicMock()), {})
        self.assertEqual(resolve_creator_emails(Counter({('user', A, 'file'): 1}),
                                                fetch, lambda: True, MagicMock()), {})
        fetch.assert_not_called()


class DeltaIntegrationTest(unittest.TestCase):
    def test_projection_pagination_and_deduplication(self):
        def item(item_id, facet, parent='root', user=A):
            return {'id': item_id, facet: {}, 'size': 1,
                    'parentReference': {'driveId': 'drive', 'id': parent},
                    'createdBy': {'user': {'id': user}}}

        root = {'id': 'root', 'folder': {}, 'parentReference': {'driveId': 'drive'}}
        first = item('file1', 'file')
        folder = item('folder1', 'folder')
        second = item('file2', 'file', user=B)
        shortcut = item('shortcut1', 'remoteItem')
        first_page = {'value': [root, first, folder],
                      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next'}
        second_page = {'value': [first, second, shortcut], '@odata.deltaLink': 'done'}

        def run(enabled):
            urls = []
            invoker = MagicMock()
            def invoke(base, batch, *args):
                urls.extend(req['url'] for req in batch)
                body = first_page if len(urls) == 1 else second_page
                return [{'id': req['id'], 'status': 200, 'body': body} for req in batch]
            invoker.invoke.side_effect = invoke
            config = ScanConfig(tenant_id='t', client_ids=['id'], client_secrets=['s'],
                                user_source='tenant', csv_path='', concurrency=2)
            estimator = FileEstimator(config, invoker, logger=lambda *args: None,
                                      progress_update_callback=lambda *args, **kwargs: None)
            progress = ThreadSafeMap()
            for key in ('folderCount', 'fileCount', 'shortcutCount', 'versionCount'):
                progress.update(key, 0)
            counts = {'drive': Counter()} if enabled else None
            completed = set() if enabled else None
            failures = []
            try:
                result = estimator._create_in_memory_tree(['drive'], progress, failures,
                    creator_counts_by_drive=counts, creator_completed_drives=completed)
            finally:
                estimator.shutdown()
            self.assertEqual(failures, [])
            self.assertEqual(len(urls), 2)
            return urls, progress.get_all(), counts, completed, result[3:]

        off, baseline, _, _, base_metrics = run(False)
        on, new_metrics, counts, completed, new_tree_metrics = run(True)
        self.assertNotIn('createdBy', off[0])
        self.assertIn('createdBy', on[0])
        self.assertEqual(off[1], on[1])
        self.assertEqual(baseline, new_metrics)
        self.assertEqual(base_metrics, new_tree_metrics)
        self.assertEqual(completed, {'drive'})
        self.assertEqual(counts['drive'], Counter({('user', A, 'file'): 1,
                                                  ('user', A, 'folder'): 1,
                                                  ('user', B, 'file'): 1}))


if __name__ == '__main__':
    unittest.main()
