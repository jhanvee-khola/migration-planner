"""Offline creator attribution tests, including the actual delta traversal."""
import csv
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

from estimators.file_estimator import FileEstimator
from util.creator_counts import creator_key, creator_report, creator_counts_enabled, write_creator_report, is_terminal_page
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
        self.assertEqual(creator_key({'createdBy': {'application': {'id': B}}}), ('application', B))
        self.assertEqual(creator_key({'createdBy': {'siteUser': {'id': '42'}}}), ('unresolved_user', ''))
        for bad in ({}, {'createdBy': None}, {'createdBy': []}, {'createdBy': {'user': None}}):
            self.assertEqual(creator_key(bad), ('unknown', ''))

    def test_merge_and_coverage(self):
        first = Counter({('user', A, 'file'): 3, ('unknown', '', 'folder'): 1})
        first.update(Counter({('user', A, 'folder'): 2, ('user', B, 'file'): 1}))
        report = creator_report(first, 2, 1, True)
        self.assertEqual(report['rows'][0]['Total Items Created'], 5)
        self.assertEqual(report['coverage']['observed_items'], 7)
        self.assertEqual(report['coverage']['items_with_user_guid'], 6)
        self.assertEqual(report['coverage']['status'], 'partial')

    def test_terminal_page(self):
        self.assertTrue(is_terminal_page({'status': 200, 'body': {'value': []}}))
        for response in ({}, {'status': 200, 'body': None},
                         {'status': 403, 'body': {'value': []}},
                         {'status': 200, 'body': {'value': [], '@odata.nextLink': 'next'}}):
            self.assertFalse(is_terminal_page(response))

    def test_empty_report(self):
        report = creator_report(Counter(), 0, 0, False)
        with tempfile.TemporaryDirectory() as directory:
            path = write_creator_report(directory, report)
            with path.open() as f:
                self.assertEqual(list(csv.DictReader(f)), [])
            self.assertEqual({p.name for p in Path(directory).iterdir()}, {'creator_counts.csv'})

    def test_nonempty_report_csv_only(self):
        report = creator_report(Counter({('user', A, 'file'): 2,
                                         ('user', A, 'folder'): 1}), 2, 1, True)
        with tempfile.TemporaryDirectory() as directory:
            path = write_creator_report(directory, report)
            with path.open() as f:
                rows = list(csv.DictReader(f))
            self.assertEqual(rows, [{'Creator Type': 'user', 'Creator ID': A,
                                     'Files Created': '2', 'Folders Created': '1',
                                     'Total Items Created': '3'}])
            self.assertEqual({p.name for p in Path(directory).iterdir()}, {'creator_counts.csv'})


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
