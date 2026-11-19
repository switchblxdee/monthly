"""Keep source provenance and counts with the report throughout the graph."""
from collections import Counter, defaultdict


class SourceInventory:
    def __call__(self, state):
        sheets = defaultdict(lambda: {'rows': 0, 'sources': Counter(), 'missing_source_rows': 0})
        for record in state.get('signals', []):
            stats = sheets[record['sheet']]
            stats['rows'] += 1
            name = str(record.get('source') or '').strip()
            if name:
                stats['sources'][name] += 1
            else:
                stats['missing_source_rows'] += 1
        detail = [{'sheet': sheet, 'rows_in_period': s['rows'],
                   'source_count': len(s['sources']), 'sources': dict(sorted(s['sources'].items())),
                   'missing_source_rows': s['missing_source_rows']}
                  for sheet, s in sorted(sheets.items())]
        represented = {x['sheet'] for x in detail}
        detail.extend({'sheet': name, 'rows_in_period': 0, 'source_count': 0,
                       'sources': {}, 'missing_source_rows': 0}
                      for name in state.get('detail_sheets') or [] if name not in represented)
        date_filter = state.get('date_filter') or {}
        other = Counter(ref['sheet'] for ref in date_filter.get('other_month_references', []))
        unknown = Counter(ref['sheet'] for ref in date_filter.get('unknown_date_references', []))
        skipped = Counter(ref['sheet'] for ref in state.get('skipped_rows', []) if ref.get('kind') == 'detail')
        for sheet in detail:
            name = sheet['sheet']
            sheet['rows_excluded_other_month'] = other[name]
            sheet['rows_excluded_unknown_date'] = unknown[name]
            sheet['skipped_invalid_rows'] = skipped[name]
            sheet['parsed_rows_all_periods'] = sheet['rows_in_period'] + other[name] + unknown[name]
            sheet['nonempty_rows_seen'] = sheet['parsed_rows_all_periods'] + skipped[name]
        return {'source_inventory': {'priority_sheet': state.get('current_sheet'),
                'priority_rows': len(state.get('topics', [])),
                'previous_sheet': state.get('previous_sheet'),
                'previous_rows': len(state.get('previous_topics', [])),
                'detail_sheets': detail, 'detail_rows_in_period': len(state.get('signals', [])),
                'skipped_rows': len(state.get('skipped_rows', [])),
                'date_filter': state.get('date_filter')}}
