"""
update_dashboard.py
Reads claims_analysis_output.xlsx and splices fresh data into dashboard.html.
Run standalone:  python update_dashboard.py
Run in notebook: %run update_dashboard.py
"""
import json
import re
import pathlib
import pandas as pd
import numpy as np
from datetime import date

XL   = pathlib.Path(__file__).parent / 'claims_analysis_output.xlsx'
DASH = pathlib.Path(__file__).parent / 'dashboard.html'


def _j(v):
    if isinstance(v, np.integer):  return int(v)
    if isinstance(v, np.floating): return float(v)
    return v


def _summary_rows(df):
    rows = []
    for _, r in df.sort_values('Total Open Claims', ascending=False).iterrows():
        rows.append({
            'wh':    str(r['Warehouse']),
            'lt30':  _j(r['<30 days']),
            'd30':   _j(r['30-<60 days']),
            'd60':   _j(r['60-<90 days']),
            'ge90':  _j(r['>= 90 days']),
            'total': _j(r['Total Open Claims']),
            'pct90': round(float(r['% >= 90 days']) * 100, 1) if r['Total Open Claims'] > 0 else 0,
        })
    return rows


def _oldest_rows(combined_all, dataset_name, n=10):
    open_df = combined_all[
        (combined_all['Dataset'] == dataset_name) &
        (combined_all['Progress'] != 'Completed')
    ].copy()
    open_df['Start Date'] = pd.to_datetime(open_df['Start Date'], errors='coerce')
    today = pd.Timestamp('today').normalize()
    open_df['days_open'] = (today - open_df['Start Date']).dt.days
    top = open_df.sort_values('days_open', ascending=False).head(n)
    rows = []
    for _, r in top.iterrows():
        parts = str(r['Task Name']).split(' - ')
        label = ' \u2013 '.join(parts[:3]) if len(parts) >= 3 else str(r['Task Name'])[:70]
        rows.append({
            'task':      label,
            'warehouse': str(r['Warehouse']),
            'days':      int(r['days_open']) if pd.notna(r['days_open']) else 0,
            'start':     str(r['Start Date'])[:10] if pd.notna(r['Start Date']) else '',
        })
    return rows


def _action_rows(open_tasks_agg, dataset_name):
    rows = []
    sub = open_tasks_agg[open_tasks_agg['Dataset'] == dataset_name].sort_values('Count', ascending=False)
    for _, r in sub.iterrows():
        rows.append({
            'wh':     str(r['Warehouse']),
            'action': str(r['Action Required']),
            'count':  _j(r['Count']),
        })
    return rows


def build_data_js():
    aging_summary  = pd.read_excel(XL, sheet_name='Open_Aging_Summary')
    combined_all   = pd.read_excel(XL, sheet_name='All_Claims_Combined')
    open_tasks_agg = pd.read_excel(XL, sheet_name='Open_Tasks_Aggregate')

    ford_sum   = _summary_rows(aging_summary[aging_summary['Dataset'] == 'Ford Claims'])
    chrys_sum  = _summary_rows(aging_summary[aging_summary['Dataset'] == 'Chrysler Claims'])
    ford_old   = _oldest_rows(combined_all, 'Ford Claims')
    chrys_old  = _oldest_rows(combined_all, 'Chrysler Claims')
    ford_acts  = _action_rows(open_tasks_agg, 'Ford Claims')
    chrys_acts = _action_rows(open_tasks_agg, 'Chrysler Claims')

    ford_total        = sum(r['total'] for r in ford_sum)
    chrys_total       = sum(r['total'] for r in chrys_sum)
    ford_ge90         = sum(r['ge90']  for r in ford_sum)
    chrys_ge90        = sum(r['ge90']  for r in chrys_sum)
    ford_biggest      = max(ford_sum,  key=lambda r: r['total']) if ford_sum  else {'wh': '', 'total': 0}
    chrys_biggest     = max(chrys_sum, key=lambda r: r['total']) if chrys_sum else {'wh': '', 'total': 0}
    ford_oldest_days  = max((r['days'] for r in ford_old),  default=0)
    chrys_oldest_days = max((r['days'] for r in chrys_old), default=0)
    ford_oldest_wh    = next((r['warehouse'] for r in ford_old  if r['days'] == ford_oldest_days),  '')
    chrys_oldest_wh   = next((r['warehouse'] for r in chrys_old if r['days'] == chrys_oldest_days), '')

    as_of = date.today().strftime('%B %d, %Y')

    lines = [
        'const AS_OF         = {};'.format(json.dumps(as_of)),
        'const fordSummary   = {};'.format(json.dumps(ford_sum,   indent=2)),
        'const chrysSummary  = {};'.format(json.dumps(chrys_sum,  indent=2)),
        'const fordOldest    = {};'.format(json.dumps(ford_old,   indent=2)),
        'const chrysOldest   = {};'.format(json.dumps(chrys_old,  indent=2)),
        'const fordActions   = {};'.format(json.dumps(ford_acts,  indent=2)),
        'const chrysActions  = {};'.format(json.dumps(chrys_acts, indent=2)),
        'const fordTotal     = {};'.format(ford_total),
        'const chrysTotal    = {};'.format(chrys_total),
        'const fordGe90      = {};'.format(ford_ge90),
        'const chrysGe90     = {};'.format(chrys_ge90),
        'const fordBiggestWh = {};'.format(json.dumps(ford_biggest['wh'])),
        'const fordBiggestN  = {};'.format(ford_biggest['total']),
        'const chrysBiggestWh= {};'.format(json.dumps(chrys_biggest['wh'])),
        'const chrysBiggestN = {};'.format(chrys_biggest['total']),
        'const fordOldestDays  = {};'.format(ford_oldest_days),
        'const fordOldestWh    = {};'.format(json.dumps(ford_oldest_wh)),
        'const chrysOldestDays = {};'.format(chrys_oldest_days),
        'const chrysOldestWh   = {};'.format(json.dumps(chrys_oldest_wh)),
    ]
    return '\n'.join(lines)


def update():
    data_js = build_data_js()
    replacement = '// @@DATA_START@@\n' + data_js + '\n// @@DATA_END@@'
    html = DASH.read_text(encoding='utf-8')
    html = re.sub(
        r'// @@DATA_START@@.*?// @@DATA_END@@',
        lambda _: replacement,
        html,
        flags=re.DOTALL,
    )
    DASH.write_text(html, encoding='utf-8')
    print('Dashboard updated:', DASH)


if __name__ == '__main__':
    update()
