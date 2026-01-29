#!/usr/bin/env python3
import csv
import pathlib
import uuid
from datetime import datetime, timezone

BASE_DIR = pathlib.Path(__file__).resolve().parents[1]
INPUT_PATH = BASE_DIR / 'OKRs' / 'OKRs_DO_NOT_EDIT_Cleaned.csv'
OUTPUT_OBJECTIVES = BASE_DIR / 'OKRs' / 'Import_Objectives.tsv'
OUTPUT_KEY_RESULTS = BASE_DIR / 'OKRs' / 'Import_KeyResults.tsv'

OBJECTIVE_HEADERS = [
    'id', 'title', 'level', 'parentId', 'ownerEmail', 'ownerName', 'department', 'team',
    'quarter', 'status', 'priority', 'impact', 'progress', 'description', 'rationale',
    'createdAt', 'updatedAt'
]

KEY_RESULT_HEADERS = [
    'id', 'objectiveId', 'title', 'metric', 'baseline', 'target', 'current',
    'timeline', 'status', 'confidence', 'ownerEmail', 'updatedAt', 'createdAt'
]


def normalize_progress(value):
    if value is None or value == '':
        return ''
    try:
        num = float(value)
    except ValueError:
        return value
    if num <= 1:
        num *= 100
    return round(num, 2)


def split_owner(owner):
    if not owner:
        return ('', '')
    if ':' in owner:
        parts = [p.strip() for p in owner.split(':', 1)]
        return (parts[0], parts[1])
    return (owner.strip(), '')


def build_status(baseline_type, comments):
    baseline_type = (baseline_type or '').strip()
    comments = (comments or '').strip()
    if baseline_type and comments:
        return f'{baseline_type}; {comments}'
    return baseline_type or comments


now = datetime.now(timezone.utc).isoformat()

objectives = {}
key_results = []

with INPUT_PATH.open(newline='') as f:
    reader = csv.DictReader(f)
    for row in reader:
        objective_title = (row.get('Objective') or '').strip()
        deadline = (row.get('Deadline') or '').strip()
        owner = (row.get('Owner') or '').strip()
        progress = normalize_progress((row.get('Progress') or '').strip())
        department, team = split_owner(owner)

        obj_key = (objective_title, deadline)
        if obj_key not in objectives:
            obj_id = str(uuid.uuid4())
            objectives[obj_key] = {
                'id': obj_id,
                'title': objective_title,
                'level': 'Executive',
                'parentId': '',
                'ownerEmail': '',
                'ownerName': owner,
                'department': department,
                'team': team,
                'quarter': deadline,
                'status': '',
                'priority': '',
                'impact': '',
                'progress': progress,
                'description': objective_title,
                'rationale': 'Imported from OKRs_DO_NOT_EDIT_Cleaned.csv',
                'createdAt': now,
                'updatedAt': now
            }
        else:
            obj_id = objectives[obj_key]['id']

        kr_title = (row.get('Key Results') or '').strip()
        kr_metric = (row.get('Metric') or '').strip()
        kr_baseline = 0
        kr_target = (row.get('Target') or '').strip()
        kr_current = (row.get('Current') or '').strip()
        baseline_type = (row.get('Baseline/Stretch') or '').strip()
        comments = (row.get('Status / Comments') or '').strip()
        kr_status = build_status(baseline_type, comments)

        key_results.append({
            'id': str(uuid.uuid4()),
            'objectiveId': obj_id,
            'title': kr_title,
            'metric': kr_metric,
            'baseline': kr_baseline,
            'target': kr_target,
            'current': kr_current,
            'timeline': deadline,
            'status': kr_status,
            'confidence': '',
            'ownerEmail': '',
            'updatedAt': now,
            'createdAt': now
        })

OUTPUT_OBJECTIVES.parent.mkdir(parents=True, exist_ok=True)

with OUTPUT_OBJECTIVES.open('w', newline='') as f:
    writer = csv.writer(f, delimiter='\t')
    writer.writerow(OBJECTIVE_HEADERS)
    for record in objectives.values():
        writer.writerow([record.get(h, '') for h in OBJECTIVE_HEADERS])

with OUTPUT_KEY_RESULTS.open('w', newline='') as f:
    writer = csv.writer(f, delimiter='\t')
    writer.writerow(KEY_RESULT_HEADERS)
    for record in key_results:
        writer.writerow([record.get(h, '') for h in KEY_RESULT_HEADERS])

print(f'Wrote {OUTPUT_OBJECTIVES}')
print(f'Wrote {OUTPUT_KEY_RESULTS}')
