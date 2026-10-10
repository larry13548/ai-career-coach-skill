#!/usr/bin/env python3
"""Deterministic calculation for evidence judgments; no network or inference."""
import argparse
import hashlib
import json
import sys
from decimal import Decimal, ROUND_HALF_UP

POINTS = {'matched': 2, 'partial': 1, 'unknown': 0, 'unmatched': 0}


def validate(data):
    if not isinstance(data, dict) or not isinstance(data.get('job'), str) or not data['job'].strip():
        raise ValueError('job must identify the target position')
    rows = data.get('requirements')
    if not isinstance(rows, list):
        raise ValueError('requirements must be a list')
    ids = set()
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError('each requirement must be an object')
        for key in ('id', 'requirement', 'jd_quote', 'category'):
            if not isinstance(row.get(key), str) or not row[key].strip():
                raise ValueError(f'{key} must be a nonempty string')
        if row['id'] in ids:
            raise ValueError('duplicate requirement id')
        ids.add(row['id'])
        if row.get('kind') not in ('required', 'bonus') or row.get('status') not in POINTS:
            raise ValueError('invalid kind or status')
        if type(row.get('hard_gate')) is not bool:
            raise ValueError('hard_gate must be boolean and grounded in JD')
        if row['kind'] == 'bonus' and row['hard_gate']:
            raise ValueError('a bonus cannot be a hard gate')
        for key in ('evidence', 'source'):
            if not isinstance(row.get(key), str):
                raise ValueError(f'{key} must be a string')
            if row['status'] != 'unknown' and not row[key].strip():
                raise ValueError('matched/partial/unmatched judgments need evidence and source')
    return rows


def manifest(data):
    validate(data)
    keys = ('id', 'requirement', 'jd_quote', 'category', 'kind', 'hard_gate')
    return {'job': data['job'], 'requirements': sorted(
        [{k: row[k] for k in keys} for row in data['requirements']], key=lambda r: r['id'])}


def score(data, baseline=None):
    rows = validate(data)
    m = manifest(data)
    if baseline is not None and m != manifest(baseline):
        raise ValueError('requirements changed: explain the JD correction and start an explicitly revised baseline')
    required = [r for r in rows if r['kind'] == 'required']
    bonus = [r for r in rows if r['kind'] == 'bonus']
    units = sum(POINTS[r['status']] for r in required)
    n = len(required)
    unknown = [r['id'] for r in required if r['status'] == 'unknown']
    hard_gaps = [r['id'] for r in required if r['hard_gate'] and r['status'] in ('partial', 'unmatched')]
    hard_unknown = [r['id'] for r in required if r['hard_gate'] and r['status'] == 'unknown']
    # Compare integer half-point units, never the rounded display value.
    branch = 'no_requirements' if not n else ('above_90' if units * 5 > n * 9 else 'at_90' if units * 5 == n * 9 else 'below_90')
    messages = {
        'no_requirements': '没有可识别的常规或必备要求，请补充JD，不生成分数。',
        'above_90': '材料基本匹配，建议投递尝试，并说明剩余缺口。',
        'at_90': '达到90%临界值，可投递尝试，同时说明剩余缺口。',
        'below_90': '精准说明差距与补充方向，不仅报低分，也不直接劝退。',
    }
    advice = messages[branch]
    if hard_gaps:
        advice = '存在明确硬性条件未完全满足，应先提示并确认企业是否接受放宽条件，不能仅凭总分宣布基本匹配。'
    elif hard_unknown:
        advice = '有硬性条件尚待确认，应优先澄清；当前匹配度为暂评。'
    if n and unknown:
        advice += ' 信息缺失暂计0分，不等于候选人不具备能力。'
    percent = ((Decimal(units) * 50 / Decimal(n)).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)) if n else None
    return {
        'job': data['job'], 'score': units / 2, 'total': n,
        'percentage': str(percent) if percent is not None else None,
        'threshold_branch': branch, 'provisional': bool(unknown),
        'unknown_ids': unknown,
        'unmatched_ids': [r['id'] for r in required if r['status'] == 'unmatched'],
        'hard_gap_ids': hard_gaps, 'hard_unknown_ids': hard_unknown,
        'all_required_matched': bool(n) and all(r['status'] == 'matched' for r in required),
        'bonus': bonus, 'requirements': required, 'advice': advice,
        'manifest_sha256': hashlib.sha256(json.dumps(m, ensure_ascii=False, sort_keys=True).encode()).hexdigest(),
    }


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('input'); p.add_argument('--baseline')
    args = p.parse_args()
    try:
        with open(args.input, encoding='utf-8') as f: data = json.load(f)
        previous = None
        if args.baseline:
            with open(args.baseline, encoding='utf-8') as f: previous = json.load(f)
        print(json.dumps(score(data, previous), ensure_ascii=False, indent=2))
    except (OSError, ValueError, TypeError) as exc:
        print(f'Cannot calculate: {exc}', file=sys.stderr); return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
