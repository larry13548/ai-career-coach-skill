import copy
import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from score_match import score


def sample(states):
    return {'job':'测试岗位','requirements':[{'id':str(i),'requirement':f'岗位要求{i}','jd_quote':f'JD要求{i}',
        'category':'职责','dimension':'experience','kind':'required','hard_gate':False,'status':s,
        'evidence':'' if s=='unknown' else '测试确认的证据','source':'' if s=='unknown' else '合成案例'} for i,s in enumerate(states)]}


class ScoreTests(unittest.TestCase):
    def test_weighted_tenure_only_costs_ten_points(self):
        d=sample(['matched','matched','partial'])
        d['requirements'][1]['dimension']='capability'
        d['requirements'][2].update(dimension='tenure',required_months=36,actual_months=12)
        r=score(d)
        self.assertEqual(r['percentage'],'93.3')
        self.assertAlmostEqual(r['dimensions']['tenure']['score'],10/3)
        d['requirements'][2].update(status='unmatched',actual_months=0)
        self.assertEqual(score(d)['percentage'],'90.0')

    def test_education_gap_does_not_change_tenure(self):
        d=sample(['matched','unmatched','matched','matched','matched','matched'])
        for row in d['requirements'][1:5]: row['dimension']='capability'
        d['requirements'][1]['hard_gate']=True
        d['requirements'][5].update(dimension='tenure',required_months=60,actual_months=60)
        r=score(d)
        self.assertEqual(r['percentage'],'92.5')
        self.assertEqual(r['dimensions']['tenure']['score'],10)
        self.assertEqual(r['hard_gap_ids'],['1'])

    def test_transferable_and_missing_dimension(self):
        r=score(sample(['transferable']))
        self.assertEqual(r['percentage'],'75.0')
        self.assertTrue(r['weight_redistributed'])
        self.assertEqual(r['dimensions']['tenure']['weight'],0)

    def test_tenure_validation_and_baseline(self):
        d=sample(['partial']);row=d['requirements'][0]
        row.update(dimension='tenure',required_months=60,actual_months=24)
        self.assertEqual(score(d)['percentage'],'40.0')
        newer=copy.deepcopy(d);newer['requirements'][0]['actual_months']=36
        self.assertEqual(score(newer,d)['percentage'],'60.0')
        newer['requirements'][0]['required_months']=48
        with self.assertRaises(ValueError):score(newer,d)
        for bad in [0,-1,True]:
            row['required_months']=bad
            with self.assertRaises(ValueError):score(d)
        row.update(required_months=60,actual_months=70)
        with self.assertRaises(ValueError):score(d)

    def test_unknown_tenure_is_provisional(self):
        d=sample(['unknown']);d['requirements'][0].update(dimension='tenure',required_months=60)
        r=score(d);self.assertTrue(r['provisional']);self.assertEqual(r['percentage'],'0.0')

    def test_thresholds(self):
        for states, branch, percent in [(['matched']*8+['partial']*2,'at_90','90.0'),
                (['matched']*9+['partial'],'above_90','95.0'),
                (['matched']*8+['unmatched']*2,'below_90','80.0')]:
            r=score(sample(states));self.assertEqual(r['threshold_branch'],branch);self.assertEqual(r['percentage'],percent)

    def test_rounding_does_not_change_branch(self):
        r=score(sample(['matched']*1800+['unmatched']*201))
        self.assertEqual(r['percentage'],'90.0');self.assertEqual(r['threshold_branch'],'below_90')

    def test_unknown_is_provisional_not_unmatched(self):
        r=score(sample(['matched','unknown']))
        self.assertTrue(r['provisional']);self.assertEqual(r['unknown_ids'],['1']);self.assertFalse(r['unmatched_ids'])

    def test_bonus_excluded(self):
        d=sample(['matched','unmatched']);d['requirements'][1]['kind']='bonus'
        r=score(d);self.assertEqual(r['requirement_count'],1);self.assertEqual(r['percentage'],'100.0')

    def test_hard_gate_overrides_high_score(self):
        d=sample(['matched']*20+['unmatched']);d['requirements'][-1]['hard_gate']=True
        r=score(d);self.assertEqual(r['threshold_branch'],'above_90');self.assertTrue(r['hard_gap_ids']);self.assertIn('先提示',r['advice'])

    def test_empty_and_only_bonus_have_no_score(self):
        for d in [sample([]),sample(['matched'])]:
            if d['requirements']:d['requirements'][0]['kind']='bonus'
            self.assertIsNone(score(d)['percentage'])

    def test_fixed_denominator_allows_evidence_update(self):
        d=sample(['unknown']);new=copy.deepcopy(d)
        new['requirements'][0].update(status='matched',evidence='新确认经历',source='用户补充')
        self.assertEqual(score(new,d)['percentage'],'100.0')
        new['requirements'][0]['requirement']='偷偷换要求'
        with self.assertRaises(ValueError):score(new,d)

    def test_rejects_duplicate_and_missing_evidence(self):
        d=sample(['matched']);d['requirements'].append(copy.deepcopy(d['requirements'][0]))
        with self.assertRaises(ValueError):score(d)
        d=sample(['matched']);d['requirements'][0]['source']=''
        with self.assertRaises(ValueError):score(d)


if __name__=='__main__':unittest.main()
