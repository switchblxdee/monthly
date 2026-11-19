import json
import pytest
from golos_mas.promises import PromisesAgent

def record(id,owner,due,outcome,**extra):
    return {'id':id,'owner':owner,'description':'Задача '+id,'source':'Встреча','initial_due':due,'outcome':outcome,'in_work_start':True,'in_work_current':False,**extra}

def test_counts_and_denominators(tmp_path):
    data=[record('1','A','2026-10-31','postponed',current_due='2026-11-30',in_work_current=True),record('2','A','2026-10-20','completed_on_time'),record('3','B','2026-11-10','in_progress',in_work_current=True)]
    path=tmp_path/'p.json';path.write_text(json.dumps(data),encoding='utf8')
    result=PromisesAgent()({'promises_json':str(path),'report_month':'2026-10','warnings':[]})['promise_report']
    assert result['due_count']==2 and result['completed_count']==1 and result['current_total']==2
    assert result['due_in_work']==1 and result['counts']['postponed']==1
    assert result['groups'][0]=={'owner':'A','start':2,'due':2,'completed':1,'postponed':1}
    assert result['groups'][1]['due']==0

def test_postponed_requires_later_date(tmp_path):
    path=tmp_path/'p.json';path.write_text(json.dumps([record('1','A','2026-10-31','postponed',current_due='2026-10-01')]),encoding='utf8')
    with pytest.raises(ValueError,match='позднее исходного'):
        PromisesAgent()({'promises_json':str(path),'report_month':'2026-10','warnings':[]})

def test_empty_due_set_has_no_fake_percentage(tmp_path):
    from golos_mas.template_render import percent
    path=tmp_path/'p.json';path.write_text(json.dumps([record('1','A','2026-11-30','in_progress')]),encoding='utf8')
    report=PromisesAgent()({'promises_json':str(path),'report_month':'2026-10','warnings':[]})['promise_report']
    assert report['due_count']==0 and percent(report['completed_count'],report['due_count'])=='—'

def test_percentage_rounds_half_up():
    from golos_mas.template_render import percent
    assert percent(137,200)=='69%' and percent(2,3)=='67%'
