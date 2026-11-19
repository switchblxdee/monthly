import pytest
from golos_mas.methodology import evaluate_signal_score, PrioritizationAgent

@pytest.mark.parametrize('mau,counts',[
    (999,[(10,1),(19,1),(20,2),(29,2),(30,3),(39,3),(40,4)]),
    (1001,[(10,1),(29,1),(30,2),(39,2),(40,3),(49,3),(50,4)]),
    (20001,[(10,1),(49,1),(50,2),(59,2),(60,3),(69,3),(70,4)]),
])
def test_signal_thresholds_match_slide_four(mau,counts):
    for count,expected in counts: assert evaluate_signal_score(mau,count)['expected_signal_score']==expected

@pytest.mark.parametrize('mau',[1000,20000])
def test_unconfirmed_mau_boundaries_stay_ambiguous(mau):
    result=evaluate_signal_score(mau,50,'ambiguous')
    assert result['reason']=='ambiguous_boundary' and result['expected_signal_score'] is None

def test_boundary_options_and_missing_mau_default():
    assert evaluate_signal_score(1000,25,'upper')['audience_category']=='средний'
    assert evaluate_signal_score(1000,25,'lower')['audience_category']=='малый'
    assert evaluate_signal_score(20000,55,'upper')['audience_category']=='большой'
    assert evaluate_signal_score(20000,55,'lower')['audience_category']=='средний'
    assert evaluate_signal_score(None,30)['audience_category']=='средний'
    assert evaluate_signal_score('н.д.',30)['audience_basis']=='default_missing'
    assert evaluate_signal_score(1500,9)['reason']=='below_threshold'


def test_methodology_detects_mismatch_without_rewriting_excel_values():
    topic={'sheet':'P','row':2,'mau':25000,'count':70,'signals_score':2,'score':9}
    result=PrioritizationAgent()({'topics':[topic],'warnings':[]})
    assert result['topics'][0]['signals_score']==2
    assert result['topics'][0]['methodology_check']['expected_signal_score']==4
    assert any('балл за сигналы в Excel 2, по методике 4' in warning for warning in result['warnings'])


def test_confirmed_boundaries_are_default_and_recorded():
    assert evaluate_signal_score(1000,25)['audience_category']=='средний'
    result=evaluate_signal_score(20000,55)
    assert result['audience_category']=='большой' and result['boundary_rule']=='upper'

def test_missing_signal_score_and_total_are_derived_with_provenance(tmp_path):
    from openpyxl import Workbook
    from golos_mas.agents import ExcelAgent, AnalyticsAgent
    w=Workbook();s=w.active;s.title='P'
    s.append(['продукт','тема','MAU','кол-во сигналов','сигналы','длительность','влияние','общий балл'])
    s.append(['P','T',1000,30,'=2',2,3,'=7'])
    path=tmp_path/'i.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'current_sheet':'P'})
    assert set(state['topics'][0]['uncached_formulas'])=={'signals_score','score'}
    state.update(PrioritizationAgent()(state));state.update(AnalyticsAgent()(state))
    result=state['topics'][0]
    assert result['signals_score_input'] is None and result['signals_score']==2 and result['signals_score_origin']=='methodology'
    assert result['score']==7 and result['score_origin']=='sum_of_criteria' and result['zone']=='потенциальный риск'
    assert any('нет сохранённых результатов формул' in w for w in state['warnings'])


def test_uncached_mau_formula_is_not_treated_as_absent_mau(tmp_path):
    from openpyxl import Workbook
    from golos_mas.agents import ExcelAgent, AnalyticsAgent
    w=Workbook();s=w.active;s.title='P'
    s.append(['продукт','тема','MAU','кол-во сигналов','сигналы','длительность','влияние','общий балл'])
    s.append(['P','T','=25000',30,None,2,3,None]);path=tmp_path/'i.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'current_sheet':'P'});state.update(PrioritizationAgent()(state));state.update(AnalyticsAgent()(state))
    result=state['topics'][0]
    assert 'mau' in result['uncached_formulas']
    assert result['methodology_check']['reason']=='uncached_mau'
    assert result['signals_score'] is None and result['score'] is None and result['zone']=='не определена'


def test_previous_period_uncached_formulas_are_reported(tmp_path):
    from openpyxl import Workbook
    from golos_mas.agents import ExcelAgent
    w=Workbook();current=w.active;current.title='текущий'
    current.append(['продукт','тема','кол-во сигналов','общий балл']);current.append(['P','T',40,9])
    previous=w.create_sheet('предыдущий')
    previous.append(['продукт','тема','кол-во сигналов','общий балл']);previous.append(['P','T','=30','=8'])
    path=tmp_path/'i.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'current_sheet':'текущий','previous_sheet':'предыдущий'})
    assert set(state['previous_topics'][0]['uncached_formulas'])=={'count','score'}
    assert any('предыдущий!2' in warning and 'нет сохранённых результатов формул' in warning for warning in state['warnings'])


def test_ux_and_invalid_entered_score_are_not_replaced():
    base={'sheet':'P','mau':5000,'count':50,'duration':1,'impact':3}
    items=[dict(base,row=2,score='UX',signals_score=None),dict(base,row=3,score=9,signals_score='ошибка')]
    result=PrioritizationAgent()({'topics':items,'warnings':[]})['topics']
    assert result[0]['signals_score'] is None and result[1]['signals_score']=='ошибка'

def test_truly_missing_mau_uses_medium_default_and_preserves_reason():
    item={'sheet':'P','row':2,'mau':None,'count':30,'signals_score':None,'duration':2,'impact':3,'score':None}
    result=PrioritizationAgent()({'topics':[item],'warnings':[]})
    topic=result['topics'][0]
    assert topic['methodology_check']['audience_basis']=='default_missing' and topic['signals_score']==2
    assert any('средняя аудитория' in warning for warning in result['warnings'])


def test_explicit_ambiguous_boundary_prevents_derivation():
    item={'sheet':'P','row':2,'mau':20000,'count':70,'signals_score':None,'duration':2,'impact':3,'score':None}
    result=PrioritizationAgent()({'topics':[item],'warnings':[],'mau_boundaries':'ambiguous'})['topics'][0]
    assert result['methodology_check']['reason']=='ambiguous_boundary' and result['signals_score'] is None
