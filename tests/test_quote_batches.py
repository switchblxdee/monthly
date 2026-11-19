import json
from types import SimpleNamespace
from golos_mas.agents import EvidenceAgent, QuoteSelectionAgent


def record(i): return {'text':f'Обращение {i}','source':f'Канал {i%2}','sheet':'SW','row':i,'likes':100-i}
def topic(n):
    candidates=[record(i) for i in range(1,n+1)]
    return {'sheet':'P','row':2,'topic':'Проблема','zone':'высокое внимание','quote_candidates':candidates,'quotes':candidates[:3]}


def test_model_considers_all_batches_including_low_reaction_tail():
    class Model:
        def __init__(self): self.seen=[]
        def invoke(self,messages):
            payload=json.loads(messages[1].content);batch=payload['кандидаты'];self.seen.extend(x['текст'] for x in batch)
            assert all(set(item)=={'номер','источник','текст'} for item in batch)
            return SimpleNamespace(content=json.dumps({'номера':[len(batch)]}))
    model=Model();result=QuoteSelectionAgent(model,True)({'topics':[topic(25)],'warnings':[]})
    chosen=result['topics'][0]
    assert {q['row'] for q in chosen['quotes']}=={12,24,25}
    assert 'Обращение 25' in model.seen and chosen['quote_selection_stats']['initial_candidates_attempted']==25
    assert chosen['quote_selection_stats']['model_calls']==3 and not chosen['quote_selection_stats']['truncated']

def test_quote_limit_reranks_survivors_from_different_batches():
    from golos_mas.agents import QuoteLimitAgent, maximum_quote_model_calls
    class Model:
        def __init__(self): self.calls=0
        def invoke(self,messages):
            self.calls+=1
            assert 'по убыванию соответствия' in messages[0].content
            return SimpleNamespace(content=json.dumps({'номера':[3,2,1] if self.calls==4 else [1]}))
    model=Model();state={'topics':[topic(25)],'warnings':[],'max_quotes_per_topic':1}
    state.update(QuoteSelectionAgent(model,True)(state))
    state.update(QuoteLimitAgent()(state))
    assert model.calls==4 and state['topics'][0]['quotes'][0]['row']==25
    assert maximum_quote_model_calls(25,final_limit=1)>=model.calls
    assert state['topics'][0]['omitted_quote_references']==[{'sheet':'SW','row':13},{'sheet':'SW','row':1}]

def test_rerank_failure_discards_partial_model_selection_before_quote_limit():
    from golos_mas.agents import QuoteLimitAgent
    class Model:
        def __init__(self): self.calls=0
        def invoke(self,messages):
            self.calls+=1
            if self.calls==4: raise TimeoutError('private endpoint')
            return SimpleNamespace(content='{"номера":[1]}')
    state={'topics':[topic(25)],'warnings':[],'max_quotes_per_topic':1}
    state.update(QuoteSelectionAgent(Model(),True)(state));state.update(QuoteLimitAgent()(state))
    assert state['topics'][0]['quote_selection']=='likes_fallback'
    assert [q['row'] for q in state['topics'][0]['quotes']]==[1]
    assert state['topics'][0]['quote_selection_stats']['model_calls']==4
    assert any('отбор GigaChat не удался' in warning for warning in state['warnings'])
    assert all('private endpoint' not in warning for warning in state['warnings'])

def test_model_call_estimate_never_understates_possible_reductions():
    from functools import lru_cache
    from golos_mas.agents import maximum_quote_model_calls
    @lru_cache(None)
    def worst(size,limit):
        if not size: return 0
        batches=(size+11)//12
        largest=(size//12)*3+min(3,size%12)
        outcomes=[]
        for selected in range(largest+1):
            if selected<=3:
                outcomes.append(batches+int(size>12 and selected>limit))
            elif selected<size:
                outcomes.append(batches+worst(selected,limit))
            else:
                outcomes.append(batches)
        return max(outcomes)
    for limit in (1,2,3):
        for size in range(1,151):
            assert maximum_quote_model_calls(size,final_limit=limit)>=worst(size,limit)


def test_reduction_is_bounded_and_returns_existing_rows():
    class Model:
        def invoke(self,messages):
            size=len(json.loads(messages[1].content)['кандидаты'])
            return SimpleNamespace(content=json.dumps({'номера':list(range(1,min(size,3)+1))}))
    result=QuoteSelectionAgent(Model(),True)({'topics':[topic(50)],'warnings':[]})['topics'][0]
    assert len(result['quotes'])==3 and all(1<=q['row']<=50 for q in result['quotes'])
    assert result['quote_selection_stats']['model_calls']==8 and result['quote_selection_stats']['initial_candidates_attempted']==50


def test_batch_failure_discards_partial_model_selection():
    class Model:
        def __init__(self): self.calls=0
        def invoke(self,messages):
            self.calls+=1
            if self.calls==2: raise TimeoutError('private endpoint')
            return SimpleNamespace(content='{"номера":[12]}')
    result=QuoteSelectionAgent(Model(),True)({'topics':[topic(30)],'warnings':[]})
    t=result['topics'][0]
    assert t['quote_selection']=='likes_fallback' and [q['row'] for q in t['quotes']]==[1,2,3]
    assert t['quote_selection_stats']['initial_candidates_attempted']==24 and t['quote_selection_stats']['initial_candidates_validated']==12 and t['quote_selection_stats']['model_calls']==2
    assert 'private endpoint' not in ' '.join(result['warnings'])


def test_explicit_limit_is_visible_in_audit_data():
    class Model:
        def invoke(self,messages): return SimpleNamespace(content='{"номера":[1]}')
    result=QuoteSelectionAgent(Model(),True,12)({'topics':[topic(25)],'warnings':[]})
    stats=result['topics'][0]['quote_selection_stats']
    assert stats['truncated'] and stats['eligible_after_limit']==12 and stats['initial_candidates_attempted']==12
    assert any('12 из 25' in w for w in result['warnings'])


def test_identical_text_in_different_sources_remains_distinct():
    matches=[{'text':'Одинаковый текст','source':'Канал A','sheet':'SW','row':2,'likes':0},
             {'text':'Одинаковый текст','source':'Канал B','sheet':'SW','row':3,'likes':0},
             {'text':'Одинаковый текст','source':'Канал A','sheet':'SW','row':4,'likes':0}]
    result=EvidenceAgent()({'topics':[{'matches':matches}]})['topics'][0]
    assert result['quote_text_rows']==3 and result['quote_candidate_count']==2
    assert {q['source'] for q in result['quote_candidates']}=={'Канал A','Канал B'}


def test_cli_audit_keeps_source_references_without_copying_all_candidates(tmp_path):
    import os,subprocess,sys
    from pathlib import Path
    from openpyxl import Workbook
    import golos_mas
    w=Workbook();p=w.active;p.title='P';p.append(['продукт','тема','общий балл']);p.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['продукт','тема','источник','Маскированный текст из источника','лайки','сигнал'])
    for i in range(5): d.append(['P','T','Чат',f'Текст обращения {i}',10-i,1])
    path=tmp_path/'book.xlsx';w.save(path);out=tmp_path/'report.pptx'
    env=dict(os.environ,PYTHONPATH=str(Path(golos_mas.__file__).resolve().parent.parent))
    run=subprocess.run([sys.executable,'-m','golos_mas.cli',str(path),'--current-sheet','P','--period','месяц','--output',str(out)],env=env,capture_output=True,text=True,timeout=120)
    assert run.returncode==0,run.stderr
    audit=json.loads(out.with_suffix('.audit.json').read_text(encoding='utf8'));t=audit['topics'][0]
    assert 'matches' not in t and 'quote_candidates' not in t
    assert len(t['matched_references'])==5 and len(t['quotes'])==3 and t['quote_candidate_count']==5
    assert 'Текст обращения 4' not in json.dumps(audit,ensure_ascii=False)


def test_worst_case_call_estimate_includes_reduction():
    from golos_mas.agents import maximum_quote_model_calls
    assert maximum_quote_model_calls(0)==0
    assert maximum_quote_model_calls(1)==1
    assert maximum_quote_model_calls(12)==1
    assert maximum_quote_model_calls(25)==4
    assert maximum_quote_model_calls(50)==8
