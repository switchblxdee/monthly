from types import SimpleNamespace
from golos_mas.semantic_matching import SemanticMatchingAgent
from golos_mas.agents import AnalyticsAgent, EvidenceAgent

class Model:
    def __init__(self,response): self.response=response;self.calls=[]
    def invoke(self,messages): self.calls.append(messages);return SimpleNamespace(content=self.response)

def topic(product,text,row=2):
    return {'product':product,'topic':text,'score':9,'signals_score':None,'duration':None,'impact':None,'sheet':'P','row':row,'count':None}

def detail(product,text,row=2):
    return {'product':product,'topic':text,'text':'Текст из источника','source':'Чат','sheet':'D','row':row,'signal':1,'likes':0}

def test_model_matches_only_numbered_existing_candidates():
    p=topic('ЦОД Балаково','Недоступность инфраструктуры Балаково')
    d=detail('DropApp','Сбой инфраструктуры Балаково после отключения')
    model=Model('{"номера":[1]}');state={'topics':[p],'signals':[d],'previous_topics':[],'warnings':[]}
    state.update(SemanticMatchingAgent(model,True)(state))
    state.update(AnalyticsAgent()(state));state.update(EvidenceAgent()(state))
    assert state['topics'][0]['match_method']=='gigachat' and state['topics'][0]['quotes'][0]['row']==2
    assert 'Текст из источника' not in str(model.calls[0][1].content)
    assert any('требует проверки' in w for w in state['warnings'])

def test_invalid_model_numbers_do_not_create_matches():
    p=topic('ЦОД Балаково','Недоступность инфраструктуры Балаково');d=detail('DropApp','Сбой инфраструктуры Балаково')
    state={'topics':[p],'signals':[d],'previous_topics':[],'warnings':[]}
    state.update(SemanticMatchingAgent(Model('{"номера":[999]}'),True)(state));state.update(AnalyticsAgent()(state))
    assert not state['topics'][0]['matches'] and any('сопоставление GigaChat не удалось' in w for w in state['warnings'])

def test_exact_matches_are_reserved_and_never_sent_as_alternatives():
    exact=topic('DropApp','Сбой инфраструктуры Балаково',2)
    other=topic('ЦОД Балаково','Недоступность инфраструктуры Балаково',3)
    d=detail('DropApp','Сбой инфраструктуры Балаково')
    model=Model('{"номера":[1]}');state={'topics':[exact,other],'signals':[d],'previous_topics':[],'warnings':[]}
    state.update(SemanticMatchingAgent(model,True)(state));state.update(AnalyticsAgent()(state))
    assert len(model.calls)==0 and state['topics'][0]['match_method']=='exact' and not state['topics'][1]['matches']

def test_two_model_topics_claiming_same_detail_are_both_blocked():
    p1=topic('ЦОД Балаково','Недоступность инфраструктуры Балаково',2)
    p2=topic('ЦОД','Отключение инфраструктуры Балаково',3)
    d=detail('DropApp','Сбой инфраструктуры Балаково после отключения')
    state={'topics':[p1,p2],'signals':[d],'previous_topics':[],'warnings':[]}
    state.update(SemanticMatchingAgent(Model('{"номера":[1]}'),True)(state));state.update(AnalyticsAgent()(state))
    assert all(t['match_method']=='ambiguous' and not t['matches'] for t in state['topics'])

def test_candidate_limit_and_model_failure_are_visible_without_source_text():
    class FailingModel:
        def invoke(self,messages): raise TimeoutError('sensitive server detail')
    p=topic('Продукт','Недоступность сервиса')
    rows=[detail('Продукт',f'Сбой компонента {i}',i+2) for i in range(30)]
    state={'topics':[p],'signals':rows,'previous_topics':[],'warnings':[]}
    state.update(SemanticMatchingAgent(FailingModel(),True)(state))
    assert not state['semantic_map']
    assert any('25 из 30 кандидатов' in w for w in state['warnings'])
    assert any('TimeoutError' in w for w in state['warnings'])
    assert all('sensitive server detail' not in w for w in state['warnings'])

def test_no_candidates_is_reported_without_calling_model():
    model=Model('{"номера":[]}')
    state={'topics':[topic('Продукт','Недоступность сервиса')],'signals':[], 'warnings':[]}
    state.update(SemanticMatchingAgent(model,True)(state))
    assert not model.calls and any('нет доступных кандидатов' in w for w in state['warnings'])
