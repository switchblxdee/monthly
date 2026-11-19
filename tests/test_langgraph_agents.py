import json
import os
from types import SimpleNamespace

import pytest
from langchain_core.messages import AIMessage, HumanMessage, ToolMessage
from openpyxl import Workbook

from golos_mas.graph import create_graph
from golos_mas.review_agents import ToolReviewAgent
from golos_mas.worker_agents import ReportWorker


class Model:
    def __init__(self):
        self.match_calls = 0
        self.quote_calls = 0
        self.tool_results = []

    def bind_tools(self, tools):
        model = self

        class Bound:
            def invoke(self, messages):
                if isinstance(messages[-1], HumanMessage):
                    return AIMessage(content='', tool_calls=[{'name': 'inspect_overview', 'args': {}, 'id': 'call-1', 'type': 'tool_call'}])
                assert isinstance(messages[-1], ToolMessage)
                model.tool_results.append(messages[-1].content)
                data = json.loads(messages[-1].content)
                matching = 'сопоставление' in messages[0].content
                should_retry = data['without_matches'] > 0 if matching else data['without_quotes'] > 0
                return AIMessage(content=json.dumps({'decision': 'retry' if should_retry else 'accept'}))
        return Bound()

    def invoke(self, messages):
        prompt = messages[0].content
        if 'Выбери только варианты' in prompt:
            self.match_calls += 1
            return SimpleNamespace(content='{"номера":[]}' if self.match_calls == 1 else '{"номера":[1]}')
        if 'Выбери до трёх номеров обращений' in prompt:
            self.quote_calls += 1
            return SimpleNamespace(content='{"номера":[]}' if self.quote_calls == 1 else '{"номера":[1]}')
        return SimpleNamespace(content='{"номера":[1]}')


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE')
def test_agents_call_tools_and_route_real_pipeline_through_retries(tmp_path):
    w = Workbook(); p = w.active; p.title = 'P'
    p.append(['владелец', 'продукт', 'тема', 'кол-во сигналов', 'общий балл'])
    p.append(['Команда', 'ЦОД Балаково', 'Недоступность инфраструктуры Балаково', 12, 9])
    d = w.create_sheet('прочее')
    d.append(['дата', 'источник', 'продукт', 'тема', 'Маскированный текст из источника', 'лайки', 'сигнал'])
    d.append(['01.11.2026', 'Чат', 'DropApp', 'Сбой инфраструктуры Балаково', 'Нет доступа', 9, 10])
    book = tmp_path / 'input.xlsx'; w.save(book)
    model = Model()
    result = create_graph(model, use_model_matching=True, use_model_quotes=True, agentic=False).invoke({
        'workbook': str(book), 'template': os.environ['GOLOS_TEST_TEMPLATE'], 'current_sheet': 'P',
        'period': 'ноябрь', 'output': str(tmp_path / 'result.pptx')})
    assert [(d['agent'], d['action']) for d in result['agent_decisions']] == [
        ('matching', 'retry'), ('matching', 'accept'), ('evidence', 'retry'), ('evidence', 'accept')]
    assert all(d['tool_calls'] == 1 and d['model_calls'] == 2 for d in result['agent_decisions'])
    assert model.match_calls == 2 and model.quote_calls == 2
    assert result['topics'][0]['quotes'][0]['text'] == 'Нет доступа'
    assert result['coverage']['quote_model_calls_attempted'] == 2


def test_topic_tool_hides_quote_text_without_explicit_permission():
    class InspectModel:
        def __init__(self): self.seen = ''
        def bind_tools(self, tools):
            model = self
            class Bound:
                def invoke(self, messages):
                    if isinstance(messages[-1], HumanMessage):
                        return AIMessage(content='', tool_calls=[{'name': 'inspect_topic', 'args': {'number': 1}, 'id': 'c', 'type': 'tool_call'}])
                    model.seen = messages[-1].content
                    return AIMessage(content='{"decision":"accept"}')
            return Bound()
    topic = {'product': 'P', 'topic': 'T', 'matches': [], 'quotes': [{'text': 'PRIVATE', 'source': 'S'}]}
    for permission in (False, True):
        model = InspectModel()
        result = ToolReviewAgent(model, 'evidence', allow_text=permission)({'topics': [topic]})
        assert result['agent_decisions'][0]['action'] == 'accept'
        assert ('PRIVATE' in model.seen) is permission


def test_reviewer_failure_is_visible_and_cannot_request_retry():
    class NoTools: pass
    result = ToolReviewAgent(NoTools(), 'matching', retry_allowed=True)({'topics': [{'matches': []}]})
    assert result['matching_review_action'] == 'accept'
    assert result['agent_decisions'][0]['action'] == 'manual_review'
    assert result['warnings']


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE')
def test_four_agents_build_and_publish_deck_with_optional_retries(tmp_path):
    class TeamModel(Model):
        def bind_tools(self, tools):
            available = {t.name for t in tools}
            class Bound:
                def __init__(self): self.inspected = False
                def invoke(self, messages):
                    latest = json.loads(messages[-1].content)
                    status = latest.get('status', latest)
                    completed = status['completed']; remaining = status['remaining']
                    if not self.inspected and completed:
                        name = 'inspect_report'; self.inspected = True
                    elif 'retry_matching' in available and 'analyze_topics' in completed and 'retry_matching' not in completed:
                        name = 'retry_matching'
                    elif 'retry_quotes' in available and 'select_quotes' in completed and 'retry_quotes' not in completed:
                        name = 'retry_quotes'
                    else:
                        name = remaining[0]
                    return AIMessage(content='', tool_calls=[{'name': name, 'args': {}, 'id': f'call-{len(messages)}', 'type': 'tool_call'}])
            return Bound()
    w=Workbook(); p=w.active; p.title='P'
    p.append(['владелец','продукт','тема','кол-во сигналов','общий балл'])
    p.append(['Команда','ЦОД Балаково','Недоступность инфраструктуры Балаково',12,9])
    d=w.create_sheet('прочее')
    d.append(['дата','источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    d.append(['01.11.2026','Чат','DropApp','Сбой инфраструктуры Балаково','Нет доступа',9,10])
    book=tmp_path/'input.xlsx'; w.save(book); out=tmp_path/'result.pptx'; model=TeamModel()
    result=create_graph(model,use_model_matching=True,use_model_quotes=True).invoke({
        'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P',
        'period':'ноябрь','output':str(out)})
    decisions=result['agent_decisions']
    assert [x['agent'] for x in decisions]==['data','analyst','editor','presentation']
    assert 'retry_matching' in decisions[1]['tools'] and 'retry_quotes' in decisions[2]['tools']
    assert decisions[3]['tools']==['build_presentation','inspect_report','validate_presentation','save_presentation']
    assert result['topics'][0]['quotes'][0]['text']=='Нет доступа'
    assert result['coverage']['quote_model_calls_attempted']==2 and out.is_file()


def test_worker_cannot_skip_validation_or_loop_forever():
    invoked=[]
    class ModelIgnoringDependencies:
        def bind_tools(self, tools):
            class Bound:
                def invoke(self, messages):
                    return AIMessage(content='', tool_calls=[{'name':'save','args':{},'id':str(len(messages)),'type':'tool_call'}])
            return Bound()
    def validation(state): invoked.append('validate'); return {}
    def saving(state): invoked.append('save'); return {}
    worker=ReportWorker(ModelIgnoringDependencies(), 'presentation', 'собрать отчёт', [
        ('validate','Проверить отчёт.',validation),('save','Сохранить отчёт.',saving)])
    with pytest.raises(RuntimeError, match='превышен лимит'):
        worker({})
    assert invoked==[]
