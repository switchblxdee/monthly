from typing import TypedDict, Any
from pathlib import Path
import os
import tempfile
from langgraph.graph import StateGraph, START, END
from .agents import ExcelAgent, AnalyticsAgent, EvidenceAgent, QuoteSelectionAgent, QuoteLimitAgent, CoverageAgent, NarrativeAgent
from .render import PresentationAgent
from .template_render import TemplatePresentationAgent
from .quality import QualityAgent
from .statuses import StatusAgent
from .promises import PromisesAgent
from .topic_mapping import TopicMappingAgent
from .semantic_matching import SemanticMatchingAgent
from .methodology import PrioritizationAgent
from .review_agents import ToolReviewAgent
from .worker_agents import ReportWorker

class ReportState(TypedDict, total=False):
    schema_json: str | None
    mau_boundaries: str
    topic_map_json: str | None
    semantic_map: dict
    topic_map: dict
    promises_json: str | None
    promise_report: dict[str, Any]
    coverage: dict[str, Any]
    statuses_json: str | None
    template: str | None
    workbook: str
    current_sheet: str | None
    auto_previous: bool
    previous_sheet: str | None
    detail_sheets: list[str] | None
    report_month: str | None
    max_quotes_per_topic: int
    date_filter: dict[str, Any] | None
    skipped_rows: list[dict[str, Any]]
    period: str
    output: str
    staging_target: str
    topics: list[dict[str, Any]]
    previous_topics: list[dict[str, Any]]
    signals: list[dict[str, Any]]
    warnings: list[str]
    narrative: str
    narrative_sources: list[dict[str, Any]]
    slide_count: int
    agent_decisions: list[dict[str, Any]]
    matching_review_attempts: int
    matching_review_action: str
    evidence_review_attempts: int
    evidence_review_action: str

class SelectPresentationAgent:
    def __call__(self, state):
        if (state.get('promise_report') or state.get('statuses_json')) and not state.get('template'):
            raise ValueError('Для формирования обещаний или значков статуса укажите --template')
        target=Path(state['output']).resolve()
        sources=[Path(state[k]).resolve() for k in ('workbook','template','statuses_json','promises_json','topic_map_json','schema_json') if state.get(k)]
        if target in sources: raise ValueError('Выходная презентация совпадает с исходным файлом')
        target.parent.mkdir(parents=True,exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix='.golos-',suffix='.pptx',dir=target.parent,delete=False) as tmp:
            staging=Path(tmp.name)
        try:
            staged_state=dict(state,output=str(staging))
            result=(TemplatePresentationAgent() if state.get('template') else PresentationAgent())(staged_state)
            return {**result,'staging_target':str(target)}
        except BaseException:
            staging.unlink(missing_ok=True)
            raise

class PublishAgent:
    def __call__(self,state):
        staging=Path(state['output']); target=Path(state['staging_target'])
        try: os.replace(staging,target)
        except BaseException:
            staging.unlink(missing_ok=True)
            raise
        return {'output':str(target)}

class QuoteSelectionStage:
    def __init__(self, llm, enabled, maximum):
        self.delegate = QuoteSelectionAgent(llm, enabled, maximum)

    def __call__(self, state):
        result = self.delegate(state)
        if result and state.get('evidence_review_action') == 'retry':
            for previous, current in zip(state['topics'], result['topics']):
                history = list(previous.get('quote_attempt_history', []))
                if previous.get('quote_selection_stats'):
                    history.append(previous['quote_selection_stats'])
                current['quote_attempt_history'] = history
        return result

def create_graph(llm=None, use_model_quotes=False, use_model_matching=False, max_quote_candidates=0, agentic=True):
    if agentic and llm is not None:
        if not callable(getattr(llm, 'bind_tools', None)):
            raise ValueError('Для агентного режима модель должна поддерживать bind_tools')
        return create_agent_team(llm, use_model_quotes, use_model_matching, max_quote_candidates)
    graph=StateGraph(ReportState)
    for name,agent in [('excel',ExcelAgent()),('prioritization',PrioritizationAgent()),('topic_mapping',TopicMappingAgent()),('semantic_matching',SemanticMatchingAgent(llm,use_model_matching)),('analytics',AnalyticsAgent()),('matching_review',ToolReviewAgent(llm,'matching',retry_allowed=use_model_matching)),('statuses',StatusAgent()),('promises',PromisesAgent()),('evidence',EvidenceAgent()),('quote_selection',QuoteSelectionStage(llm,use_model_quotes,max_quote_candidates)),('evidence_review',ToolReviewAgent(llm,'evidence',allow_text=use_model_quotes,retry_allowed=use_model_quotes)),('quote_limit',QuoteLimitAgent()),('coverage',CoverageAgent()),('narrative',NarrativeAgent(llm)),('presentation',SelectPresentationAgent()),('quality',QualityAgent()),('publish',PublishAgent())]:
        graph.add_node(name,agent)
    graph.add_edge(START,'excel')
    for left,right in [('excel','prioritization'),('prioritization','topic_mapping'),('topic_mapping','semantic_matching'),('semantic_matching','analytics'),('analytics','matching_review'),('statuses','promises'),('promises','evidence'),('evidence','quote_selection'),('quote_selection','evidence_review'),('quote_limit','coverage'),('coverage','narrative'),('narrative','presentation'),('presentation','quality'),('quality','publish'),('publish',END)]:
        graph.add_edge(left,right)
    graph.add_conditional_edges('matching_review',lambda state: state.get('matching_review_action','accept'),{'retry':'semantic_matching','accept':'statuses'})
    graph.add_conditional_edges('evidence_review',lambda state: state.get('evidence_review_action','accept'),{'retry':'quote_selection','accept':'quote_limit'})
    return graph.compile()

def create_agent_team(llm, use_model_quotes=False, use_model_matching=False, max_quote_candidates=0):
    """Four tool-using workers share the typed report state and build the PPTX."""
    graph = StateGraph(ReportState)
    selector = QuoteSelectionStage(llm, use_model_quotes, max_quote_candidates)

    def calculate(state):
        updates = SemanticMatchingAgent(llm, use_model_matching)(state)
        intermediate = dict(state, **updates)
        return {**updates, **AnalyticsAgent()(intermediate)}

    def retry_matching(state):
        return calculate(state)

    def choose_quotes(state):
        return selector(state)

    def retry_quotes(state):
        return selector(dict(state, evidence_review_action='retry'))

    workers = [
        ReportWorker(llm, 'data', 'прочитай книгу и подготовь данные и методику', [
            ('read_workbook', 'Прочитать листы Excel и проверить структуру.', ExcelAgent()),
            ('apply_prioritization', 'Применить методику приоритета к темам.', PrioritizationAgent()),
            ('prepare_mapping', 'Подготовить заданные сопоставления тем.', TopicMappingAgent())]),
        ReportWorker(llm, 'analyst', 'сопоставь темы с сигналами и рассчитай показатели', [
            ('analyze_topics', 'Сопоставить темы с детализацией и вычислить показатели.', calculate),
            ('apply_statuses', 'Применить предоставленные статусы тем.', StatusAgent()),
            ('analyze_promises', 'Обработать предоставленный реестр обещаний.', PromisesAgent())],
            [('retry_matching', 'При несопоставленных темах повторно сопоставить их; используй до применения статусов.', retry_matching, ['analyze_topics'], ['apply_statuses'])] if use_model_matching else []),
        ReportWorker(llm, 'editor', 'подготовь подтверждённые цитаты и текст сводки', [
            ('collect_evidence', 'Собрать обращения и кандидатов для цитат.', EvidenceAgent()),
            ('select_quotes', 'Выбрать цитаты из исходных обращений.', choose_quotes),
            ('apply_quote_limit', 'Применить ограничение числа цитат.', QuoteLimitAgent()),
            ('calculate_coverage', 'Рассчитать полноту и происхождение данных.', CoverageAgent()),
            ('prepare_summary', 'Подготовить сводку по подтверждённым темам.', NarrativeAgent(llm))],
            [('retry_quotes', 'Повторить выбор цитат, если необходимо; используй до ограничения цитат.', retry_quotes, ['select_quotes'], ['apply_quote_limit'])] if use_model_quotes else []),
        ReportWorker(llm, 'presentation', 'собери, проверь и сохрани готовую презентацию', [
            ('build_presentation', 'Создать временную презентацию по шаблону и данным отчёта.', SelectPresentationAgent()),
            ('validate_presentation', 'Проверить структуру PPTX и содержимое отчёта.', QualityAgent()),
            ('save_presentation', 'Сохранить проверенную презентацию по указанному пути.', PublishAgent())]),
    ]
    for worker in workers:
        graph.add_node(worker.name, worker)
    graph.add_edge(START, 'data')
    graph.add_edge('data', 'analyst')
    graph.add_edge('analyst', 'editor')
    graph.add_edge('editor', 'presentation')
    graph.add_edge('presentation', END)
    return graph.compile()
