from openpyxl import Workbook
from golos_mas.graph import create_graph
from zipfile import ZipFile

def test_full_pipeline(tmp_path):
    w=Workbook(); cur=w.active; cur.title='приоритет октябрь'
    cur.append(['владелец','продукт','тема','MAU','кол-во сигналов','сигналы','длительность','влияние','общий балл'])
    cur.append(['SberWorks','GigaCode','Ошибка',30000,3,1,2,3,9])
    old=w.create_sheet('приоритет сентябрь');old.append(['продукт','тема']);old.append(['GigaCode','Ошибка'])
    ws=w.create_sheet('SW');ws.append(['дата','источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    ws.append(['01.10.2026','Чат','GigaCode','Ошибка','Пример обращения',2,3])
    source=tmp_path/'input.xlsx';w.save(source);output=tmp_path/'result.pptx'
    state=create_graph().invoke({'workbook':str(source),'current_sheet':cur.title,'previous_sheet':old.title,'period':'октябрь 2026','output':str(output)})
    topic=state['topics'][0]
    assert topic['zone']=='высокое внимание' and topic['recurrence']=='есть в прошлом месяце'
    assert topic['signal_sum']==3 and topic['quotes'][0]['row']==2
    assert output.exists() and state['slide_count']==5
    with ZipFile(output) as ppt: assert 'ppt/presentation.xml' in ppt.namelist()

def test_ambiguous_matches_are_not_reused():
    from golos_mas.agents import AnalyticsAgent, EvidenceAgent
    topic='Проблема запуска приложения после обновления операционной системы'
    base={'owner':'A','product':'P','topic':topic,'score':9,'count':1,'sheet':'приоритет','duration':1}
    detail={'product':'P','topic':topic,'text':'тест','signal':1,'likes':0,'source':'чат','sheet':'SW','row':2}
    result=AnalyticsAgent()({'topics':[dict(base,row=2),dict(base,row=3)],'signals':[detail],'previous_topics':[],'warnings':[]})
    result.update(EvidenceAgent()(result))
    assert all(t['match_method']=='ambiguous' and not t['quotes'] and t['signal_sum'] is None for t in result['topics'])
    assert any('Неоднозначное' in w for w in result['warnings'])


def test_month_filter_excludes_unreadable_dates(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook(); ws=w.active;ws.title='priorities';ws.append(['продукт','тема']);ws.append(['P','T'])
    s=w.create_sheet('SW');s.append(['дата','продукт','тема']);s.append(['01.10.2026','P','T']);s.append(['01.09.2026','P','T']);s.append(['неизвестно','P','T'])
    path=tmp_path/'input.xlsx';w.save(path)
    result=ExcelAgent()({'workbook':str(path),'current_sheet':'priorities','report_month':'2026-10'})
    assert len(result['signals'])==1 and result['signals'][0]['row']==2
    assert any('Исключено 1' in w for w in result['warnings'])
    assert result['date_filter']['unknown_date_references']==[{'sheet':'SW','row':4}]
    assert result['date_filter']['other_month_references']==[{'sheet':'SW','row':3,'month':'2026-09'}]

def test_detail_uncached_formulas_are_distinguished_from_empty_cells(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['дата','продукт','тема','лайки','сигнал'])
    d.append(['01.10.2026','P','T','=2','=D2+1'])
    d.append(['02.10.2026','P','T',None,None])
    book=tmp_path/'i.xlsx';w.save(book)
    state=ExcelAgent()({'workbook':str(book),'current_sheet':'P','report_month':'2026-10'})
    assert state['signals'][0]['uncached_formulas']==['likes','signal']
    assert 'uncached_formulas' not in state['signals'][1]
    assert any('SW!2' in warning and 'likes, signal' in warning for warning in state['warnings'])

def test_formula_warnings_follow_reporting_month_and_include_unknown_dates(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['дата','продукт','тема','лайки','сигнал'])
    d.append(['01.09.2026','P','T','=2','=3'])
    d.append(['01.10.2026','P','T','=2','=3'])
    d.append(['=TODAY()','P','T',None,1])
    book=tmp_path/'i.xlsx';w.save(book)
    state=ExcelAgent()({'workbook':str(book),'current_sheet':'P','report_month':'2026-10'})
    formula_warnings=[warning for warning in state['warnings'] if 'нет сохранённых результатов формул' in warning]
    assert not any('SW!2' in warning for warning in formula_warnings)
    assert any('SW!3' in warning for warning in formula_warnings)
    assert any('SW!4' in warning and 'date' in warning for warning in formula_warnings)
    assert state['date_filter']['unknown_date_references']==[{'sheet':'SW','row':4}]

def test_scores_are_checked_without_overwriting_source():
    from golos_mas.agents import AnalyticsAgent
    topics=[{'product':'P','topic':'T','score':10,'signals_score':2,'duration':2,'impact':3,'sheet':'priorities','row':2,'count':1}]
    result=AnalyticsAgent()({'topics':topics,'signals':[],'previous_topics':[],'warnings':[]})
    assert result['topics'][0]['score']==10
    assert any('не совпадает с суммой критериев 7' in w for w in result['warnings'])

def test_missing_score_can_be_derived_from_complete_components():
    from golos_mas.agents import AnalyticsAgent
    topic={'product':'P','topic':'T','score':None,'signals_score':4,'duration':2,'impact':3,'sheet':'priorities','row':2,'count':1}
    result=AnalyticsAgent()({'topics':[topic],'signals':[],'previous_topics':[],'warnings':[]})
    assert result['topics'][0]['score']==9 and result['topics'][0]['score_origin']=='sum_of_criteria'
    assert result['topics'][0]['zone']=='высокое внимание'

def test_quote_model_can_only_choose_existing_source_rows():
    from types import SimpleNamespace
    from golos_mas.agents import QuoteSelectionAgent
    class Stub:
        def invoke(self, messages):
            return SimpleNamespace(content='{"номера": [2, 1]}')
    candidates=[{'text':'Оригинал 1','sheet':'SW','row':2}, {'text':'Оригинал 2','sheet':'SW','row':3}]
    topic={'sheet':'P','row':2,'topic':'Ошибка','zone':'высокое внимание','quote_candidates':candidates,'quotes':candidates,'quote_selection':'likes'}
    result=QuoteSelectionAgent(Stub(),True)({'topics':[topic],'warnings':[]})
    assert [q['row'] for q in result['topics'][0]['quotes']]==[3,2]
    assert result['topics'][0]['quote_selection']=='gigachat'


def test_quote_ranking_ignores_invalid_likes_and_reports_source_rows():
    from golos_mas.agents import EvidenceAgent
    bad={'sheet':'SW','row':2,'source':'A','text':'Некорректное','likes':True}
    good={'sheet':'SW','row':3,'source':'B','text':'Достоверное','likes':1}
    result=EvidenceAgent()({'topics':[{'matches':[bad,good]}],'warnings':[]})
    assert [q['row'] for q in result['topics'][0]['quotes']]==[3,2]
    assert [q['likes_for_ordering'] for q in result['topics'][0]['quotes']]==[1,0]
    assert result['topics'][0]['quotes'][1]['likes'] is True
    assert any('SW!2' in warning and 'некорректное число лайков' in warning for warning in result['warnings'])

def test_invalid_model_output_uses_deterministic_fallback():
    from types import SimpleNamespace
    from golos_mas.agents import QuoteSelectionAgent
    class Stub:
        def invoke(self,messages): return SimpleNamespace(content='{"номера": [999], "цитата": "выдуманный текст"}')
    quote={'text':'Оригинал','sheet':'SW','row':2}
    topic={'sheet':'P','row':2,'topic':'Ошибка','zone':'UX','quote_candidates':[quote],'quotes':[quote],'quote_selection':'likes'}
    result=QuoteSelectionAgent(Stub(),True)({'topics':[topic],'warnings':[]})
    assert result['topics'][0]['quotes']==[quote] and result['topics'][0]['quote_selection']=='likes_fallback'
    assert result['warnings']

def test_model_summary_failure_keeps_pipeline_usable():
    from golos_mas.agents import NarrativeAgent
    class Failed:
        def invoke(self,messages): raise TimeoutError('private endpoint details')
    state=NarrativeAgent(Failed())({'topics':[{'zone':'высокое внимание','score':9,'count':10,'product':'P','topic':'T','sheet':'P','row':2}],'warnings':[]})
    assert state['narrative']=='' and 'TimeoutError' in state['warnings'][0]
    assert 'private endpoint details' not in state['warnings'][0]

def test_optional_status_is_attached_only_to_exact_topic(tmp_path):
    import json
    from golos_mas.statuses import StatusAgent
    records=[{'product':'P','topic':'Ошибка','controls':['backlog'],'mitigation':'release','deadline':'4Q 2026'}]
    path=tmp_path/'statuses.json';path.write_text(json.dumps(records),encoding='utf8')
    result=StatusAgent()({'statuses_json':str(path),'topics':[{'product':'P','topic':'Ошибка','sheet':'P','row':2},{'product':'P','topic':'Другое','sheet':'P','row':3}],'warnings':[]})
    assert result['topics'][0]['status_record']['mitigation']=='release'
    assert result['topics'][0]['status_source']['record']==1
    assert 'status_record' not in result['topics'][1]


def test_duplicate_status_records_fail_closed(tmp_path):
    import json, pytest
    from golos_mas.statuses import StatusAgent
    path=tmp_path/'s.json';path.write_text(json.dumps([{'product':'P','topic':'T'},{'product':'P','topic':'T'}]),encoding='utf8')
    with pytest.raises(ValueError,match='повторяются'):
        StatusAgent()({'statuses_json':str(path),'topics':[],'warnings':[]})

def test_mau_header_with_approximation_sign_is_read(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='приоритет';s.append(['продукт','тема','~ MAU','общий балл']);s.append(['P','T',30000,9])
    path=tmp_path/'i.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'current_sheet':'приоритет'})
    assert state['topics'][0]['mau']==30000

def test_workbook_reports_missing_headers_and_partial_rows(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',9]);s.append(['P',None,5])
    d=w.create_sheet('SW');d.append(['продукт','тема']);d.append(['P','T'])
    path=tmp_path/'i.xlsx';w.save(path)
    result=ExcelAgent()({'workbook':str(path),'current_sheet':'P'})
    assert len(result['topics'])==1
    assert any('пропущено 1 строк' in x for x in result['warnings'])
    assert any('маскированный текст' in x and 'сигнал' in x for x in result['warnings'])

def test_detail_rows_without_both_join_keys_are_reported(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['продукт','тема','источник','маскированный текст из источника','сигнал'])
    d.append([None,None,'Чат','Обращение без продукта и темы',1])
    d.append(['  ','  ',None,None,None])
    path=tmp_path/'i.xlsx';w.save(path)
    result=ExcelAgent()({'workbook':str(path),'current_sheet':'P'})
    assert not result['signals']
    assert any('Лист SW: пропущено 1 строк' in warning and '2' in warning for warning in result['warnings'])
    assert result['skipped_rows']==[{'sheet':'SW','row':2,'kind':'detail','missing_fields':['product','topic']}]

def test_all_skipped_row_references_are_available_beyond_warning_examples(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['продукт','тема','маскированный текст из источника'])
    for i in range(15): d.append([None,None,f'Частное обращение {i}'])
    book=tmp_path/'i.xlsx';w.save(book)
    state=ExcelAgent()({'workbook':str(book),'current_sheet':'P'})
    assert [r['row'] for r in state['skipped_rows']]==list(range(2,17))
    assert all(set(r)=={'sheet','row','kind','missing_fields'} for r in state['skipped_rows'])

def test_skipped_row_identifies_uncached_key_formula_without_exposing_it(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['продукт','тема','маскированный текст из источника'])
    d.append(['="PRIVATE_PRODUCT"','T','PRIVATE_TEXT'])
    d.append([None,'T','PRIVATE_TEXT_2'])
    book=tmp_path/'i.xlsx';w.save(book)
    state=ExcelAgent()({'workbook':str(book),'current_sheet':'P'})
    assert state['skipped_rows'][0]['uncached_key_formulas']==['product']
    assert 'uncached_key_formulas' not in state['skipped_rows'][1]
    assert any('SW!2' in warning and 'формулы ключа' in warning for warning in state['warnings'])
    assert 'PRIVATE' not in str(state['skipped_rows']) and 'PRIVATE' not in str(state['warnings'])

def test_formula_only_rows_are_not_mistaken_for_empty_rows(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['продукт','тема','сигнал'])
    d.append(['="PRIVATE_PRODUCT"','="PRIVATE_TOPIC"',None])
    d.append([None,None,'=1'])
    d.append([None,None,None])
    book=tmp_path/'i.xlsx';w.save(book)
    state=ExcelAgent()({'workbook':str(book),'current_sheet':'P'})
    assert [row['row'] for row in state['skipped_rows']]==[2,3]
    assert state['skipped_rows'][0]['uncached_key_formulas']==['product','topic']
    assert 'uncached_key_formulas' not in state['skipped_rows'][1]
    assert 'PRIVATE' not in str(state['warnings']) and 'PRIVATE' not in str(state['skipped_rows'])

def test_mismatched_value_and_formula_ranges_fail_instead_of_dropping_rows():
    import pytest
    from golos_mas.agents import sheet_rows
    values=Workbook().active;values.append(['продукт','тема']);values.append(['P','T']);values.append(['P2','T2'])
    raw=Workbook().active;raw.append(['продукт','тема']);raw.append(['P','T'])
    with pytest.raises(ValueError,match='диапазоны значений и формул не совпадают'):
        list(sheet_rows(values,'detail',raw_ws=raw))


def test_coverage_counts_unique_rows_and_unmatched_references():
    from golos_mas.agents import CoverageAgent
    r1={'sheet':'SW','row':2};r2={'sheet':'SW','row':3};r3={'sheet':'SH','row':4}
    topics=[{'matches':[r1],'quotes':[r1],'match_method':'exact'},{'matches':[r1],'quotes':[r1],'match_method':'exact'}, {'matches':[],'quotes':[],'match_method':'none'}]
    out=CoverageAgent()({'topics':topics,'signals':[r1,r2,r3]})['coverage']
    assert out['detail_rows']==3 and out['matched_detail_rows']==1 and out['selected_quote_rows']==1
    assert out['unmatched_by_sheet']=={'SH':1,'SW':1}
    assert out['unmatched_references']==[{'sheet':'SH','row':4},{'sheet':'SW','row':3}]

def test_quote_coverage_distinguishes_selected_from_presented_rows():
    from golos_mas.agents import CoverageAgent
    r1={'sheet':'SW','row':2};r2={'sheet':'SW','row':3}
    topics=[{'zone':'высокое внимание','matches':[r1],'quotes':[r1],'match_method':'exact'},
            {'zone':'потенциальный риск','matches':[r2],'quotes':[r2],'match_method':'exact'}]
    coverage=CoverageAgent()({'signals':[r1,r2],'topics':topics,'warnings':[]})['coverage']
    assert coverage['selected_quote_rows']==2 and coverage['presentation_quote_rows']==1
    assert coverage['presentation_quote_references']==[{'sheet':'SW','row':2}]

def test_explicit_detail_sheets_and_self_comparison_are_validated(tmp_path):
    from golos_mas.agents import ExcelAgent
    import pytest
    w=Workbook();s=w.active;s.title='Приоритет';s.append(['продукт','тема']);s.append(['P','T'])
    d=w.create_sheet('Новый источник');d.append(['продукт','тема']);d.append(['P','T']);path=tmp_path/'i.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'current_sheet':'Приоритет','detail_sheets':['Новый источник']})
    assert len(state['signals'])==1
    with pytest.raises(ValueError,match='должны различаться'):
        ExcelAgent()({'workbook':str(path),'current_sheet':'Приоритет','previous_sheet':'Приоритет'})
    with pytest.raises(ValueError,match='несуществующий'):
        ExcelAgent()({'workbook':str(path),'current_sheet':'Приоритет','detail_sheets':['нет']})

def test_explicit_mapping_links_different_products_and_several_topics(tmp_path):
    import json
    from golos_mas.agents import ExcelAgent, AnalyticsAgent, EvidenceAgent
    from golos_mas.topic_mapping import TopicMappingAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл','кол-во сигналов']);s.append(['ЦОД','Инцидент',9,5])
    d=w.create_sheet('прочее');d.append(['продукт','тема','источник','Маскированный текст из источника','сигнал'])
    d.append(['Сервис A','Ошибка A','Чат A','Текст A',2]);d.append(['Сервис B','Ошибка B','Чат B','Текст B',3])
    path=tmp_path/'book.xlsx';w.save(path)
    mapping=[{'priority_product':'ЦОД','priority_topic':'Инцидент','detail_product':'Сервис A','detail_topic':'Ошибка A'}, {'priority_product':'ЦОД','priority_topic':'Инцидент','detail_product':'Сервис B','detail_topic':'Ошибка B'}]
    mp=tmp_path/'map.json';mp.write_text(json.dumps(mapping,ensure_ascii=False),encoding='utf8')
    state={'workbook':str(path),'current_sheet':'P','topic_map_json':str(mp)}
    state.update(ExcelAgent()(state));state.update(TopicMappingAgent()(state));state.update(AnalyticsAgent()(state));state.update(EvidenceAgent()(state))
    t=state['topics'][0]
    assert t['match_method']=='explicit_mapping' and t['signal_sum']==5
    assert {q['row'] for q in t['quotes']}=={2,3}


def test_mapping_rejects_reusing_detail_for_different_priorities(tmp_path):
    import json, pytest
    from golos_mas.topic_mapping import TopicMappingAgent
    a={'priority_product':'A','priority_topic':'T1','detail_product':'D','detail_topic':'X'}
    b={**a,'priority_topic':'T2'};path=tmp_path/'map.json';path.write_text(json.dumps([a,b]),encoding='utf8')
    with pytest.raises(ValueError,match='разным темам'):
        TopicMappingAgent()({'topic_map_json':str(path),'topics':[],'warnings':[]})

def test_explicit_recurrence_is_checked_and_not_inferred_as_new(tmp_path):
    import json
    from golos_mas.statuses import StatusAgent
    from golos_mas.template_render import recurrence_code
    topic={'product':'P','topic':'T','sheet':'P','row':2,'duration':4,'extra':{}}
    path=tmp_path/'s.json';path.write_text(json.dumps([{'product':'P','topic':'T','recurrence':'new'}]),encoding='utf8')
    state=StatusAgent()({'statuses_json':str(path),'topics':[topic],'warnings':[]})
    assert recurrence_code(state['topics'][0])=='new' and any('противоречит' in w for w in state['warnings'])
    assert recurrence_code({'duration':1,'extra':{}}) is None
    assert recurrence_code({'duration':1,'extra':{'примечание':'обсуждение не возобновилось'}}) is None

def test_duration_and_missing_previous_topic_warn_but_do_not_rewrite_icon():
    from golos_mas.agents import AnalyticsAgent
    from golos_mas.template_render import recurrence_code
    topic={'product':'P','topic':'T','sheet':'P','row':2,'score':9,'count':10,'duration':2,'signals_score':3,'impact':4}
    state={'topics':[topic],'signals':[],'previous_topics':[],'previous_sheet':'прошлый','warnings':[]}
    result=AnalyticsAgent()(state)
    assert recurrence_code(result['topics'][0])=='two'
    assert any('точная тема не найдена в предыдущем периоде' in w for w in result['warnings'])

def test_contents_product_uses_highest_explicit_zone():
    from golos_mas.template_render import contents_cards
    groups={'A':[{'product':'One','zone':'потенциальный риск'}, {'product':' one ','zone':'высокое внимание'}, {'product':'Two','zone':'UX'}, {'product':'Three','zone':'наблюдение'}]}
    cards=contents_cards(groups)
    assert cards==[[('A',[('One','высокое внимание'),('Three','наблюдение'),('Two',None)])]]

def test_contents_reserves_room_for_wrapped_product_names():
    from golos_mas.template_render import contents_cards
    products=[f'Продукт {i}: система доступа и управления инфраструктурой организации' for i in range(5)]
    cards=contents_cards({'A':[{'product':p,'zone':'высокое внимание'} for p in products]})
    flattened=[product for page in cards for _,items in page for product,_ in items]
    assert flattened==products
    assert len(cards[0][0][1])==len(products)

def test_contents_moves_long_product_from_short_card_to_next_page():
    from golos_mas.template_render import contents_cards
    products=[f'Продукт {i}: система доступа и управления инфраструктурой организации' for i in range(8)]
    pages=contents_cards({'A':[{'product':p,'zone':'высокое внимание'} for p in products]})
    flattened=[product for page in pages for _,items in page for product,_ in items]
    assert flattened==products and len(pages)==1 and len(pages[0])==1
    assert len(pages[0][0][1])==8

def test_failed_validation_does_not_replace_existing_deck(tmp_path,monkeypatch):
    import pytest
    from golos_mas.quality import QualityAgent
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',5]);book=tmp_path/'i.xlsx';w.save(book)
    output=tmp_path/'existing.pptx';output.write_bytes(b'old deck remains')
    def fail(self,state): raise ValueError('failure in quality checks')
    monkeypatch.setattr(QualityAgent,'_inspect',fail)
    with pytest.raises(ValueError,match='quality checks'):
        create_graph().invoke({'workbook':str(book),'current_sheet':'P','period':'месяц','output':str(output)})
    assert output.read_bytes()==b'old deck remains'
    assert not list(tmp_path.glob('.golos-*'))

def test_missing_libreoffice_explains_how_to_proceed_and_preserves_output(tmp_path,monkeypatch):
    import pytest
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',5]);book=tmp_path/'i.xlsx';w.save(book)
    output=tmp_path/'existing.pptx';output.write_bytes(b'previous report')
    monkeypatch.setenv('PATH',str(tmp_path/'no-executables'))
    with pytest.raises(RuntimeError,match='LibreOffice.*--template'):
        create_graph().invoke({'workbook':str(book),'current_sheet':'P','period':'месяц','output':str(output)})
    assert output.read_bytes()==b'previous report'
    assert not list(tmp_path.glob('.golos-*'))

def test_cli_without_gigachat_creates_deck_and_audit(tmp_path):
    import os, subprocess, sys, json
    from pathlib import Path
    import golos_mas
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',5])
    detail=w.create_sheet('SW');detail.append(['продукт','тема','маскированный текст из источника']);detail.append([None,None,'Пропущенное обращение'])
    book=tmp_path/'i.xlsx';w.save(book)
    output=tmp_path/'report.pptx';env=dict(os.environ,PYTHONPATH=str(Path(golos_mas.__file__).resolve().parent.parent))
    result=subprocess.run([sys.executable,'-m','golos_mas.cli',str(book),'--current-sheet','P','--period','октябрь','--output',str(output)],capture_output=True,text=True,env=env,timeout=120)
    assert result.returncode==0, result.stderr
    assert output.exists() and output.with_suffix('.audit.json').exists()
    from golos_mas.verify import verify
    assert verify(output)
    with output.open('ab') as deck: deck.write(b'change')
    assert not verify(output)
    audit=json.loads(output.with_suffix('.audit.json').read_text(encoding='utf8'))
    assert audit['coverage']['priority_topics']==1
    assert audit['skipped_rows']==[{'sheet':'SW','row':2,'kind':'detail','missing_fields':['product','topic']}]
    assert not list(tmp_path.glob('.golos-*'))

def test_quality_rejects_missing_package_part(tmp_path):
    from zipfile import ZipFile
    import pytest
    from golos_mas.quality import QualityAgent
    path=tmp_path/'broken.pptx'
    with ZipFile(path,'w') as z:
        z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="ppt/missing.xml" Type="test"/></Relationships>')
    with pytest.raises(ValueError,match='отсутствует связанная часть'):
        QualityAgent()({'output':str(path),'slide_count':0,'topics':[],'warnings':[]})

def test_quality_rejects_embedded_objects(tmp_path):
    from zipfile import ZipFile
    import pytest
    from golos_mas.quality import QualityAgent
    path=tmp_path/'embedded.pptx'
    with ZipFile(path,'w') as z:
        z.writestr('ppt/embeddings/hidden.bin',b'test')
    with pytest.raises(ValueError,match='запрещённые встроенные части'):
        QualityAgent()({'output':str(path),'slide_count':0,'topics':[],'warnings':[]})

def test_coverage_calculates_partial_sum_and_flags_possible_duplicates():
    from golos_mas.agents import CoverageAgent
    records=[
        {'sheet':'SW','row':2,'date':'01.10.2026','source':'Чат','text':'Ошибка сети','signal':3},
        {'sheet':'прочее','row':3,'date':'01.10.2026','source':'ЧАТ','text':'Ошибка  сети','signal':3},
        {'sheet':'SW','row':4,'date':'01.10.2026','source':'Чат B','text':'Другая','signal':None},
    ]
    result=CoverageAgent()({'signals':records,'topics':[],'warnings':[]})
    c=result['coverage']
    assert c['signal_sum']==6 and c['numeric_signal_rows']==2 and c['invalid_signal_rows']==1
    assert c['distinct_named_sources']==2 and c['possible_duplicate_groups']==1
    assert len(c['duplicate_examples'][0])==2 and any('не удалены' in w for w in result['warnings'])


def test_invalid_signal_values_are_not_summed():
    from golos_mas.agents import AnalyticsAgent
    topic={'product':'P','topic':'T','score':9,'signals_score':None,'duration':None,'impact':None,'sheet':'P','row':2,'count':None}
    details=[{'product':'P','topic':'T','signal':x,'sheet':'SW','row':i} for i,x in enumerate([3,-1,1.5,True,None],2)]
    state=AnalyticsAgent()({'topics':[topic],'signals':details,'previous_topics':[],'warnings':[]})
    assert state['topics'][0]['signal_sum']==3 and any('корректного' in w for w in state['warnings'])

def test_summary_model_selects_only_rows_and_numbers_are_derived_locally():
    from types import SimpleNamespace
    from golos_mas.agents import NarrativeAgent
    class Model:
        def invoke(self,messages): return SimpleNamespace(content='{"номера":[1],"текст":"Ложная цифра 999999"}')
    current={'owner':'A','product':'Продукт','topic':'Проблема','score':10,'zone':'высокое внимание','count':70,'sheet':'Текущий','row':5}
    previous={'owner':'A','product':'Продукт','topic':'Проблема','count':35,'sheet':'Прошлый','row':3}
    result=NarrativeAgent(Model())({'topics':[current],'previous_topics':[previous],'warnings':[]})
    assert 'сигналов в приоритетах: 70' in result['narrative'] and 'изменение: +100%' in result['narrative']
    assert '999999' not in result['narrative']
    assert result['narrative_sources']==[{'sheet':'Текущий','row':5,'previous_sheet':'Прошлый','previous_row':3}]


def test_summary_never_infers_change_from_ambiguous_or_zero_baseline():
    from types import SimpleNamespace
    from golos_mas.agents import NarrativeAgent
    class Model:
        def invoke(self,messages): return SimpleNamespace(content='{"номера":[1]}')
    current={'owner':'A','product':'P','topic':'T','score':9,'zone':'высокое внимание','count':10,'sheet':'P','row':2}
    old={'owner':'A','product':'P','topic':'T','count':0,'sheet':'old','row':2}
    result=NarrativeAgent(Model())({'topics':[current],'previous_topics':[old],'warnings':[]})
    assert 'изменение:' not in result['narrative'] and 'previous_row' not in result['narrative_sources'][0]
    old['count']=5
    result=NarrativeAgent(Model())({'topics':[current],'previous_topics':[old,dict(old,row=3)],'warnings':[]})
    assert 'изменение:' not in result['narrative']


def test_summary_invalid_model_selection_is_rejected():
    from types import SimpleNamespace
    from golos_mas.agents import NarrativeAgent
    class Model:
        def invoke(self,messages): return SimpleNamespace(content='{"номера":[10]}')
    t={'product':'P','topic':'T','score':9,'zone':'высокое внимание','count':2,'sheet':'P','row':2}
    result=NarrativeAgent(Model())({'topics':[t],'previous_topics':[],'warnings':[]})
    assert result['narrative']=='' and not result['narrative_sources'] and result['warnings']

def test_summary_change_rounds_half_up_and_preserves_sign():
    from types import SimpleNamespace
    from golos_mas.agents import NarrativeAgent
    class Model:
        def invoke(self,messages): return SimpleNamespace(content='{"номера":[1]}')
    now={'owner':'A','product':'P','topic':'T','score':9,'zone':'высокое внимание','count':63,'sheet':'P','row':2}
    old={'owner':'A','product':'P','topic':'T','count':200,'sheet':'Old','row':2}
    result=NarrativeAgent(Model())({'topics':[now],'previous_topics':[old],'warnings':[]})
    assert 'изменение: −69%' in result['narrative']

def test_fallback_renders_the_end_of_long_topic_and_quote(tmp_path):
    from zipfile import ZipFile
    w=Workbook();p=w.active;p.title='P';p.append(['владелец','продукт','тема','общий балл'])
    topic=('Длинная тема ' * 80)+'КОНЕЦТЕМЫ';quote=('Подробное обращение ' * 100)+'КОНЕЦЦИТАТЫ'
    p.append(['Команда','Продукт',topic,9]);d=w.create_sheet('SW')
    d.append(['источник','продукт','тема','Маскированный текст из источника','сигнал']);d.append(['Чат','Продукт',topic,quote,1])
    book=tmp_path/'i.xlsx';w.save(book);out=tmp_path/'o.pptx'
    state=create_graph().invoke({'workbook':str(book),'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        data=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        assert 'КОНЕЦТЕМЫ'.encode() in data
        assert state['topics'][0]['quotes'][0]['text']==quote
        assert 'КОНЕЦЦИТАТЫ'.encode() not in data
    assert state['slide_count']==5


def test_status_metadata_requires_template_and_keeps_existing_output(tmp_path):
    import json,pytest
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','тема','общий балл']);s.append(['P','T',5]);book=tmp_path/'i.xlsx';w.save(book)
    statuses=tmp_path/'status.json';statuses.write_text(json.dumps([{'product':'P','topic':'T','mitigation':'done'}]),encoding='utf8')
    output=tmp_path/'o.pptx';output.write_bytes(b'existing')
    with pytest.raises(ValueError,match='укажите --template'):
        create_graph().invoke({'workbook':str(book),'current_sheet':'P','period':'месяц','statuses_json':str(statuses),'output':str(output)})
    assert output.read_bytes()==b'existing'
