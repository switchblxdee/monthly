import os
from pathlib import Path
from zipfile import ZipFile
import pytest
from openpyxl import Workbook
from golos_mas.graph import create_graph

@pytest.mark.parametrize('number,expected',[(0,'0 источников'),(1,'1 источник'),(2,'2 источника'),(5,'5 источников'),(11,'11 источников'),(21,'21 источник'),(24,'24 источника')])
def test_source_count_uses_correct_russian_form(number,expected):
    from golos_mas.template_render import source_count_label
    assert source_count_label(number)==expected

def test_long_product_title_is_reduced_without_losing_name():
    import xml.etree.ElementTree as ET
    from golos_mas.template_render import N, product_title, content
    shape=ET.fromstring('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Старое</a:t></a:r></a:p></p:txBody></p:sp>')
    name='Система учёта и управления ПУЗ и ТУЗ (инфраструктурный IDM)'
    product_title(shape,name)
    assert content(shape)==name
    assert 1400<=int(shape.find('.//a:rPr',N).get('sz'))<2800

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_long_product_name_is_reduced_in_overview_table(tmp_path):
    import xml.etree.ElementTree as ET
    from golos_mas.template_render import N, content
    name='Учебная система управления инфраструктурой и доступом пользователей в технологических средах организации'
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл']);s.append(['Команда',name,'Ошибка доступа',9])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'result.pptx'
    create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    with ZipFile(out) as deck:
        root=ET.fromstring(deck.read('ppt/slides/slide6.xml'))
        cell=next(c for c in root.findall('.//a:tc',N) if content(c)==name)
        assert 800<=int(cell.find('.//a:rPr',N).get('sz'))<1200

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_template_deck_contains_only_current_values(tmp_path):
    w=Workbook();ws=w.active;ws.title='приоритет октябрь';ws.append(['владелец','продукт','тема','кол-во сигналов','общий балл','длительность'])
    ws.append(['Учебная платформа','Новый продукт','Тестовая тема',2,9,2])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','лайки','сигнал']);d.append(['Тестовый источник','Новый продукт','Тестовая тема','Тестовый отзыв',1,2])
    path=tmp_path/'test.xlsx';w.save(path);out=tmp_path/'test.pptx'
    result=create_graph().invoke({'workbook':str(path),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':ws.title,'period':'декабрь 2030','output':str(out)})
    with ZipFile(out) as deck:
        data=b''.join(deck.read(name) for name in deck.namelist() if name.endswith('.xml'))
        assert 'Тестовый отзыв'.encode() in data
        assert 'SberESM'.encode() not in data and 'октябрь 2026'.encode() not in data
        from golos_mas.template_render import forbidden_embedded_parts
        assert not forbidden_embedded_parts(deck.namelist())
        assert b'oleObj' not in deck.read('ppt/slides/slide2.xml')
    with ZipFile(os.environ['GOLOS_TEST_TEMPLATE']) as original:
        assert b'oleObj' in original.read('ppt/slides/slide2.xml')
    assert result['slide_count']==7
    with ZipFile(out) as deck:
        toc=deck.read('ppt/slides/slide3.xml').decode()
        assert 'Слайды 5–7' in toc and 'Новый продукт' in toc

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_multi_page_contents_matches_generated_sections(tmp_path):
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    w=Workbook(); ws=w.active;ws.title='priorities';ws.append(['владелец','продукт','тема','общий балл'])
    expected=[]
    for i in range(8):
        for j in range(3):
            product=f'Продукт {i}-{j}';expected.append(product);ws.append([f'Группа {i}',product,f'Тема {i}-{j}',5])
    path=tmp_path/'input.xlsx';w.save(path);out=tmp_path/'result.pptx'
    result=create_graph().invoke({'workbook':str(path),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'priorities','period':'декабрь 2030','output':str(out)})
    with ZipFile(out) as deck:
        slides=[]
        for i in range(1,result['slide_count']+1):
            root=ET.fromstring(deck.read(f'ppt/slides/slide{i}.xml'))
            slides.append(' '.join(el.text or '' for el in root.findall('.//a:t',ns)))
        contents=[s for s in slides if 'Продуктов с темами высокого внимания' in s]
        assert len(contents)==1
        for product in expected: assert any(product in page for page in contents)
        for i in range(8):
            start=next(pos for pos,text in enumerate(slides,1) if f'Продукты и сервисы Группа {i}' in text)
            end=(next((pos for pos,text in enumerate(slides,1) if f'Продукты и сервисы Группа {i+1}' in text),len(slides)+1)-1)
            assert any(f'Слайды {start}–{end}' in text and f'Группа {i}' in text for text in contents)

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_repeated_contents_cards_keep_group_number(tmp_path):
    import xml.etree.ElementTree as ET
    from golos_mas.template_render import N, content
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл'])
    for i in range(5): s.append(['Команда',f'Продукт {i}: система доступа и управления инфраструктурой организации','Ошибка',9])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'out.pptx'
    create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    with ZipFile(out) as deck:
        root=ET.fromstring(deck.read('ppt/slides/slide3.xml'))
        texts=[content(shape) for shape in root.findall('.//p:sp',N)]
        assert sum(text.startswith('01 Команда') for text in texts)==1
        assert not any(text.startswith('02 Команда') for text in texts)

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_product_name_normalization_is_consistent_in_summary_and_contents(tmp_path):
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл'])
    s.append(['Команда','Продукт Ё','Ошибка 1',9]);s.append(['Команда','  продукт   Е  ','Ошибка 2',9])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'out.pptx'
    create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    with ZipFile(out) as deck:
        summary=deck.read('ppt/slides/slide2.xml').decode();toc=deck.read('ppt/slides/slide3.xml').decode()
        assert 'Продуктов: 1' in summary and 'Продуктов с темами высокого внимания: 1' in toc
        assert 'Продукт Ё' in toc and 'продукт   Е' not in toc

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_high_attention_product_count_is_unique_across_owners(tmp_path):
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл'])
    s.append(['Команда A','Портал ДИ','Проблема 1',9]);s.append(['Команда B','Портал ДИ','Проблема 2',9])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'out.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    with ZipFile(out) as deck:
        summary=deck.read('ppt/slides/slide2.xml').decode();toc=deck.read('ppt/slides/slide3.xml').decode()
        assert 'Продуктов: 1' in summary and 'Продуктов с темами высокого внимания: 1' in toc
        assert 'Команда A' in toc and 'Команда B' in toc
    assert any('нескольких владельцев' in warning for warning in state['warnings'])

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_long_text_is_split_without_omitting_segments(tmp_path):
    from golos_mas.template_render import split_text
    w=Workbook();ws=w.active;ws.title='priorities';ws.append(['владелец','продукт','тема','общий балл'])
    topic=' '.join(f'длинноеописание{i}' for i in range(70))
    quote=' '.join(f'текстобращения{i}' for i in range(75))
    ws.append(['Платформа','Продукт',topic,9]);s=w.create_sheet('SW')
    s.append(['источник','продукт','тема','Маскированный текст из источника','сигнал'])
    s.append(['Тестовый источник','Продукт',topic,quote,1]);path=tmp_path/'long.xlsx';w.save(path)
    out=tmp_path/'long.pptx'
    result=create_graph().invoke({'workbook':str(path),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'priorities','period':'октябрь','output':str(out)})
    with ZipFile(out) as deck:
        slides=b''.join(deck.read(name) for name in deck.namelist() if name.startswith('ppt/slides/slide') and name.endswith('.xml')).decode()
        assert result['topics'][0]['topic']==topic and result['topics'][0]['quotes'][0]['text']==quote
        assert result['topics'][0]['presentation_quotes'][0]['text'] in slides
        assert result['slide_count'] < 15
        assert 'прод.' in slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_unbroken_long_technical_identifier_is_preserved(tmp_path):
    from golos_mas.template_render import split_text
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл']);s.append(['Команда','Продукт','Ошибка',9])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','сигнал'])
    quote='https://example.invalid/'+('ABCDEF0123456789'*40)
    d.append(['Чат','Продукт','Ошибка',quote,1]);book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'result.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    assert state['topics'][0]['quotes'][0]['text']==quote
    with ZipFile(out) as deck:
        slides=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert state['topics'][0]['presentation_quotes'][0]['text'] in slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_model_choice_is_rendered_from_original_record(tmp_path):
    from types import SimpleNamespace
    class Stub:
        def invoke(self,messages):
            if 'Выбери до трёх' in messages[0].content: return SimpleNamespace(content='{"номера":[2]}')
            return SimpleNamespace(content='Данные по выбранной теме.')
    w=Workbook();ws=w.active;ws.title='P';ws.append(['владелец','продукт','тема','кол-во сигналов','общий балл']);ws.append(['Команда','Продукт','Ошибка',2,9])
    s=w.create_sheet('SW');s.append(['источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    s.append(['Ист1','Продукт','Ошибка','Текст первый',10,1]);s.append(['Ист2','Продукт','Ошибка','Текст второй',1,1])
    path=tmp_path/'input.xlsx';w.save(path);out=tmp_path/'out.pptx'
    state=create_graph(Stub(),use_model_quotes=True,agentic=False).invoke({'workbook':str(path),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'октябрь','output':str(out)})
    assert state['topics'][0]['quotes'][0]['row']==3
    with ZipFile(out) as deck:
        data=b''.join(deck.read(name) for name in deck.namelist() if name.startswith('ppt/slides/slide') and name.endswith('.xml'))
        assert 'Текст второй'.encode() in data and 'Текст первый'.encode() not in data

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_status_metadata_reaches_summary_and_detail(tmp_path):
    import json
    w=Workbook();ws=w.active;ws.title='P';ws.append(['владелец','продукт','тема','общий балл']);ws.append(['Владелец','Продукт','Тема',9])
    s=w.create_sheet('SW');s.append(['источник','продукт','тема','Маскированный текст из источника','сигнал']);s.append(['Источник','Продукт','Тема','Цитата',1])
    book=tmp_path/'i.xlsx';w.save(book)
    status=tmp_path/'s.json';status.write_text(json.dumps([{'product':'Продукт','topic':'Тема','controls':['backlog'],'mitigation':'release','version':'7.0','deadline':'4Q 2026'}]),encoding='utf8')
    out=tmp_path/'o.pptx';create_graph().invoke({'workbook':str(book),'statuses_json':str(status),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'октябрь','output':str(out)})
    with ZipFile(out) as deck:
        summary=deck.read('ppt/slides/slide2.xml').decode()
        all_slides=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert 'Релиз 7.0 4Q 2026' in summary and 'Бэклог команды' in all_slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_promises_generate_calculated_pages_without_old_chart_data(tmp_path):
    import json
    w=Workbook();ws=w.active;ws.title='P';ws.append(['владелец','продукт','тема','общий балл']);ws.append(['Команда','Продукт','Тема',5]);book=tmp_path/'i.xlsx';w.save(book)
    promises=[{'id':'P1','owner':'Команда','description':'Задача 1','source':'Встреча','initial_due':'2026-10-01','current_due':'2026-11-01','outcome':'postponed','in_work_start':True,'in_work_current':True,'comment':'Новый срок'},
              {'id':'P2','owner':'Команда','description':'Задача 2','source':'Встреча','initial_due':'2026-10-02','outcome':'completed_on_time','in_work_start':True,'in_work_current':False}]
    data=tmp_path/'promises.json';data.write_text(json.dumps(promises),encoding='utf8');out=tmp_path/'out.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'октябрь 2026','report_month':'2026-10','promises_json':str(data),'output':str(out)})
    assert state['promise_report']['due_count']==2 and state['promise_report']['completed_count']==1
    with ZipFile(out) as deck:
        xml=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert '50%' in xml and 'Новый срок' in xml and 'P1' in xml
        assert '69%' not in xml and not any(n.startswith('ppt/charts/') for n in deck.namelist())

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_ux_has_separate_five_column_table(tmp_path):
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    w=Workbook();ws=w.active;ws.title='P';ws.append(['владелец','продукт','тема','общий балл']);ws.append(['Команда','Продукт','Обычная тема',5]);ws.append(['Команда','Продукт','Проблема UX','UX'])
    book=tmp_path/'i.xlsx';w.save(book);out=tmp_path/'o.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        slides=[ET.fromstring(deck.read(f'ppt/slides/slide{i}.xml')) for i in range(1,state['slide_count']+1)]
        texts=[' '.join((t.text or '') for t in s.findall('.//a:t',ns)) for s in slides]
        ux=next(i for i,t in enumerate(texts) if 'UX-проблемы' in t)
        assert 'Проблема UX' in texts[ux] and 'Обычная тема' not in texts[ux]
        assert any('Темы, требующие внимания' in t and 'Обычная тема' in t and 'Проблема UX' not in t for t in texts)
        header=slides[ux].find('.//a:tbl/a:tr',ns)
        assert len(header.findall('a:tc',ns))==5

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_model_topic_matching_drives_presentation_with_source_trace(tmp_path):
    from types import SimpleNamespace
    class Model:
        def invoke(self,messages):
            if 'Выбери только варианты' in messages[0].content: return SimpleNamespace(content='{"номера":[1]}')
            return SimpleNamespace(content='Данные по инциденту.')
    w=Workbook();ws=w.active;ws.title='P';ws.append(['владелец','продукт','тема','общий балл']);ws.append(['Прочее','ЦОД Балаково','Недоступность инфраструктуры Балаково',9])
    s=w.create_sheet('прочее');s.append(['источник','продукт','тема','Маскированный текст из источника','сигнал']);s.append(['Чат','DropApp','Сбой инфраструктуры Балаково','Исходный комментарий',1])
    book=tmp_path/'book.xlsx';w.save(book);out=tmp_path/'out.pptx'
    state=create_graph(Model(),use_model_matching=True,agentic=False).invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'октябрь','output':str(out)})
    assert state['topics'][0]['match_method']=='gigachat' and state['topics'][0]['quotes'][0]['row']==2
    with ZipFile(out) as deck:
        alltext=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        assert 'Исходный комментарий'.encode() in alltext

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_model_matching_quote_selection_and_summary_run_together(tmp_path):
    from types import SimpleNamespace
    class Model:
        def __init__(self): self.calls=[]
        def invoke(self,messages):
            prompt=messages[0].content;self.calls.append(prompt)
            if 'Выбери до трёх номеров обращений' in prompt:
                return SimpleNamespace(content='{"номера":[2],"текст":"ВЫДУМАННОЕ ОБРАЩЕНИЕ"}')
            return SimpleNamespace(content='{"номера":[1],"текст":"ВЫДУМАННЫЙ ФАКТ"}')
    w=Workbook();s=w.active;s.title='P'
    s.append(['владелец','продукт','тема','кол-во сигналов','общий балл'])
    s.append(['Команда','ЦОД Балаково','Недоступность инфраструктуры Балаково',12,9])
    old=w.create_sheet('Old');old.append(['владелец','продукт','тема','кол-во сигналов'])
    old.append(['Команда','ЦОД Балаково','Недоступность инфраструктуры Балаково',6])
    d=w.create_sheet('прочее');d.append(['дата','источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    d.append(['01.11.2026','Чат','DropApp','Сбой инфраструктуры Балаково','Первый отзыв',9,10])
    d.append(['02.11.2026','Чат','DropApp','Сбой инфраструктуры Балаково','Второй отзыв',1,2])
    d.append(['02.10.2026','Чат','DropApp','Сбой инфраструктуры Балаково','Прошлый месяц',0,1])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'result.pptx';model=Model()
    state=create_graph(model,use_model_matching=True,use_model_quotes=True,agentic=False).invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','previous_sheet':'Old','period':'ноябрь 2026','report_month':'2026-11','output':str(out)})
    t=state['topics'][0]
    assert t['match_method']=='gigachat' and t['quote_selection']=='gigachat'
    assert t['quotes'][0]['row']==3 and t['signal_sum']==12
    assert state['narrative_sources']==[{'sheet':'P','row':2,'previous_sheet':'Old','previous_row':2}]
    assert len(model.calls)==3 and state['date_filter']['other_month_rows']==1
    with ZipFile(out) as deck:
        slides=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert 'Второй отзыв' in slides and 'Первый отзыв' not in slides and 'Прошлый месяц' not in slides
        assert 'ВЫДУМАНН' not in slides and 'изменение: +100%' in slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_model_outage_preserves_exact_evidence_and_leaves_unverified_matches_out(tmp_path):
    class Unavailable:
        def invoke(self,messages): raise TimeoutError('private endpoint detail')
    w=Workbook();s=w.active;s.title='P'
    s.append(['владелец','продукт','тема','кол-во сигналов','общий балл'])
    s.append(['Команда','Продукт A','Ошибка авторизации',5,9])
    s.append(['Команда','ЦОД Балаково','Недоступность инфраструктуры Балаково',1,9])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    d.append(['Чат A','Продукт A','Ошибка авторизации','Подтверждённое обращение',4,5])
    d.append(['Чат B','DropApp','Сбой инфраструктуры Балаково','Неподтверждённое обращение',0,1])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'result.pptx'
    state=create_graph(Unavailable(),use_model_matching=True,use_model_quotes=True,agentic=False).invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    exact,unmatched=state['topics']
    assert exact['match_method']=='exact' and exact['quote_selection']=='likes_fallback'
    assert not unmatched['matches'] and not unmatched['quotes']
    assert any('сопоставление GigaChat не удалось' in warning for warning in state['warnings'])
    assert any('отбор GigaChat не удался' in warning for warning in state['warnings'])
    assert all('private endpoint detail' not in warning for warning in state['warnings'])
    with ZipFile(out) as deck:
        slides=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert 'Подтверждённое обращение' in slides and 'Неподтверждённое обращение' not in slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_invalid_priority_signal_count_is_not_presented_as_a_number(tmp_path):
    w=Workbook();s=w.active;s.title='P'
    s.append(['владелец','продукт','тема','кол-во сигналов','общий балл'])
    s.append(['Команда','Продукт','Ошибка',-42,9])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','сигнал'])
    d.append(['Чат','Продукт','Ошибка','Комментарий',1])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'result.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    assert state['topics'][0]['count']==-42
    assert any('некорректное количество сигналов' in warning for warning in state['warnings'])
    with ZipFile(out) as deck:
        slides=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert 'сигналов: нет данных' in slides and '-42 сигналов' not in slides and 'сигналов: -42' not in slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
@pytest.mark.parametrize('mau',[-100,True,0.5])
def test_invalid_mau_is_not_displayed_as_an_audience_size(tmp_path,mau):
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','MAU','общий балл']);s.append(['Команда','Продукт','Ошибка',mau,9])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'result.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь','output':str(out)})
    assert any('некорректное значение MAU' in warning for warning in state['warnings'])
    with ZipFile(out) as deck:
        slides=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert 'MAU: —' in slides and f'{mau} MAU' not in slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_quote_limit_reduces_slides_and_preserves_omitted_references(tmp_path):
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл']);s.append(['Команда','Продукт','Ошибка',9])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    for i in range(3): d.append([f'Канал {i}','Продукт','Ошибка',f'Обращение {i} '+('длинный комментарий '*18),3-i,1])
    book=tmp_path/'input.xlsx';w.save(book)
    common={'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'ноябрь'}
    full=create_graph().invoke(dict(common,output=str(tmp_path/'full.pptx')))
    limited=create_graph().invoke(dict(common,output=str(tmp_path/'limited.pptx'),max_quotes_per_topic=1))
    assert limited['slide_count']==full['slide_count']
    assert limited['topics'][0]['omitted_quote_references']==[{'sheet':'SW','row':3},{'sheet':'SW','row':4}]
    assert limited['coverage']['presentation_quote_rows']==1
    assert any('по заданному ограничению' in warning for warning in limited['warnings'])
    with ZipFile(tmp_path/'limited.pptx') as deck:
        slides=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')).decode()
        assert 'Обращение 0' in slides and 'Обращение 1' not in slides and 'Обращение 2' not in slides

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_explicit_new_topic_uses_original_legend_icon(tmp_path):
    import json
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    w=Workbook();ws=w.active;ws.title='P';ws.append(['владелец','продукт','тема','общий балл','длительность']);ws.append(['Команда','Продукт','Тема',9,1])
    book=tmp_path/'i.xlsx';w.save(book)
    status=tmp_path/'s.json';status.write_text(json.dumps([{'product':'Продукт','topic':'Тема','recurrence':'new'}]),encoding='utf8')
    out=tmp_path/'o.pptx';state=create_graph().invoke({'workbook':str(book),'statuses_json':str(status),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'октябрь','output':str(out)})
    with ZipFile(out) as deck:
        slide=ET.fromstring(deck.read('ppt/slides/slide6.xml'))
        blips=[x.get('{'+ns['r']+'}embed') for x in slide.findall('.//a:blip',ns)]
        assert blips.count('rId7')>=2
        allxml=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        assert 'Новая тема'.encode() in allxml and state['topics'][0]['status_record']['recurrence']=='new'

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_contents_uses_calculated_zone_color(tmp_path):
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл'])
    s.append(['Команда','Продукт','Риск',5]);s.append(['Команда','Продукт','Высокий',9]);s.append(['Команда','UX продукт','Интерфейс','UX'])
    book=tmp_path/'i.xlsx';w.save(book);out=tmp_path/'o.pptx'
    create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        slide=ET.fromstring(deck.read('ppt/slides/slide3.xml'))
        texts=' '.join(el.text or '' for el in slide.findall('.//a:t',ns))
        colors=[el.get('val') for el in slide.findall('.//a:rPr/a:solidFill/a:schemeClr',ns)]
        assert 'Продуктов с темами высокого внимания: 1' in texts
        assert 'accent6' in colors and 'UX / зона не указана' not in texts
        assert 'Мониторинг обещаний' in texts and 'Нет данных' in texts
        assert 'темы в зоне «высокое внимание»' in texts

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_contents_runs_precede_paragraph_end_properties(tmp_path):
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл']);s.append(['Команда','Продукт','Тема',9]);path=tmp_path/'i.xlsx';w.save(path);out=tmp_path/'o.pptx'
    create_graph().invoke({'workbook':str(path),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        slide=ET.fromstring(deck.read('ppt/slides/slide3.xml'))
        for para in slide.findall('.//a:p',ns):
            tags=[e.tag.split('}')[-1] for e in para]
            if 'endParaRPr' in tags: assert tags[-1]=='endParaRPr'

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_summary_qualifies_partial_detail_sum(tmp_path):
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл']);s.append(['Команда','Продукт','Тема',5])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','сигнал']);d.append(['Чат','Продукт','Тема','Текст 1',3]);d.append(['Чат','Продукт','Тема','Текст 2',None])
    book=tmp_path/'i.xlsx';w.save(book);out=tmp_path/'o.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    assert state['coverage']['numeric_signal_rows']==1 and state['coverage']['invalid_signal_rows']==1
    with ZipFile(out) as deck:
        summary=deck.read('ppt/slides/slide2.xml').decode()
        assert 'Детализация (SW): строк: 2' in summary and 'сумма сигналов (неполная): 3 (строк с числом: 1)' in summary

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_detail_uses_original_icons_and_accessible_descriptions(tmp_path):
    import json
    import xml.etree.ElementTree as ET
    ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл','длительность']);s.append(['Команда','Продукт','Тема',9,1])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','сигнал']);d.append(['Чат','Продукт','Тема','Комментарий',1])
    book=tmp_path/'i.xlsx';w.save(book);status=tmp_path/'s.json'
    status.write_text(json.dumps([{'product':'Продукт','topic':'Тема','recurrence':'new','controls':['aibolit','backlog'],'mitigation':'release','version':'7.0','deadline':'4Q 2026','promnenie':True,'after_close':True}]),encoding='utf8')
    out=tmp_path/'o.pptx';result=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'statuses_json':str(status),'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        detail=ET.fromstring(deck.read(f"ppt/slides/slide{result['slide_count']}.xml"))
        descriptors={e.get('descr') for e in detail.findall('.//p:pic/p:nvPicPr/p:cNvPr',ns)}
        assert {'Новая тема','АйболИТ+','Бэклог команды','Релиз','Повтор после закрытия','Упоминание в ПроМнении'} <= descriptors
        overview=ET.fromstring(deck.read('ppt/slides/slide6.xml'))
        overview_descriptors={e.get('descr') for e in overview.findall('.//p:pic/p:nvPicPr/p:cNvPr',ns)}
        assert {'Новая тема','АйболИТ+','Бэклог команды','Релиз','Повтор после закрытия','Упоминание в ПроМнении'} <= overview_descriptors
        embeds=[b.get('{'+ns['r']+'}embed') for b in detail.findall('.//a:blip',ns)]
        assert sum(e.startswith('rIdDetailIcon') for e in embeds if e)>=6

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_missing_required_legend_icon_fails_instead_of_disappearing():
    import xml.etree.ElementTree as ET
    from golos_mas.template_render import table_slide, collect_legend, shapes, N
    with ZipFile(os.environ['GOLOS_TEST_TEMPLATE']) as deck:
        root=ET.fromstring(deck.read('ppt/slides/slide6.xml'))
    tree=shapes(root);tree.remove(collect_legend(root)['aibolit'])
    item={'product':'P','topic':'T','zone':'потенциальный риск','duration':1,'extra':{},'status_record':{'controls':['aibolit']}}
    with pytest.raises(ValueError,match='не найден необходимый значок'):
        table_slide(root,'Команда',[item],1,1)

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_grounded_summary_on_slide_has_traceable_previous_period(tmp_path):
    from types import SimpleNamespace
    class Model:
        def invoke(self,messages): return SimpleNamespace(content='{"номера":[1],"текст":"выдуманное значение 999999"}')
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','кол-во сигналов','общий балл']);s.append(['Команда','Продукт','Проблема',70,9])
    old=w.create_sheet('Old');old.append(['владелец','продукт','тема','кол-во сигналов']);old.append(['Команда','Продукт','Проблема',35])
    path=tmp_path/'i.xlsx';w.save(path);out=tmp_path/'o.pptx'
    state=create_graph(Model(),agentic=False).invoke({'workbook':str(path),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','previous_sheet':'Old','period':'месяц','output':str(out)})
    assert state['narrative_sources']==[{'sheet':'P','row':2,'previous_sheet':'Old','previous_row':2}]
    with ZipFile(out) as deck:
        summary=deck.read('ppt/slides/slide2.xml').decode()
        assert 'изменение: +100%' in summary and '999999' not in summary

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_short_quotes_share_slide_and_all_selected_evidence_is_rendered(tmp_path):
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл']);s.append(['Команда','Продукт','Тема',9])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    d.append(['Канал один','Продукт','Тема','Первое обращение',3,4]);d.append(['Канал два','Продукт','Тема','Второе обращение',2,3]);d.append(['Канал три','Продукт','Тема','Третье обращение',1,2])
    book=tmp_path/'i.xlsx';w.save(book);out=tmp_path/'o.pptx'
    result=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        slides=[ET.fromstring(deck.read(f'ppt/slides/slide{i}.xml')) for i in range(1,result['slide_count']+1)]
        texts=[' '.join(t.text or '' for t in slide.findall('.//a:t',ns)) for slide in slides]
        first=next(i for i,text in enumerate(texts) if 'Первое обращение' in text)
        assert 'Второе обращение' in texts[first] and 'Канал один' in texts[first] and 'Канал два' in texts[first]
        assert not any('Третье обращение' in text for text in texts)
        assert result['topics'][0]['presentation_omitted_quote_references']==[{'sheet':'SW','row':4}]
        quote_runs=[r for r in slides[first].findall('.//a:r',ns) if 'Первое обращение' in ''.join(t.text or '' for t in r.findall('.//a:t',ns))]
        assert quote_runs[0].find('a:rPr',ns).get('b')=='0'

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_overview_groups_products_and_merges_adjacent_cells(tmp_path):
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл'])
    s.append(['Команда','A','Тема риска',5]);s.append(['Команда','B','Другая',5]);s.append(['Команда','A','Важная',9]);s.append(['Команда','A','Наблюдение',4])
    book=tmp_path/'i.xlsx';w.save(book);out=tmp_path/'o.pptx'
    create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        slide=ET.fromstring(deck.read('ppt/slides/slide6.xml'));rows=slide.findall('.//a:tbl/a:tr',ns)[1:]
        texts=[''.join(x.text or '' for x in row.findall('a:tc',ns)[1].findall('.//a:t',ns)) for row in rows]
        assert texts==['Важная','Тема риска','Наблюдение','Другая']
        cells=[row.findall('a:tc',ns)[0] for row in rows]
        assert cells[0].get('rowSpan')=='3' and cells[1].get('vMerge')=='1' and cells[2].get('vMerge')=='1'
        assert not cells[3].get('vMerge') and ''.join(x.text or '' for x in cells[3].findall('.//a:t',ns))=='B'


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_product_merge_restarts_at_new_slide(tmp_path):
    import xml.etree.ElementTree as ET
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл'])
    for i in range(8): s.append(['Команда','Продукт',f'Тема {i}',5])
    book=tmp_path/'i.xlsx';w.save(book);out=tmp_path/'o.pptx'
    create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    with ZipFile(out) as deck:
        first=ET.fromstring(deck.read('ppt/slides/slide6.xml'));second=ET.fromstring(deck.read('ppt/slides/slide7.xml'))
        c1=first.findall('.//a:tbl/a:tr',ns)[1].findall('a:tc',ns)[0]
        c2=second.findall('.//a:tbl/a:tr',ns)[1].findall('a:tc',ns)[0]
        assert c1.get('rowSpan')=='7' and c2.get('vMerge') is None
        assert ''.join(x.text or '' for x in c2.findall('.//a:t',ns))=='Продукт'

@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_batched_quote_selection_reaches_presentation(tmp_path):
    import json
    from types import SimpleNamespace
    class Model:
        def invoke(self,messages):
            payload=json.loads(messages[1].content)
            if isinstance(payload,dict) and 'кандидаты' in payload:
                return SimpleNamespace(content=json.dumps({'номера':[len(payload['кандидаты'])]}))
            return SimpleNamespace(content='{"номера":[1]}')
    w=Workbook();p=w.active;p.title='P';p.append(['владелец','продукт','тема','общий балл']);p.append(['Команда','Продукт','Тема',9])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    for i in range(1,26): d.append(['Чат','Продукт','Тема',f'Уникальный отзыв номер {i}',30-i,1])
    path=tmp_path/'book.xlsx';w.save(path);out=tmp_path/'out.pptx'
    state=create_graph(Model(),use_model_quotes=True,agentic=False).invoke({'workbook':str(path),'template':os.environ['GOLOS_TEST_TEMPLATE'],'current_sheet':'P','period':'месяц','output':str(out)})
    assert state['topics'][0]['quote_selection_stats']['initial_candidates_attempted']==25
    assert any(q['text']=='Уникальный отзыв номер 25' for q in state['topics'][0]['quotes'])
    with ZipFile(out) as deck:
        data=b''.join(deck.read(n) for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        for quote in state['topics'][0]['presentation_quotes']:
            assert quote['text'].encode() in data
        assert len(state['topics'][0]['presentation_quotes'])==2
