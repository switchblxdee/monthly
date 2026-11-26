import os
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook

from golos_mas.agents import ExcelAgent, customer_quote
from golos_mas.graph import create_graph
from golos_mas.product_metrics import product_metrics, format_metric
from golos_mas.template_render import recurrence_code, order_topics_by_product
from golos_mas.source_inventory import product_source_rows


@pytest.mark.parametrize('duration,comment,expected',[
    (1,None,'new'),(1,'возобновилось','resumed'),(1,'возобновилась','resumed'),
    (1,'другая заметка',None),(2,'возобновилась','two'),(3,None,'long'),
    (4,None,'long'),(5,None,'long'),(None,'возобновилось',None),(0,None,None)])
def test_excel_duration_controls_icon(duration,comment,expected):
    assert recurrence_code({'duration':duration,'comment':comment,'status_record':{'recurrence':'new'}})==expected


def test_metrics_repeated_values_signed_delta_and_conflict():
    topics=[{'sheet':'P','row':2,'csat':4.46,'csat_delta':-0.03,'opinion_score':'4,3','opinion_delta':1},
            {'sheet':'P','row':3,'csat':'4,46','csat_delta':-0.03,'opinion_score':4.3,'opinion_delta':1}]
    values=product_metrics(topics)
    assert format_metric(values['csat'])=='4,46'
    assert format_metric(values['csat_delta'],True)=='-0,03'
    assert format_metric(values['opinion_delta'],True)=='+1'
    assert len(values['csat']['references'])==2
    topics[1]['csat']=3.5
    assert product_metrics(topics)['csat']['conflict'] and format_metric(product_metrics(topics)['csat'])=='3,5…4,46'


def test_month_inferred_from_priority_tabs_and_detail_dates(tmp_path):
    w=Workbook();current=w.active;current.title='приоритет ноябрь';current.append(['продукт','тема']);current.append(['P','T'])
    old=w.create_sheet('приоритет октябрь');old.append(['продукт','тема']);old.append(['P','T'])
    d=w.create_sheet('SW');d.append(['дата','источник','продукт','тема']);d.append(['01.11.2026','Чат','P','T']);d.append(['01.10.2026','Чат','P','T'])
    path=tmp_path/'book.xlsx';w.save(path)
    result=ExcelAgent()({'workbook':str(path)})
    assert result['current_sheet']=='приоритет ноябрь' and result['previous_sheet']=='приоритет октябрь'
    assert result['report_month']=='2026-11' and result['period']=='ноябрь 2026'
    assert len(result['signals'])==1 and result['date_filter']['other_month_rows']==1


def test_without_dates_or_year_month_is_not_guessed(tmp_path):
    w=Workbook();w.active.title='приоритет ноябрь';w.active.append(['продукт','тема']);w.active.append(['P','T'])
    path=tmp_path/'book.xlsx';w.save(path)
    with pytest.raises(ValueError,match='укажите --report-month'):
        ExcelAgent()({'workbook':str(path)})


def test_customer_words_are_not_rewritten_and_image_metadata_is_removed():
    assert customer_quote('185 - 14:00:07 - 151245702 - 151245702: Пользователь 1 Коллеги, не работает!')=='Коллеги, не работает!'
    assert customer_quote('510 - 11:06:45 - 708139329 - [АНАЛИЗ ИЗОБРАЖЕНИЯ: image.png] - ОС: Windows')==''


def test_product_order_uses_signal_sum_then_high_count():
    rows=[{'product':'B','count':5,'zone':'потенциальный риск'}, {'product':'A','count':5,'zone':'высокое внимание'},
          {'product':'C','count':6,'zone':'потенциальный риск'}]
    assert [x['product'] for x in order_topics_by_product(rows)]==['C','A','B']


def test_product_sources_include_other_topics_without_assigning_their_quotes():
    topics=[{'product':'P','matches':[{'sheet':'SW','row':2,'source':'Чат A'}]}]
    signals=[{'sheet':'SW','row':2,'product':'P','source':'Чат A'},
             {'sheet':'SW','row':3,'product':' P ','source':'Чат B'},
             {'sheet':'SW','row':4,'product':'Другое','source':'Чат C'}]
    refs=product_source_rows(topics,signals)
    assert [(x['source'],x['basis']) for x in refs]==[('Чат A','topic_match'),('Чат B','product')]
    assert len(product_source_rows(topics,signals,False))==1


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'),reason='Set GOLOS_TEST_TEMPLATE')
def test_new_columns_and_recurrence_are_shown_in_generated_deck(tmp_path):
    w=Workbook();p=w.active;p.title='приоритет ноябрь'
    p.append(['владелец','продукт','тема','кол-во сигналов','общий балл','длительность','Комментарий','CSAT','дельта CSAT','ПроМнение','дельта ПроМнения'])
    p.append(['Команда','P','Ошибка A',12,9,1,'возобновилось','4,46','−0,03','4,3','+1'])
    p.append(['Команда','P','Ошибка B',10,10,3,None,'=4.46',None,'—',None])
    d=w.create_sheet('SW');d.append(['дата','источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    d.append(['01.11.2026','Чат A','P','Ошибка A','185 - 14:00:07 - 151245702 - Пользователь 1 Клиент не может войти',0,1])
    d.append(['02.11.2026','Чат B','P','Ошибка B','Вторая ошибка',0,1])
    book=tmp_path/'book.xlsx';w.save(book);out=tmp_path/'out.pptx'
    result=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],'output':str(out)})
    assert result['report_month']=='2026-11'
    assert [recurrence_code(t) for t in result['topics']]==['resumed','long']
    with ZipFile(out) as deck:
        data=' '.join(deck.read(n).decode() for n in deck.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        detail = next(ET.fromstring(deck.read(n)) for n in deck.namelist()
                      if n.startswith('ppt/slides/slide') and n.endswith('.xml')
                      and 'Клиент не может войти'.encode() in deck.read(n))
        exact_text = [(node.text or '').strip() for node in detail.iter() if node.tag.endswith('}t')]
        assert all(x in exact_text for x in ('CSAT', 'ПроМнение', '4,46', '-0,03', '4,3', '+1'))
        assert exact_text.count('+1') == exact_text.count('-0,03') == 1
        assert '151245702' not in data and '14:00:07' not in data
        assert 'Возобновилось обсуждение' in data and 'Повторяется 3 месяца или более' in data

def test_historical_duplicate_does_not_block_latest_month_and_ambiguous_year_does(tmp_path):
    w = Workbook()
    w.active.title = 'Приоритет ноября'
    for name in ('Приоритет сентября', 'Приоритет сентябрь', 'Приоритет октября'):
        w.create_sheet(name)
    for sheet in w:
        sheet.append(['продукт', 'тема'])
        sheet.append(['P', 'T'])
    detail = w.create_sheet('SW')
    detail.append(['дата', 'продукт', 'тема'])
    for value in ('01.09.2026', '01.11.2026'):
        detail.append([value, 'P', 'T'])
    path = tmp_path / 'book.xlsx'
    w.save(path)
    assert ExcelAgent()({'workbook': str(path)})['report_month'] == '2026-11'
    detail.append(['01.11.2025', 'P', 'T'])
    w.save(path)
    with pytest.raises(ValueError, match='за несколько лет'):
        ExcelAgent()({'workbook': str(path)})


def test_uncomputed_comment_is_not_assumed_empty():
    assert recurrence_code({'duration': 1, 'comment': None, 'uncached_formulas': ['comment']}) is None


def test_quote_keeps_times_and_reference_numbers_inside_customer_words():
    assert customer_quote('185 - 14:00:07 - 151245702: Пользователь 1 С 10:30 ошибка, обращение 123456789.') == 'С 10:30 ошибка, обращение 123456789.'


def test_large_source_list_fits_footer_with_audit_marker():
    from golos_mas.detail_widgets import source_widget_layout

    sources = [f'Источник {i} с длинным названием для проверки вёрстки' for i in range(40)]
    widgets, hidden = source_widget_layout(sources)
    assert 0 < hidden < len(sources)
    assert [name for name, *_ in widgets[:-1]] == sources[:len(sources) - hidden]
    assert widgets[-1][0] == f'Ещё {hidden} · в аудите'
    assert all(y + height <= 7.42 for _, _, y, height in widgets)
