import os
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pytest
from openpyxl import Workbook

from golos_mas.graph import create_graph
from golos_mas.source_inventory import SourceInventory
from golos_mas.worker_agents import ReportWorker
from langchain_core.messages import AIMessage, ToolMessage

NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE')
def test_compact_details_preserve_sources_and_rank_high_topics_by_signals(tmp_path):
    w=Workbook(); p=w.active; p.title='приоритет октябрь'
    p.append(['владелец','продукт','тема','кол-во сигналов','общий балл'])
    for i in range(12): p.append(['Команда','Продукт',f'Проблема {i}',i+1,12 if i==0 else 9])
    p.append(['Команда','Продукт','Проблема UX',100,'UX'])
    d=w.create_sheet('SW')
    d.append(['дата','источник','продукт','тема','Маскированный текст из источника','лайки','сигнал'])
    for i in range(12):
        for j in range(3):
            d.append(['01.10.2026',f'Канал {j}','Продукт',f'Проблема {i}',f'Сообщение {j} '+('длинный текст ' * 400),3-j,1])
    d.append(['02.10.2026','UX чат','Продукт','Проблема UX','UX отзыв',0,1])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'result.pptx'
    state=create_graph().invoke({'workbook':str(book),'template':os.environ['GOLOS_TEST_TEMPLATE'],
        'current_sheet':'приоритет октябрь','report_month':'2026-10','period':'октябрь 2026','output':str(out)})
    assert state['slide_count'] < 30
    assert state['source_inventory']['detail_rows_in_period']==37
    sheet=state['source_inventory']['detail_sheets'][0]
    assert sheet['source_count']==4 and sheet['sources']['Канал 0']==12
    assert state['coverage']['presentation_quote_rows']==24
    with ZipFile(out) as deck:
        def text(i):
            return ' '.join(t.text or '' for t in ET.fromstring(deck.read(f'ppt/slides/slide{i}.xml')).findall('.//a:t',NS))
        summary=text(2)
        assert 'Проблема 11' in summary and 'Проблема 10' in summary and 'Проблема 9' in summary
        slides=[text(i) for i in range(1,state['slide_count']+1)]
        details=[s for s in slides if 'Темы с высоким приоритетом' in s]
        assert len(details)==12 and all('UX отзыв' not in s for s in details)
        assert all('Канал 0' in s and 'Канал 1' in s and 'Канал 2' in s for s in details)
        assert all('SW!' in s for s in details)
    assert len(state['topics'][0]['quotes'][0]['text']) > 1000
    assert len(state['topics'][0]['presentation_quotes'][0]['text']) <= 180


def test_data_agent_can_page_through_all_sources_and_preserve_inventory():
    class Model:
        def __init__(self): self.results=[]
        def bind_tools(self, tools):
            model=self
            class Bound:
                def __init__(self): self.calls=0
                def invoke(self, messages):
                    if isinstance(messages[-1], ToolMessage): model.results.append(messages[-1].content)
                    calls=[('inventory_sources',{}),('inspect_sheets',{}),('inspect_sources',{'sheet':'SW','offset':50}),('finish',{})]
                    name,args=calls[self.calls]; self.calls+=1
                    return AIMessage(content='',tool_calls=[{'name':name,'args':args,'id':str(self.calls),'type':'tool_call'}])
            return Bound()
    signals=[{'sheet':'SW','row':i+2,'source':f'Источник {i:02}'} for i in range(55)]
    model=Model()
    worker=ReportWorker(model,'data','собрать данные',[
        ('inventory_sources','Собрать учёт.',SourceInventory()),('finish','Передать данные.',lambda state: {})])
    state=worker({'signals':signals,'topics':[],'detail_sheets':['SW']})
    assert state['source_inventory']['detail_sheets'][0]['source_count']==55
    assert any('Источник 54' in result for result in model.results)
    assert state['agent_decisions'][0]['tools']==['inventory_sources','inspect_sheets','inspect_sources','finish']


def test_inventory_distinguishes_period_rows_exclusions_and_invalid_rows(tmp_path):
    w=Workbook(); p=w.active; p.title='P'; p.append(['продукт','тема','общий балл']); p.append(['P','T',9])
    d=w.create_sheet('SW'); d.append(['дата','источник','продукт','тема','сигнал'])
    d.append(['01.10.2026','Чат','P','T',1])
    d.append(['01.09.2026','Чат','P','T',1])
    d.append(['не дата','Чат','P','T',1])
    d.append(['01.10.2026','Чат',None,'T',1])
    book=tmp_path/'book.xlsx';w.save(book)
    from golos_mas.agents import ExcelAgent
    loaded=ExcelAgent()({'workbook':str(book),'current_sheet':'P','report_month':'2026-10'})
    sheet=SourceInventory()(loaded)['source_inventory']['detail_sheets'][0]
    assert sheet['rows_in_period']==1 and sheet['rows_excluded_other_month']==1
    assert sheet['rows_excluded_unknown_date']==1 and sheet['skipped_invalid_rows']==1
    assert sheet['parsed_rows_all_periods']==3 and sheet['nonempty_rows_seen']==4
