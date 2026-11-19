import pytest
from openpyxl import Workbook
from golos_mas.periods import discover_priority_sheet,format_period,previous_month
from golos_mas.agents import ExcelAgent,AnalyticsAgent


def test_year_boundary_and_qualified_precedence():
    names=['приоритет январь','Приоритет января 2026','приоритет декабрь 2025','приоритет декабрь 2024']
    assert discover_priority_sheet(names,'2026-01',True)==('Приоритет января 2026',False)
    assert previous_month('2026-01')=='2025-12'
    assert discover_priority_sheet(names,previous_month('2026-01'),False)==('приоритет декабрь 2025',False)
    assert format_period('2026-01')=='январь 2026'


def test_ambiguous_month_requires_explicit_choice():
    names=['приоритет ноябрь','приоритет ноября']
    with pytest.raises(ValueError,match='Несколько листов'):
        discover_priority_sheet(names,'2026-11',True)


def test_auto_discovery_drives_previous_period_and_can_be_disabled(tmp_path):
    w=Workbook();cur=w.active;cur.title='приоритет январь 2026';cur.append(['продукт','тема','общий балл']);cur.append(['P','T',9])
    old=w.create_sheet('приоритет декабрь 2025');old.append(['продукт','тема']);old.append(['P','T'])
    path=tmp_path/'i.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'report_month':'2026-01'})
    assert state['current_sheet']==cur.title and state['previous_sheet']==old.title
    assert AnalyticsAgent()(state)['topics'][0]['recurrence']=='есть в прошлом месяце'
    without=ExcelAgent()({'workbook':str(path),'report_month':'2026-01','auto_previous':False})
    assert without['previous_sheet'] is None and AnalyticsAgent()(without)['topics'][0]['recurrence']=='неизвестно'


def test_yearless_discovery_is_explicitly_flagged(tmp_path):
    w=Workbook();ws=w.active;ws.title='приоритет ноябрь';ws.append(['продукт','тема']);ws.append(['P','T']);path=tmp_path/'i.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'report_month':'2026-11'})
    assert state['current_sheet']==ws.title and any('Год' in warning for warning in state['warnings'])
