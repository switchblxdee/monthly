import json
from openpyxl import Workbook
from golos_mas.inspect_workbook import inspect_workbook, main


def test_inspector_excludes_sensitive_cell_values_and_sheet_names(tmp_path):
    w=Workbook();s=w.active;s.title='SECRET_SHEET_NAME'
    s.append(['продукт','тема','общий балл','владелец'])
    s.append(['SECRET_PRODUCT','SECRET_TOPIC','=4+5','SECRET_OWNER'])
    s.append(['ANOTHER_PRODUCT','ANOTHER_TOPIC',8,'SECRET_OWNER'])
    d=w.create_sheet('SECRET_SOURCE_SHEET');d.append(['продукт','тема','источник','Маскированный текст из источника','сигнал'])
    d.append(['SECRET_PRODUCT','SECRET_TOPIC','SECRET_SOURCE','SECRET_MESSAGE',3])
    path=tmp_path/'SECRET_FILENAME.xlsx';w.save(path)
    report=inspect_workbook(path,max_rows=1);payload=json.dumps(report,ensure_ascii=False)
    for secret in ('SECRET','ANOTHER_PRODUCT','=4+5','SECRET_FILENAME.xlsx'):
        assert secret not in payload
    assert report['sheet_count']==2 and report['sheets'][0]['sheet_alias']=='Sheet_1'
    assert report['sheets'][0]['truncated'] and report['sheets'][0]['rows_scanned']==1
    assert report['sheets'][0]['formulas']['score']=={'total':1,'missing_cached_value':1,'error_cached_value':0}
    assert report['sheets'][1]['kind']=='detail'


def test_inspector_will_not_overwrite_input(tmp_path):
    import pytest
    w=Workbook();path=tmp_path/'input.xlsx';w.save(path);before=path.read_bytes()
    with pytest.raises(SystemExit): main([str(path),'--output',str(path)])
    assert path.read_bytes()==before


def test_inspector_counts_nonempty_rows_missing_both_keys_without_exposing_values(tmp_path):
    w=Workbook();s=w.active;s.title='PRIVATE'
    s.append(['продукт','тема','источник','маскированный текст из источника','сигнал'])
    s.append([None,None,'PRIVATE_SOURCE','PRIVATE_COMMENT',1])
    s.append(['  ','  ',None,None,None])
    s.append(['P','  ','Чат','Текст',1])
    s.append(['="PRIVATE_PRODUCT"','="PRIVATE_TOPIC"',None,None,None])
    path=tmp_path/'book.xlsx';w.save(path)
    report=inspect_workbook(path);sheet=report['sheets'][0]
    assert sheet['rows_with_no_keys_but_other_values']==1 and sheet['rows_with_only_one_key']==1
    assert sheet['rows_with_no_keys_and_uncached_formula']==1
    assert 'PRIVATE' not in json.dumps(report,ensure_ascii=False)

def test_inspector_rejects_mismatched_formula_and_value_row_counts(tmp_path,monkeypatch):
    import pytest
    from golos_mas import inspect_workbook as inspector
    w1=Workbook();w1.active.title='P';w1.active.append(['продукт','тема']);w1.active.append(['P','T']);w1.active.append(['P2','T2'])
    w2=Workbook();w2.active.title='P';w2.active.append(['продукт','тема']);w2.active.append(['P','T']);w2.active.append(['P2','T2'])
    raw=w2.active;original=raw.iter_rows
    def truncated(*args,**kwargs):
        rows=original(*args,**kwargs)
        if kwargs.get('min_row',1)>1: return iter(list(rows)[:1])
        return rows
    monkeypatch.setattr(raw,'iter_rows',truncated)
    monkeypatch.setattr(inspector,'load_workbook',lambda path,read_only,data_only: w1 if data_only else w2)
    with pytest.raises(ValueError,match='Диапазоны значений и формул не совпадают'):
        inspector.inspect_workbook(tmp_path/'ignored.xlsx')

def test_dimensionless_streamed_workbook_is_read_and_limit_is_accurate(tmp_path):
    from golos_mas.agents import ExcelAgent
    w=Workbook(write_only=True);p=w.create_sheet('P');p.append(['продукт','тема','общий балл']);p.append(['P','T',9])
    d=w.create_sheet('SW');d.append(['дата','продукт','тема','сигнал']);d.append(['01.10.2026','P','T',1]);d.append(['02.10.2026','P','T',1])
    path=tmp_path/'streamed.xlsx';w.save(path)
    state=ExcelAgent()({'workbook':str(path),'current_sheet':'P','report_month':'2026-10','auto_previous':False})
    assert len(state['signals'])==2
    exact=inspect_workbook(path,max_rows=2)['sheets'][1]
    limited=inspect_workbook(path,max_rows=1)['sheets'][1]
    assert exact['rows_scanned']==2 and not exact['truncated']
    assert limited['rows_scanned']==1 and limited['truncated']
