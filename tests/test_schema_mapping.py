import json
from pathlib import Path
import pytest
from openpyxl import Workbook
from golos_mas.agents import ExcelAgent, AnalyticsAgent, EvidenceAgent
from golos_mas.inspect_workbook import inspect_workbook


def make_workbook(path):
    w=Workbook();p=w.active;p.title='PRIVATE_PRIORITY'
    p.append(['introduction']);p.append(['group','asset','issue','audience','volume','rating'])
    p.append(['PRIVATE_OWNER','PRIVATE_PRODUCT','PRIVATE_TOPIC',30000,20,9])
    d=w.create_sheet('PRIVATE_DETAIL');d.append(['intro']);d.append(['intro'])
    d.append(['when','where','asset','issue','message','reactions','volume'])
    d.append(['01.11.2026','PRIVATE_SOURCE','PRIVATE_PRODUCT','PRIVATE_TOPIC','PRIVATE_MESSAGE',19,20])
    w.save(path)


def schema():
    return {'version':1,'sheets':{
        'PRIVATE_PRIORITY':{'header_row':2,'columns':{'owner':'A','product':'B','topic':'C','mau':'D','count':'E','score':'F'}},
        'PRIVATE_DETAIL':{'header_row':3,'columns':{'date':'A','source':'B','product':'C','topic':'D','text':'E','likes':'F','signal':'G'}}}}


def test_nonstandard_columns_are_shared_by_ingest_and_inspection(tmp_path):
    path=tmp_path/'input.xlsx';make_workbook(path);config=tmp_path/'schema.json';config.write_text(json.dumps(schema()),encoding='utf8')
    state={'workbook':str(path),'current_sheet':'PRIVATE_PRIORITY','detail_sheets':['PRIVATE_DETAIL'],'schema_json':str(config)}
    state.update(ExcelAgent()(state));state.update(AnalyticsAgent()(state));state.update(EvidenceAgent()(state))
    assert state['topics'][0]['product']=='PRIVATE_PRODUCT' and state['topics'][0]['mau']==30000
    assert state['topics'][0]['signal_sum']==20 and state['topics'][0]['quotes'][0]['text']=='PRIVATE_MESSAGE'
    report=inspect_workbook(path,schema_path=config)
    assert [s['kind'] for s in report['sheets']]==['priority','detail']
    assert all(s['schema_applied'] and s['rows_with_product_and_topic']==1 for s in report['sheets'])
    assert 'PRIVATE_' not in json.dumps(report,ensure_ascii=False)


@pytest.mark.parametrize('field,value',[('topic','B'),('product','XFE'),('unknown','A')])
def test_invalid_column_mappings_fail_closed(tmp_path,field,value):
    path=tmp_path/'i.xlsx';make_workbook(path);config=schema()
    config['sheets']['PRIVATE_PRIORITY']['columns'][field]=value
    cfg=tmp_path/'s.json';cfg.write_text(json.dumps(config),encoding='utf8')
    with pytest.raises(ValueError):
        ExcelAgent()({'workbook':str(path),'current_sheet':'PRIVATE_PRIORITY','schema_json':str(cfg)})


def test_duplicate_headers_require_explicit_mapping_and_inspection_reports_it(tmp_path):
    w=Workbook();s=w.active;s.title='P';s.append(['продукт','продукт','тема','общий балл']);s.append(['WRONG','RIGHT','T',9])
    path=tmp_path/'i.xlsx';w.save(path)
    with pytest.raises(ValueError,match='повторяются распознанные заголовки'):
        ExcelAgent()({'workbook':str(path),'current_sheet':'P'})
    assert 'product' in inspect_workbook(path)['sheets'][0]['duplicate_recognized_fields']
    config={'version':1,'sheets':{'P':{'header_row':1,'columns':{'product':'B','topic':'C','score':'D'}}}}
    schema_file=tmp_path/'schema.json';schema_file.write_text(json.dumps(config),encoding='utf8')
    result=ExcelAgent()({'workbook':str(path),'current_sheet':'P','schema_json':str(schema_file)})
    assert result['topics'][0]['product']=='RIGHT'
    assert 'product' not in inspect_workbook(path,schema_path=schema_file)['sheets'][0]['duplicate_recognized_fields']


def test_custom_schema_cli_generates_expected_audit(tmp_path):
    import os,subprocess,sys
    import golos_mas
    path=tmp_path/'i.xlsx';make_workbook(path);cfg=tmp_path/'s.json';cfg.write_text(json.dumps(schema()),encoding='utf8')
    out=tmp_path/'o.pptx';env=dict(os.environ,PYTHONPATH=str(Path(golos_mas.__file__).resolve().parent.parent))
    run=subprocess.run([sys.executable,'-m','golos_mas.cli',str(path),'--current-sheet','PRIVATE_PRIORITY','--detail-sheet','PRIVATE_DETAIL','--schema-json',str(cfg),'--period','месяц','--output',str(out)],env=env,capture_output=True,text=True,timeout=120)
    assert run.returncode==0,run.stderr
    audit=json.loads(out.with_suffix('.audit.json').read_text(encoding='utf8'))
    assert audit['topics'][0]['quotes'][0]['text']=='PRIVATE_MESSAGE' and audit['coverage']['signal_sum']==20
    assert audit['source_sheets']['schema']==str(cfg)
