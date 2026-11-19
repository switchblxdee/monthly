import json
from pathlib import Path
import pytest
from zipfile import ZipFile
from golos_mas.demo import create_demo
from golos_mas.graph import create_graph


def test_demo_does_not_replace_existing_directory(tmp_path):
    folder=tmp_path/'exists';folder.mkdir();marker=folder/'important.txt';marker.write_text('original')
    with pytest.raises(FileExistsError): create_demo(folder)
    assert marker.read_text()=='original'


@pytest.mark.skipif(not __import__('os').environ.get('GOLOS_TEST_TEMPLATE'),reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_demo_runs_across_all_agents_and_known_counts(tmp_path):
    import os
    root=create_demo(tmp_path/'generated');expected=json.loads((root/'expected.json').read_text(encoding='utf8'));out=root/'report.pptx'
    state=create_graph().invoke({'workbook':str(root/'demo.xlsx'),'current_sheet':'приоритет ноябрь','previous_sheet':'приоритет октябрь',
        'period':'ноябрь 2026','report_month':'2026-11','template':os.environ['GOLOS_TEST_TEMPLATE'],
        'statuses_json':str(root/'statuses.json'),'promises_json':str(root/'promises.json'),'topic_map_json':str(root/'topic-map.json'),'output':str(out)})
    c=state['coverage'];p=state['promise_report']
    assert c['priority_topics']==expected['priority_topics'] and c['detail_rows']==expected['detail_rows_after_filter']
    assert c['signal_sum']==expected['sum_signals'] and c['unmatched_detail_rows']==0
    assert p['due_count']==expected['promise_due'] and p['completed_count']==expected['promise_completed'] and p['counts']['postponed']==expected['promise_postponed']
    assert not state['warnings']
    with ZipFile(out) as deck:
        cover=deck.read('ppt/slides/slide1.xml').decode()
        assert 'мониторинг обещаний' in cover
        allslides=[deck.read(f'ppt/slides/slide{i}.xml').decode() for i in range(1,state['slide_count']+1)]
        promises_separator=next(slide for slide in allslides if 'Мониторинг обещаний' in slide and '>04<' in slide)
        assert promises_separator

def test_cli_can_discover_month_and_period_from_demo(tmp_path):
    import os, subprocess, sys
    import golos_mas
    root=create_demo(tmp_path/'generated');out=root/'report.pptx'
    env=dict(os.environ,PYTHONPATH=str(Path(golos_mas.__file__).resolve().parent.parent))
    result=subprocess.run([sys.executable,'-m','golos_mas.cli',str(root/'demo.xlsx'),'--report-month','2026-11','--output',str(out)],capture_output=True,text=True,env=env,timeout=120)
    assert result.returncode==0,result.stderr
    audit=json.loads(out.with_suffix('.audit.json').read_text(encoding='utf8'))
    assert audit['source_sheets']['current']=='приоритет ноябрь' and audit['source_sheets']['previous']=='приоритет октябрь'
    assert audit['coverage']['priority_topics']==4 and audit['coverage']['detail_rows']==5
    assert audit['date_filter']['other_month_references']==[{'sheet':'SU','row':4,'month':'2026-10'}]
    assert audit['date_filter']['kept_rows']==5 and audit['date_filter']['unknown_date_rows']==0
    assert any('Год' in w for w in audit['warnings'])

@pytest.mark.skipif(not __import__('os').environ.get('GOLOS_TEST_TEMPLATE'),reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_owner_named_like_promises_section_keeps_both_contents_entries(tmp_path):
    import os
    from openpyxl import load_workbook
    root=create_demo(tmp_path/'generated');book=root/'demo.xlsx'
    w=load_workbook(book);w['приоритет ноябрь'].cell(2,1).value='Мониторинг обещаний';w.save(book)
    out=root/'report.pptx'
    create_graph().invoke({'workbook':str(book),'current_sheet':'приоритет ноябрь','period':'ноябрь 2026',
        'report_month':'2026-11','template':os.environ['GOLOS_TEST_TEMPLATE'],'promises_json':str(root/'promises.json'),
        'topic_map_json':str(root/'topic-map.json'),'output':str(out)})
    with ZipFile(out) as deck:
        toc=deck.read('ppt/slides/slide3.xml').decode()
        assert 'Рабочая среда A' in toc and 'Статистика и переносы' in toc
        assert 'Слайды 5–7' in toc and 'Мониторинг обещаний' in toc
