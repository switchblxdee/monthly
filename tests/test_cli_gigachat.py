import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'), reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_cli_configures_model_and_processes_selected_evidence(tmp_path):
    import golos_mas
    fake=tmp_path/'langchain_gigachat.py'
    fake.write_text('''import json, os\nfrom types import SimpleNamespace\nfrom langchain_core.messages import AIMessage\nclass GigaChat:\n def __init__(self, **kwargs):\n  self.__dict__.update(kwargs)\n  open(os.environ["FAKE_GIGA_LOG"], "w").write(json.dumps(kwargs))\n def bind_tools(self, tools):\n  class Bound:\n   def invoke(self, messages):\n    data=json.loads(messages[-1].content); status=data.get("status", data)\n    return AIMessage(content="", tool_calls=[{"name":status["remaining"][0],"args":{},"id":"call", "type":"tool_call"}])\n  return Bound()\n def invoke(self, messages):\n  return SimpleNamespace(content=\'{"номера":[1]}\')\n''',encoding='utf8')
    w=Workbook();s=w.active;s.title='P';s.append(['владелец','продукт','тема','общий балл']);s.append(['Команда','Продукт','Ошибка',9])
    d=w.create_sheet('SW');d.append(['источник','продукт','тема','Маскированный текст из источника','лайки','сигнал']);d.append(['Чат','Продукт','Ошибка','Тестовое обращение',7,8])
    book=tmp_path/'input.xlsx';w.save(book);out=tmp_path/'out.pptx';log=tmp_path/'constructor.json'
    env=dict(os.environ,PYTHONPATH=os.pathsep.join([str(tmp_path),str(Path(golos_mas.__file__).resolve().parent.parent)]),
             GIGACHAT_BASE_URL='https://model.internal/v1',GIGACHAT_AUTH_URL='https://auth.internal/token',
             GIGACHAT_CREDENTIALS='fake',GIGACHAT_MODEL='Internal-GigaChat',FAKE_GIGA_LOG=str(log))
    cmd=[sys.executable,'-m','golos_mas.cli',str(book),'--current-sheet','P','--period','месяц',
         '--template',os.environ['GOLOS_TEST_TEMPLATE'],'--gigachat','--gigachat-quotes','--output',str(out)]
    result=subprocess.run(cmd,env=env,capture_output=True,text=True,timeout=120)
    assert result.returncode==0,result.stderr
    assert json.loads(log.read_text())=={'base_url':'https://model.internal/v1','auth_url':'https://auth.internal/token','model':'Internal-GigaChat'}
    audit=json.loads(out.with_suffix('.audit.json').read_text(encoding='utf8'))
    assert audit['topics'][0]['quote_selection']=='gigachat' and audit['topics'][0]['quotes'][0]['row']==2
    quote=audit['topics'][0]['quotes'][0]
    assert quote['likes']==7 and quote['likes_for_ordering']==7 and quote['signal']==8
    assert audit['narrative_sources']==[{'sheet':'P','row':2}]


def test_cli_rejects_ignored_endpoint_before_model_request_or_report_write(tmp_path):
    import golos_mas
    fake=tmp_path/'langchain_gigachat.py'
    fake.write_text('''import os\nclass GigaChat:\n def __init__(self, **kwargs): self.base_url=kwargs["base_url"]\n def invoke(self, messages): open(os.environ["FAKE_GIGA_LOG"], "w").write("called")\n''',encoding='utf8')
    out=tmp_path/'existing.pptx';out.write_bytes(b'previous report');log=tmp_path/'called'
    env=dict(os.environ,PYTHONPATH=os.pathsep.join([str(tmp_path),str(Path(golos_mas.__file__).resolve().parent.parent)]),
             GIGACHAT_BASE_URL='https://model.internal/v1',GIGACHAT_AUTH_URL='https://auth.internal/token',
             GIGACHAT_CREDENTIALS='fake-secret',FAKE_GIGA_LOG=str(log))
    cmd=[sys.executable,'-m','golos_mas.cli',str(tmp_path/'unread.xlsx'),'--current-sheet','P','--period','месяц','--gigachat','--output',str(out)]
    result=subprocess.run(cmd,env=env,capture_output=True,text=True,timeout=30)
    assert result.returncode!=0 and 'не подтвердила параметр auth_url' in result.stderr
    assert 'auth.internal' not in result.stderr and 'fake-secret' not in result.stderr
    assert not log.exists() and out.read_bytes()==b'previous report'
