"""Optional, explicit mitigation metadata. No statuses are inferred from text."""
import json
from collections import Counter
from pathlib import Path
from .agents import norm

CONTROLS={'aibolit','promise_monitor','pain_candidate','backlog'}
MITIGATIONS={'done','analysis','release','limitations'}
RECURRENCE={'new','two','long','resumed'}
FIELDS={'product','topic','controls','mitigation','deadline','version','promnenie','after_close','recurrence'}

class StatusAgent:
    def __call__(self,state):
        path=state.get('statuses_json')
        if not path: return {}
        records=json.loads(Path(path).read_text(encoding='utf8'))
        if not isinstance(records,list): raise ValueError('Файл статусов должен содержать JSON-массив')
        bykey={}
        for i,record in enumerate(records,1):
            if not isinstance(record,dict): raise ValueError(f'Статус {i}: ожидается объект')
            unknown=set(record)-FIELDS
            if unknown: raise ValueError(f'Статус {i}: неизвестные поля {sorted(unknown)}')
            if not isinstance(record.get('product'),str) or not record['product'].strip() or not isinstance(record.get('topic'),str) or not record['topic'].strip(): raise ValueError(f'Статус {i}: нужны product и topic')
            key=(norm(record['product']),norm(record['topic']))
            if key in bykey: raise ValueError(f'Статус {i}: продукт и тема повторяются')
            controls=record.get('controls',[])
            if not isinstance(controls,list) or any(not isinstance(x,str) or x not in CONTROLS for x in controls) or len(set(controls))!=len(controls): raise ValueError(f'Статус {i}: недопустимые controls')
            if record.get('recurrence') is not None and (not isinstance(record['recurrence'],str) or record['recurrence'] not in RECURRENCE): raise ValueError(f'Статус {i}: недопустимая recurrence')
            if record.get('mitigation') is not None and (not isinstance(record['mitigation'],str) or record['mitigation'] not in MITIGATIONS): raise ValueError(f'Статус {i}: недопустимая mitigation')
            for field in ('promnenie','after_close'):
                if field in record and not isinstance(record[field],bool): raise ValueError(f'Статус {i}: {field} должен быть boolean')
            for field in ('deadline','version'):
                if field in record and not isinstance(record[field],str): raise ValueError(f'Статус {i}: {field} должен быть строкой')
            bykey[key]=(i,record)
        counts=Counter((norm(t['product']),norm(t['topic'])) for t in state['topics'])
        used=set();topics=[];warnings=list(state.get('warnings',[]))
        for topic in state['topics']:
            t=dict(topic);key=(norm(t['product']),norm(t['topic']))
            if key in bykey and counts[key]==1:
                n,status=bykey[key];t['status_record']=dict(status);t['status_source']={'file':str(path),'record':n};used.add(key)
                recurrence=status.get('recurrence');duration=t.get('duration')
                if (recurrence=='new' and duration in (2,3,4)) or (recurrence=='two' and duration in (1,3,4)) or (recurrence=='long' and duration in (1,2,3)):
                    warnings.append(f"{t['sheet']}!{t['row']}: явная повторяемость противоречит длительности; применена явная отметка")
            elif key in bykey:
                warnings.append(f"{t['sheet']}!{t['row']}: статус не применён, тема неоднозначна в приоритетах")
            topics.append(t)
        for key,(i,_) in bykey.items():
            if key not in used: warnings.append(f'Статус #{i} не сопоставлен однозначно с темой текущего периода')
        return {'topics':topics,'warnings':warnings}
