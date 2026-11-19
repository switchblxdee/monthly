"""Optional explicit monthly snapshot of promises, independent of the feedback workbook."""
import json
from collections import defaultdict
from datetime import date
from pathlib import Path

OUTCOMES={'completed_on_time','completed_early','postponed','status_missing','in_progress'}
REQUIRED={'id','owner','description','source','initial_due','outcome','in_work_start','in_work_current'}
OPTIONAL={'current_due','comment'}

def parse_date(value,label):
    if not isinstance(value,str): raise ValueError(f'{label}: дата должна быть строкой YYYY-MM-DD')
    try:
        parsed=date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f'{label}: ожидается дата YYYY-MM-DD') from exc
    if parsed.isoformat()!=value: raise ValueError(f'{label}: ожидается дата YYYY-MM-DD')
    return parsed

class PromisesAgent:
    def __call__(self,state):
        path=state.get('promises_json')
        if not path: return {}
        month=state.get('report_month')
        if not month: raise ValueError('Для обещаний требуется --report-month YYYY-MM')
        try:
            if len(month)!=7 or date.fromisoformat(month+'-01').strftime('%Y-%m')!=month: raise ValueError()
        except (TypeError, ValueError): raise ValueError('Для обещаний требуется --report-month YYYY-MM') from None
        records=json.loads(Path(path).read_text(encoding='utf8'))
        if not isinstance(records,list): raise ValueError('Файл обещаний должен содержать JSON-массив')
        validated=[];ids=set()
        for i,r in enumerate(records,1):
            prefix=f'Обещание {i}'
            if not isinstance(r,dict): raise ValueError(f'{prefix}: ожидается объект')
            if REQUIRED-set(r): raise ValueError(f'{prefix}: отсутствуют поля {sorted(REQUIRED-set(r))}')
            if set(r)-REQUIRED-OPTIONAL: raise ValueError(f'{prefix}: неизвестные поля {sorted(set(r)-REQUIRED-OPTIONAL)}')
            for key in ('id','owner','description','source'):
                if not isinstance(r[key],str) or not r[key].strip(): raise ValueError(f'{prefix}: поле {key} должно быть непустой строкой')
            if r['id'] in ids: raise ValueError(f'{prefix}: повтор id')
            ids.add(r['id'])
            if not isinstance(r['outcome'],str) or r['outcome'] not in OUTCOMES: raise ValueError(f'{prefix}: недопустимый outcome')
            for key in ('in_work_start','in_work_current'):
                if type(r[key]) is not bool: raise ValueError(f'{prefix}: {key} должен быть boolean')
            if 'comment' in r and not isinstance(r['comment'],str): raise ValueError(f'{prefix}: comment должен быть строкой')
            initial=parse_date(r['initial_due'],f'{prefix}.initial_due')
            current=parse_date(r['current_due'],f'{prefix}.current_due') if r.get('current_due') else None
            if r['outcome']=='postponed' and (current is None or current<=initial): raise ValueError(f'{prefix}: для переноса нужен новый срок позднее исходного')
            validated.append(dict(r,source_record=i))
        due=[r for r in validated if r['initial_due'][:7]==month]
        counts={k:sum(r['outcome']==k for r in due) for k in OUTCOMES}
        groups=defaultdict(lambda: {'start':0,'due':0,'completed':0,'postponed':0})
        for r in validated:
            g=groups[r['owner']];g['start']+=int(r['in_work_start'])
        for r in due:
            g=groups[r['owner']];g['due']+=1
            g['completed']+=int(r['outcome'] in ('completed_on_time','completed_early'))
            g['postponed']+=int(r['outcome']=='postponed')
        report={'month':month,'current_total':sum(r['in_work_current'] for r in validated),'due_count':len(due),'counts':counts,
                'completed_count':counts['completed_on_time']+counts['completed_early'],
                'due_in_work':sum(r['in_work_current'] for r in due),
                'groups':[{'owner':owner,**v} for owner,v in groups.items()],
                'postponed':[r for r in due if r['outcome']=='postponed'],
                'source_file':str(path)}
        warnings=list(state.get('warnings',[]))
        if counts['in_progress']:
            warnings.append(f"Обещания: {counts['in_progress']} со сроком в отчётном месяце остаются в работе; их итог не определён")
        return {'promise_report':report,'warnings':warnings}
