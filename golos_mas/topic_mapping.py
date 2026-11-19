"""Explicit exceptions for topics labelled differently in priority and detail sheets."""
import json
from collections import defaultdict
from pathlib import Path
from .agents import norm

FIELDS={'priority_product','priority_topic','detail_product','detail_topic'}

class TopicMappingAgent:
    def __call__(self,state):
        path=state.get('topic_map_json')
        if not path: return {}
        rows=json.loads(Path(path).read_text(encoding='utf8'))
        if not isinstance(rows,list): raise ValueError('Карта тем должна быть JSON-массивом')
        mapping=defaultdict(list);claimed={};all_rows=set()
        for i,row in enumerate(rows,1):
            if not isinstance(row,dict) or set(row)!=FIELDS: raise ValueError(f'Карта тем #{i}: нужны поля {sorted(FIELDS)}')
            if any(not isinstance(row[k],str) or not row[k].strip() for k in FIELDS): raise ValueError(f'Карта тем #{i}: поля должны быть непустыми строками')
            src=(norm(row['priority_product']),norm(row['priority_topic']))
            target=(norm(row['detail_product']),norm(row['detail_topic']))
            if (src,target) in all_rows: raise ValueError(f'Карта тем #{i}: повтор сопоставления')
            if target in claimed and claimed[target]!=src: raise ValueError(f'Карта тем #{i}: одна тема детализации назначена разным темам приоритетов')
            all_rows.add((src,target));claimed[target]=src;mapping[src].append(target)
        existing={(norm(t['product']),norm(t['topic'])) for t in state['topics']}
        warnings=list(state.get('warnings',[]))
        for key in mapping.keys()-existing: warnings.append('Карта тем: одна из тем приоритетов не найдена в текущем периоде')
        return {'topic_map':dict(mapping),'warnings':warnings}
