"""Product-level scores: retain source rows and refuse conflicting inputs."""
from decimal import Decimal, InvalidOperation

FIELDS=('csat','csat_delta','opinion_score','opinion_delta')

def parse_metric(value):
    if value is None or isinstance(value,bool): return None
    text=str(value).strip().replace('\u00a0','').replace(' ','').replace(',','.')
    if not text: return None
    try: number=Decimal(text)
    except InvalidOperation: return None
    return number if number.is_finite() else None

def product_metrics(topics):
    result={}
    for key in FIELDS:
        found=[];invalid=[]
        for t in topics:
            raw=t.get(key)
            if key in t.get('uncached_formulas',[]):
                invalid.append({'sheet':t.get('sheet'),'row':t.get('row')});continue
            if raw is None or (isinstance(raw,str) and not raw.strip()): continue
            value=parse_metric(raw)
            ref={'sheet':t.get('sheet'),'row':t.get('row')}
            if value is None: invalid.append(ref)
            else: found.append((value,ref))
        distinct={v for v,_ in found}
        result[key]={'value':format(next(iter(distinct)),'f') if len(distinct)==1 and not invalid else None,
                     'references':[ref for _,ref in found], 'invalid_references':invalid,
                     'conflict':len(distinct)>1}
    return result

def format_metric(entry, signed=False):
    value=entry.get('value')
    if value is None: return '—'
    num=Decimal(value)
    text=format(num,'f').replace('.',',')
    return ('+' if signed and num>0 else '')+text
