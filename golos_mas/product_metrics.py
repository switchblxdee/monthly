"""Product-level scores: retain source rows and report observed ranges."""
from decimal import Decimal, InvalidOperation

FIELDS=('csat','csat_delta','opinion_score','opinion_delta')

def parse_metric(value):
    if value is None or isinstance(value,bool): return None
    text=''.join(str(value).split()).replace(',', '.').replace('−', '-').replace('–', '-')
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
            if raw is None or (isinstance(raw, str) and raw.strip().casefold() in ('', '-', '—', '–', 'н.д.', 'н/д', 'нет данных')): continue
            value=parse_metric(raw)
            ref={'sheet':t.get('sheet'),'row':t.get('row')}
            if value is None: invalid.append(ref)
            else: found.append((value,ref))
        distinct={v for v,_ in found}
        result[key]={'value':format(next(iter(distinct)),'f') if len(distinct)==1 else None,
                     'references':[ref for _,ref in found], 'invalid_references':invalid,
                     'conflict':len(distinct)>1,
                     'observed_values':[format(value, 'f') for value in sorted(distinct)]}
    return result

def format_metric(entry, signed=False):
    def display(value):
        number = Decimal(value)
        text = format(number, 'f').replace('.', ',')
        return ('+' if signed and number > 0 else '') + text

    value = entry.get('value')
    if value is not None:
        return display(value)
    observed = entry.get('observed_values', [])
    if len(observed) > 1:
        return f'{display(observed[0])}…{display(observed[-1])}'
    return '—'
