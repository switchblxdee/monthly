"""Validate signal scoring against the methodology shown on slide 4."""
from .agents import norm, valid_signal

MISSING_MAU={'','н.д.','н.д','н/д','нет данных','n/a','na','неизвестно'}
RULES={'ambiguous','upper','lower'}


def evaluate_signal_score(mau,count,boundaries='upper'):
    if boundaries not in RULES: raise ValueError('Неизвестное правило границ MAU')
    result={'audience_category':None,'audience_basis':None,'expected_signal_score':None,'reason':None,'boundary_rule':boundaries}
    if mau is None or (isinstance(mau,str) and norm(mau) in MISSING_MAU):
        category='средний';result['audience_basis']='default_missing'
    elif not valid_signal(mau):
        result.update(audience_basis='invalid',reason='invalid_mau');return result
    else:
        result['audience_basis']='provided'
        if boundaries=='ambiguous' and mau in (1000,20000):
            result.update(audience_basis='boundary',reason='ambiguous_boundary');return result
        if mau<1000 or (boundaries=='lower' and mau==1000): category='малый'
        elif mau<20000 or (boundaries=='lower' and mau==20000): category='средний'
        else: category='большой'
    result['audience_category']=category
    if count is None:
        result['reason']='missing_count';return result
    if not valid_signal(count):
        result['reason']='invalid_count';return result
    if count<10:
        result['reason']='below_threshold';return result
    thresholds={'малый':(20,30,40),'средний':(30,40,50),'большой':(50,60,70)}[category]
    result['expected_signal_score']=1+sum(count>=x for x in thresholds)
    return result


class PrioritizationAgent:
    def __call__(self,state):
        rule=state.get('mau_boundaries','upper')
        if rule not in RULES: raise ValueError('Неизвестное правило границ MAU')
        warnings=list(state.get('warnings',[]));topics=[]
        for source in state['topics']:
            t=dict(source);uncached=set(t.get('uncached_formulas',[]))
            if 'mau' in uncached:
                check={'audience_category':None,'audience_basis':'uncached_formula','expected_signal_score':None,'reason':'uncached_mau','boundary_rule':rule}
            else: check=evaluate_signal_score(t.get('mau'),t.get('count'),rule)
            t['methodology_check']=check
            ref=f"{t['sheet']}!{t['row']}";reason=check['reason'];actual=t.get('signals_score');expected=check['expected_signal_score']
            t['signals_score_input']=actual
            t['signals_score_origin']='excel' if actual is not None else 'missing'
            if actual is None and expected is not None and norm(t.get('score'))!='ux':
                t['signals_score']=expected;t['signals_score_origin']='methodology'
                warnings.append(f'{ref}: балл за сигналы отсутствовал; рассчитан по методике: {expected}')
            if reason=='ambiguous_boundary': warnings.append(f'{ref}: MAU на неоднозначной границе; балл по сигналам не проверен')
            elif reason=='invalid_mau': warnings.append(f'{ref}: некорректное значение MAU; балл по сигналам не проверен')
            elif reason=='invalid_count': warnings.append(f'{ref}: некорректное количество сигналов; балл по сигналам не проверен')
            elif reason=='below_threshold' and norm(t.get('score'))!='ux': warnings.append(f'{ref}: меньше 10 сигналов; методика не задаёт балл')
            if expected is not None and check['audience_basis']=='default_missing': warnings.append(f'{ref}: MAU не указан; для проверки использована средняя аудитория по методике')
            if expected is not None and valid_signal(actual) and actual!=expected: warnings.append(f'{ref}: балл за сигналы в Excel {actual}, по методике {expected}; значение не изменено')
            topics.append(t)
        return {'topics':topics,'warnings':warnings}
