"""Conservative Russian month/sheet discovery; ambiguous cases never choose a sheet."""
import re
from datetime import datetime

MONTH_FORMS=(('январь','января'),('февраль','февраля'),('март','марта'),('апрель','апреля'),('май','мая'),('июнь','июня'),('июль','июля'),('август','августа'),('сентябрь','сентября'),('октябрь','октября'),('ноябрь','ноября'),('декабрь','декабря'))


def parse_month(value):
    if not isinstance(value,str) or not re.fullmatch(r'\d{4}-\d{2}',value): raise ValueError('Отчётный месяц должен иметь формат YYYY-MM')
    try: parsed=datetime.strptime(value,'%Y-%m')
    except ValueError: raise ValueError('Отчётный месяц должен иметь формат YYYY-MM') from None
    return parsed.year,parsed.month


def format_period(value):
    year,month=parse_month(value)
    return f'{MONTH_FORMS[month-1][0]} {year}'


def previous_month(value):
    year,month=parse_month(value)
    return f'{year-1:04d}-12' if month==1 else f'{year:04d}-{month-1:02d}'

def dated_priority_months(sheet_names):
    """Explicitly year-qualified priority tabs, without assuming a year."""
    result=[]
    for name in sheet_names:
        match=re.fullmatch(r'приоритет(?:ы)?[\s_-]+([а-я]+)[\s_-]+(\d{4})',name.strip().casefold().replace('ё','е'))
        if match:
            for i,forms in enumerate(MONTH_FORMS,1):
                if match.group(1) in forms: result.append(f'{int(match.group(2)):04d}-{i:02d}')
    return result


def discover_priority_sheet(sheet_names,month,required):
    year,number=parse_month(month);forms=MONTH_FORMS[number-1]
    qualified=[];unqualified=[]
    for name in sheet_names:
        normalized=re.sub(r'\s+',' ',name.strip().casefold().replace('ё','е'))
        match=re.fullmatch(r'приоритет(?:ы)?[\s_-]+([а-я]+)(?:[\s_-]+(\d{4}))?',normalized)
        if not match or match.group(1) not in forms: continue
        if match.group(2) is None: unqualified.append(name)
        elif int(match.group(2))==year: qualified.append(name)
    candidates=qualified or unqualified
    if len(candidates)>1:
        raise ValueError(f'Несколько листов приоритетов для {month}; укажите лист явно')
    if not candidates:
        if required: raise ValueError(f'Не найден лист приоритетов для {month}; укажите --current-sheet')
        return None,False
    return candidates[0],not bool(qualified)
