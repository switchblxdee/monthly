"""Generate editable PPTX slides by adapting supplied PPTX template objects."""
from __future__ import annotations
import copy, math, posixpath, re, textwrap
from collections import defaultdict, deque
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
from .agents import valid_signal, display_signal_count

P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
N={'p':P,'a':A,'r':R,'pr':PR}
PROMISES_SECTION=object()
for k,v in [('p',P),('a',A),('r',R)]: ET.register_namespace(k,v)

def xml(data): return ET.fromstring(data)
def source_count_label(count):
    last_two=count%100
    word='источников' if 11<=last_two<=14 else {1:'источник',2:'источника',3:'источника',4:'источника'}.get(count%10,'источников')
    return f'{count} {word}'

def product_title(shape, value):
    text=str(value);replace_text(shape,text)
    if len(text)>35:
        size=max(1400,min(2800,round(2800*45/len(text))))
        for props in shape.findall('.//a:rPr',N): props.set('sz',str(size))

def serial(root):
    data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    if root.tag in (f'{{{PR}}}Relationships', f'{{{CT}}}Types'):
        data=data.replace(b'xmlns:ns0=',b'xmlns=').replace(b'ns0:',b'')
    return data
def content(element): return ''.join(t.text or '' for t in element.findall('.//a:t',N))
def replace_text(element,value):
    body=element.find('p:txBody',N) if element.tag==f'{{{P}}}sp' else element.find('a:txBody',N)
    if body is None: return
    paras=body.findall('a:p',N)
    if not paras: return
    for p in paras[1:]: body.remove(p)
    p=paras[0]; runs=p.findall('a:r',N)
    if runs:
        for child in list(p):
            if child.tag in (f'{{{A}}}r',f'{{{A}}}br',f'{{{A}}}fld') and child is not runs[0]: p.remove(child)
        t=runs[0].find('a:t',N)
        if t is None: t=ET.SubElement(runs[0],f'{{{A}}}t')
        t.text=str(value)
    else:
        for child in list(p):
            if child.tag in (f'{{{A}}}br',f'{{{A}}}fld'): p.remove(child)
        run=ET.Element(f'{{{A}}}r'); rp=ET.SubElement(run,f'{{{A}}}rPr',{'lang':'ru-RU','sz':'1000'}); ET.SubElement(run,f'{{{A}}}t').text=str(value)
        end=p.find('a:endParaRPr',N)
        if end is not None: p.insert(list(p).index(end),run)
        else: p.append(run)

def position(el):
    off=el.find('.//a:xfrm/a:off',N)
    return (int(off.get('x')),int(off.get('y'))) if off is not None else (None,None)

def shapes(root): return root.find('.//p:spTree',N)

def cover(root,period,include_promises=False):
    for sp in root.findall('.//p:sp',N):
        if 'октябрь 2026' in content(sp).lower(): replace_text(sp,period)
        elif 'мониторинг обещаний' in content(sp).lower(): replace_text(sp,'+ мониторинг обещаний' if include_promises else 'Мониторинг обратной связи')
    return root

def summary(root,state):
    tree=shapes(root)
    topics=state['topics']
    counts={z:sum(t['zone']==z for t in topics) for z in ('высокое внимание','потенциальный риск','наблюдение','UX')}
    group=root.find('.//p:grpSp',N)
    coverage=state.get('coverage',{})
    priority_sheet=str(state.get('current_sheet') or 'приоритеты')
    text=f"Период: {state['period']}. Тем в листе «{priority_sheet}»: {len(topics)}. Продуктов: {len({product_key(t['product']) for t in topics})}. Высокое внимание: {counts['высокое внимание']}; потенциальный риск: {counts['потенциальный риск']}; наблюдение: {counts['наблюдение']}; UX: {counts['UX']}."
    if coverage.get('detail_rows'):
        amount=coverage.get('signal_sum')
        qualifier=' (неполная)' if coverage.get('invalid_signal_rows') and amount is not None else ''
        sheets=', '.join(state.get('detail_sheets') or []) or 'не указаны'
        text+=f" Детализация ({sheets}): строк: {coverage['detail_rows']}; сумма сигналов{qualifier}: {amount if amount is not None else 'нет данных'} (строк с числом: {coverage['numeric_signal_rows']}); указано источников: {coverage['distinct_named_sources']}."
        if coverage.get('possible_duplicate_groups'): text+=' Есть возможные повторы.'
    for sp in group.findall('.//p:sp',N):
        if content(sp): replace_text(sp,text)
    for sp in list(tree):
        if sp.tag!=f'{{{P}}}sp': continue
        txt=content(sp)
        if txt.startswith('Тренды:'):
            status_note='Статусы подключены' if any(t.get('status_record') for t in topics) else 'Статусы не предоставлены'
            promise_note='реестр обещаний подключён' if state.get('promise_report') is not None else 'реестр обещаний не предоставлен'
            replace_text(sp,state.get('narrative') or f'{status_note}; {promise_note}.')
        elif txt.startswith('Все выпуски') or txt.startswith('Не забыть'): tree.remove(sp)
    table=root.find('.//a:tbl',N)
    rows=table.findall('a:tr',N)
    high=[t for t in topics if t['zone']=='высокое внимание']
    high.sort(key=lambda t: (-int(t['count']) if valid_signal(t.get('count')) else 1, -t['score']))
    for index,row in enumerate(rows[1:]):
        cells=row.findall('a:tc',N)
        if index < len(high):
            t=high[index]; status=t.get('status_record',{})
            names={'done':'Выполнено','analysis':'На анализе','release':'Релиз','limitations':'Текущие ограничения'}
            status_text=' '.join(filter(None,[names.get(status.get('mitigation'),''),status.get('version',''),status.get('deadline','')])) or 'нет данных'
            values=[f"{t['product']}: {t['topic'][:170]}{'…' if len(t['topic'])>170 else ''}",f"{t['zone']} ({t['score']}/12, сигналов: {display_signal_count(t.get('count'))})",status_text]
        else: values=['','','']
        for cell,value in zip(cells,values): replace_text(cell,value)
    return root

ZONE_COLORS={'высокое внимание':'F44552','потенциальный риск':'F5A800','наблюдение':'AAB6C8'}
ICON_DESCRIPTIONS={'new':'Новая тема','two':'Повторяется 2 месяца','long':'Повторяется более 3 месяцев','resumed':'Возобновилось обсуждение','aibolit':'АйболИТ+','backlog':'Бэклог команды','pain_candidate':'Кандидат на боль','promise_monitor':'Мониторинг обещаний','after_close':'Повтор после закрытия','done':'Выполнено','analysis':'На анализе','release':'Релиз','limitations':'Текущие ограничения','promnenie':'Упоминание в ПроМнении'}

def replace_lines(element,header,products):
    replace_text(element,header)
    body=element.find('p:txBody',N)
    if body is None: return
    p=body.find('a:p',N);first=p.find('a:r',N)
    line_count=1+sum(max(1,len(textwrap.wrap(name,width=34,break_long_words=True,break_on_hyphens=False))) for name,_ in products)
    font_size=str(max(800,min(1100,int(1100*14/max(14,line_count)))))
    for rp in body.findall('.//a:rPr',N): rp.set('sz',font_size)
    end=p.find('a:endParaRPr',N)
    def add_child(element):
        if end is None: p.append(element)
        else: p.insert(list(p).index(end),element)
    for index,(product,zone) in enumerate(products):
        add_child(ET.Element(f'{{{A}}}br'))
        if index==0: add_child(ET.Element(f'{{{A}}}br'))
        marker=copy.deepcopy(first); marker.find('a:t',N).text='◉ ' if zone else '• '
        rp=marker.find('a:rPr',N)
        if rp is not None and zone:
            for fill in rp.findall('a:solidFill',N): rp.remove(fill)
            fill=ET.Element(f'{{{A}}}solidFill');ET.SubElement(fill,f'{{{A}}}srgbClr',{'val':ZONE_COLORS[zone]});rp.insert(0,fill)
        add_child(marker)
        run=copy.deepcopy(first);run.find('a:t',N).text=product;add_child(run)

def contents_slide(root,cards,spans,page,total,high_count):
    tree=shapes(root)
    ranges={(2.5,1.1):0,(6.8,1.1):1,(11.1,1.1):2,(11.1,3.4):3,(11.1,4.7):4,(6.7,4.4):5}
    neutral_legend=None
    for child in list(tree):
        if child.tag!=f'{{{P}}}sp': continue
        text=content(child); x,y=position(child)
        if text.startswith('По 14-ти'):
            replace_text(child, f'Продуктов с темами высокого внимания: {high_count}' + (f' ({page}/{total})' if total>1 else ''))
        elif text.startswith('темы в зоне'):
            if 'наблюдение' in text: neutral_legend=child
            replace_text(child,text.replace('темы в зоне','наивысшая зона:'))
        elif text.startswith('07 Мониторинг'):
            tree.remove(child)
        elif re.match(r'^0[1-6] ',text):
            index=int(text[1])-1
            if index>=len(cards): tree.remove(child)
            else:
                owner,products=cards[index]
                group_number=list(spans).index(owner)+1
                label='Мониторинг обещаний' if owner is PROMISES_SECTION else owner
                replace_lines(child,f'{group_number:02} {label}',products)
        elif text.startswith('Слайды'):
            key=(round(x/914400,1),round(y/914400,1)) if x is not None else None
            index=ranges.get(key)
            if index is None or index>=len(cards): tree.remove(child)
            else:
                start,end=spans[cards[index][0]]
                replace_text(child,f'Слайды {start}–{end}')
    if neutral_legend is not None:
        note=copy.deepcopy(neutral_legend); cnv=note.find('.//p:cNvPr',N)
        if cnv is not None: cnv.set('id',str(1+max((int(e.get('id','0')) for e in root.findall('.//p:cNvPr',N)),default=0)))
        off=note.find('.//a:xfrm/a:off',N)
        if off is not None: off.set('x',str(round(9.4*914400)))
        replace_text(note,'• UX / зона не указана');tree.append(note)
    return root

def split_text(value,limit):
    """Split at word boundaries, allowing a long token to be split when required."""
    remaining=str(value or '').strip(); pieces=[]
    while remaining:
        cut=min(limit,len(remaining))
        if cut<len(remaining):
            boundary=remaining.rfind(' ',0,cut)
            if boundary>limit//2: cut=boundary
        pieces.append(remaining[:cut].strip())
        remaining=remaining[cut:].strip()
    return pieces or ['']

def product_key(value):
    return re.sub(r'\s+',' ',str(value or '').strip().casefold().replace('ё','е'))


def order_topics_by_product(items):
    """Preserve product order of first appearance, then rank zones stably within product."""
    groups={};rank={'высокое внимание':0,'потенциальный риск':1,'наблюдение':2,'не определена':3,'UX':4}
    for topic in items: groups.setdefault(product_key(topic['product']),[]).append(topic)
    return [topic for product in groups.values() for topic in sorted(product,key=lambda t:rank.get(t.get('zone'),3))]

def overview_rows(items):
    rows=[]
    for item in items:
        pieces=split_text(item['topic'],300)
        for index,part in enumerate(pieces):
            row=dict(item,topic=('… ' if index else '')+part+(' …' if index<len(pieces)-1 else ''))
            row['_product_key']=product_key(item['product'])
            if index:
                row['product']=str(row['product'])+' (прод.)'
                row['duration']=None
                row['extra']={}
                row.pop('status_record',None)
            rows.append(row)
    return rows

def overview_row_height(item):
    lines=max(1,len(textwrap.wrap(str(item.get('topic') or ''),width=65,break_long_words=True)))
    return min(1300000, max(530000, 350000 + lines*180000))

def overview_pages(rows):
    pages=[]; current=[]; used=0
    for row in rows:
        height=overview_row_height(row)
        if current and (len(current)>=4 or used+height>3895000):
            pages.append(current);current=[];used=0
        current.append(row);used+=height
    if current: pages.append(current)
    return pages

def contents_cards(groups):
    # Available space varies between blocks in the supplied table of contents.
    capacities=(6,6,4,2,2,6)
    queue=[]
    rank={'высокое внимание':3,'потенциальный риск':2,'наблюдение':1}
    for owner,items in groups.items():
        byproduct={}
        for t in items:
            name=str(t['product']).strip(); key=product_key(name)
            zone=t.get('zone');zone=zone if zone in rank else None
            if key not in byproduct: byproduct[key]=(name,zone)
            elif rank.get(zone,0)>rank.get(byproduct[key][1],0): byproduct[key]=(byproduct[key][0],zone)
        queue.append([owner,list(byproduct.values())])
    return [[(owner, products) for owner,products in queue[i:i+len(capacities)]]
            for i in range(0,len(queue),len(capacities))]

def set_two_paragraphs(shape,label,value):
    body=shape.find('p:txBody',N);paras=body.findall('a:p',N)
    first=copy.deepcopy(paras[0]);last=copy.deepcopy(paras[-1])
    for run in list(last.findall('a:r',N)):
        if not ''.join(t.text or '' for t in run.findall('a:t',N)).strip(): last.remove(run)
    for para,text in ((first,label),(last,str(value))):
        holder=ET.Element(f'{{{P}}}sp');tmp=ET.SubElement(holder,f'{{{P}}}txBody');tmp.append(para);replace_text(holder,text)
    for para in paras: body.remove(para)
    body.append(first);body.append(last)

def percent(n,d): return '—' if not d else f'{(200*n+d)//(2*d)}%'

def promises_summary(root,report,period):
    tree=shapes(root);counts=report['counts'];due=report['due_count']
    for item in list(tree):
        if item.tag==f'{{{P}}}graphicFrame': tree.remove(item);continue
        if item.tag!=f'{{{P}}}sp': continue
        text=re.sub(r'\s+',' ',content(item)).strip()
        if text.startswith('В октябре были выполнены'): replace_text(item,f"Выполнено {percent(report['completed_count'],due)} обещаний со сроком в месяце" if due else f'{period}: обещаний со сроком в месяце нет')
        elif text=='Октябрь': replace_text(item,period)
        elif text.startswith('Статус не предоставлен'): set_two_paragraphs(item,'Статус не предоставлен',counts['status_missing'])
        elif text.startswith('Перенесено'): set_two_paragraphs(item,'Перенесено',counts['postponed'])
        elif text.startswith('Выполнено досрочно'): set_two_paragraphs(item,'Выполнено досрочно',counts['completed_early'])
        elif text.startswith('Выполнено в срок'): set_two_paragraphs(item,'Выполнено в срок',counts['completed_on_time'])
        elif text.startswith('Должно быть выполнено'): set_two_paragraphs(item,'Должно быть выполнено',due)
        elif text.startswith('Всего обещаний в работе'): set_two_paragraphs(item,'Всего обещаний в работе',report['current_total'])
        elif text.startswith('Всего в октябре перенесли'): replace_text(item,f"Перенесено: {counts['postponed']} из {due}. В работе на конец периода: {report['due_in_work']} из обещаний со сроком в месяце.")
        elif text.startswith('Краткий вывод октября'): replace_text(item,'Краткий вывод')
    # Reuse the original editable heading and body styles for the left panel.
    candidates=[x for x in tree.findall('p:sp',N) if content(x) in ('Краткий вывод',) or content(x).startswith('Перенесено:')]
    for source in candidates:
        clone=copy.deepcopy(source);cnv=clone.find('.//p:cNvPr',N)
        if cnv is not None: cnv.set('id',str(1+max((int(e.get('id','0')) for e in root.findall('.//p:cNvPr',N)),default=0)))
        off=clone.find('.//a:xfrm/a:off',N)
        if off is not None: off.set('x',str(round(.8*914400)))
        replace_text(clone,'Причины переноса' if content(source)=='Краткий вывод' else ('См. комментарии в детализации.' if counts['postponed'] else 'Переносов нет.'))
        tree.append(clone)
    return root

def promise_table(root,rows,index,total,details=False):
    tree=shapes(root)
    for sp in tree.findall('p:sp',N):
        text=content(sp)
        if text=='Выполнение обещаний по ответственным':
            title='Перенесённые обещания' if details else 'Выполнение обещаний по ответственным'
            replace_text(sp,title+(f' ({index}/{total})' if total>1 else ''))
    frame=root.find('.//p:graphicFrame',N);table=frame.find('.//a:tbl',N)
    old=table.findall('a:tr',N);header=old[0];prototype=old[1]
    for r in old[1:]: table.remove(r)
    header_values=['Ответственный','Обещание / источник','Статус','Сроки','Комментарий'] if details else ['Ответственный','В работе на начало','Со сроком в месяце','Выполнено','Перенос срока']
    for c,v in zip(header.findall('a:tc',N),header_values): replace_text(c,v)
    if details:
        widths=[1200000,3200000,1200000,2000000,3462255]
        for col,width in zip(table.findall('a:tblGrid/a:gridCol',N),widths):col.set('w',str(width))
    row_height=1050000 if details else 350000
    for values in rows:
        row=copy.deepcopy(prototype);row.set('h',str(row_height))
        for c,v in zip(row.findall('a:tc',N),values):
            replace_text(c,v);pr=c.find('a:tcPr',N)
            for fill in pr.findall('a:solidFill',N):pr.remove(fill)
            fill=ET.SubElement(pr,f'{{{A}}}solidFill');ET.SubElement(fill,f'{{{A}}}srgbClr',{'val':'FFFFFF'})
            for rp in c.findall('.//a:rPr',N):rp.set('sz','900' if details else '1000')
        table.append(row)
    frame.find('p:xfrm/a:ext',N).set('cy',str(int(header.get('h'))+row_height*len(rows)))
    return root

def promise_detail_rows(report):
    from itertools import zip_longest
    result=[]
    for item in report['postponed']:
        descriptions=split_text(item['id']+' · '+item['description']+' | '+item['source'],180)
        comments=split_text(item.get('comment',''),200)
        for d,c in zip_longest(descriptions,comments,fillvalue=''):
            result.append([item['owner'],d,'Перенос',item['initial_due']+' → '+item['current_due'],c])
    return result

def separator(root,owner,index):
    allsh=root.findall('.//p:sp',N)
    replace_text(allsh[0],f'Продукты и сервисы {owner}')
    replace_text(allsh[1],f'{index:02}')
    return root

def recurrence_code(topic):
    explicit=topic.get('status_record',{}).get('recurrence')
    if explicit: return explicit
    extra={re.sub(r'\s+',' ',str(v).strip().casefold()) for v in topic.get('extra',{}).values()}
    if extra & {'возобновилось','возобновилась','возобновлено','возобновилось обсуждение темы'}: return 'resumed'
    return 'long' if topic.get('duration')==4 else ('two' if topic.get('duration')==2 else None)

def promises_separator(root,index):
    for shape in root.findall('.//p:sp',N):
        if content(shape).strip()=='07': replace_text(shape,f'{index:02}')
    return root

def collect_legend(root):
    legend={}
    for pic in root.findall('.//p:pic',N):
        x,y=position(pic)
        if x is None: continue
        x,y=x/914400,y/914400
        if .5<x<.9 and 6.4<y<6.7: legend['long']=pic
        if .5<x<.9 and 6.8<y<7.1: legend['two']=pic
        if 2.3<x<2.8 and 6.8<y<7.1: legend['resumed']=pic
        if 2.3<x<2.8 and 6.4<y<6.7: legend['new']=pic
        for name,cx,cy in [('aibolit',4.6,6.5),('promise_monitor',4.6,6.9),('pain_candidate',6.0,6.5),('backlog',6.0,6.9),('done',8.1,6.5),('analysis',8.1,6.9),('release',9.4,6.5),('limitations',9.4,6.9),('promnenie',11.1,6.5),('after_close',11.1,6.9)]:
            if abs(x-cx)<.22 and abs(y-cy)<.18: legend[name]=pic
    return legend

def table_slide(root,owner,items,page_no,page_total,ux=False):
    tree=shapes(root)
    for child in list(tree):
        # Icon overlays on source rows refer to the old dataset; retain the legend only.
        if child.tag==f'{{{P}}}pic':
            x,y=position(child)
            if y is not None and y<round(6.0*914400): tree.remove(child)
        if child.tag==f'{{{P}}}sp':
            text=content(child)
            if 'Темы, требующие внимания' in text: replace_text(child,f"{'UX-проблемы' if ux else 'Темы, требующие внимания'} ({page_no}/{page_total})")
            elif text.strip()=='SberUser':
                replace_text(child,owner)
                for rp in child.findall('.//a:rPr',N): rp.set('sz','650' if len(owner)>8 else '800')
    frame=root.find('.//p:graphicFrame',N);table=frame.find('.//a:tbl',N)
    legend=collect_legend(root)
    rows=table.findall('a:tr',N); prototype=rows[1]
    if ux:
        grid=table.find('a:tblGrid',N);cols=grid.findall('a:gridCol',N)
        cols[1].set('w',str(int(cols[1].get('w'))+int(cols[2].get('w'))));grid.remove(cols[2])
        header_cells=rows[0].findall('a:tc',N);rows[0].remove(header_cells[2])
    for row in rows[1:]: table.remove(row)
    heights=[overview_row_height(item) for item in items]
    spans={};start=0
    while start<len(items):
        end=start+1;key=items[start].get('_product_key',product_key(items[start]['product']))
        while end<len(items) and items[end].get('_product_key',product_key(items[end]['product']))==key: end+=1
        spans[start]=end-start
        start=end
    for row_number,t in enumerate(items):
        height=heights[row_number]
        row=copy.deepcopy(prototype);row.set('h',str(height));cells=row.findall('a:tc',N)
        for c in cells: c.attrib.pop('rowSpan',None);c.attrib.pop('vMerge',None)
        status=t.get('status_record',{})
        timing=' '.join(status.get(k,'') for k in ('version','deadline')).strip()
        if ux:
            row.remove(cells[2]);cells=row.findall('a:tc',N)
            values=[t['product'],t['topic'],'','',timing]
        else: values=[t['product'],t['topic'],t['zone'],'','',timing]
        for cell,value in zip(cells,values): replace_text(cell,value)
        product_length=len(str(t['product']))
        if product_length>60:
            size=max(800,min(1200,round(1200*60/product_length)))
            for props in cells[0].findall('.//a:rPr',N): props.set('sz',str(size))
        if row_number in spans and spans[row_number]>1: cells[0].set('rowSpan',str(spans[row_number]))
        elif row_number not in spans:
            cells[0].set('vMerge','1');replace_text(cells[0],'')
        for rp in cells[-1].findall('.//a:rPr',N): rp.set('sz','750')
        if not ux:
            zone=cells[2].find('a:tcPr',N)
            for fill in zone.findall('a:solidFill',N): zone.remove(fill)
            fill=ET.SubElement(zone,f'{{{A}}}solidFill');ET.SubElement(fill,f'{{{A}}}srgbClr',{'val':{'высокое внимание':'FFD9DD','потенциальный риск':'FFE9BF','наблюдение':'DBECFF'}.get(t['zone'],'EEEEEE')})
        table.append(row)
        icon=recurrence_code(t)
        def put_icon(name,x_center,scale=1):
            if name is None: return
            if name not in legend: raise ValueError('В шаблоне не найден необходимый значок: '+name)
            pic=copy.deepcopy(legend[name])
            cnv=pic.find('.//p:cNvPr',N)
            cnv.set('id',str(1+max((int(e.get('id','0')) for e in root.findall('.//p:cNvPr',N)),default=0)))
            cnv.set('descr',ICON_DESCRIPTIONS[name])
            xfrm=pic.find('.//a:xfrm',N);off=xfrm.find('a:off',N);ext=xfrm.find('a:ext',N)
            if scale!=1:
                ext.set('cx',str(int(int(ext.get('cx'))*scale)));ext.set('cy',str(int(int(ext.get('cy'))*scale)))
            off.set('x',str(int(x_center-int(ext.get('cx'))/2)))
            off.set('y',str(809181+430379+sum(heights[:row_number])+(height-int(ext.get('cy')))//2))
            tree.append(pic)
        repeat_start=485888+1246136+5661020+1068117
        put_icon(icon,repeat_start+1246136/2)
        control_start=repeat_start+1246136
        controls=list(status.get('controls',[]))+(['after_close'] if status.get('after_close') else [])
        for i,name in enumerate(controls):
            put_icon(name,control_start+888360*(i+1)/(len(controls)+1), .65 if len(controls)>3 else 1)
        mitigation=status.get('mitigation')
        if mitigation: put_icon(mitigation,control_start+888360+130000)
        if status.get('promnenie'): put_icon('promnenie',485888+1246136-130000)
    return root

def high_detail_slide(root,t,quotes,legend):
    tree=shapes(root)
    source_names=list(t.get('sources', [])) or ['Источник не указан']
    midpoint=(len(source_names)+1)//2
    source_columns=[', '.join(source_names[:midpoint]), ', '.join(source_names[midpoint:])]
    for child in list(tree):
        x,y=position(child)
        if child.tag==f'{{{P}}}pic' and x is not None and y < 6*914400:
            tree.remove(child)
            continue
        if child.tag==f'{{{P}}}grpSp' and x is not None and y>6*914400:
            tree.remove(child)
            continue
        if child.tag==f'{{{P}}}cxnSp' and x is not None and x>8*914400 and y>3*914400:
            tree.remove(child); continue
        if child.tag!=f'{{{P}}}sp': continue
        txt=content(child)
        if not txt:
            if x is not None and x>8*914400 and y>3*914400: tree.remove(child)
            elif x is not None and abs(x-0.4*914400)<100000 and abs(y-1.2*914400)<100000:
                ext=child.find('.//a:xfrm/a:ext',N)
                if ext is not None and int(ext.get('cy'))>914400: ext.set('cy',str(2800000))
            continue
        if txt=='Рабочее место SberOS': product_title(child,t['product'])
        elif 'Источников' in txt: replace_text(child,source_count_label(len(t['sources'])))
        elif 'MAU' in txt: replace_text(child,f"{int(t['mau'])} MAU" if valid_signal(t.get('mau')) else 'MAU: —')
        elif 'Сигнал' in txt: replace_text(child,f"{display_signal_count(t.get('count'), '?')} сигналов")
        elif txt=='SberUser':
            replace_text(child,str(t.get('owner') or ''))
            for rp in child.findall('.//a:rPr',N): rp.set('sz','650' if len(str(t.get('owner') or ''))>8 else '800')
        elif txt=='Источники:': pass
        elif txt=='1': pass
        elif txt.startswith('Карточка ОС') or txt.startswith('чат «Почта для SberOS'):
            source_index=0 if txt.startswith('Карточка ОС') else 1
            if not source_columns[source_index]: tree.remove(child)
            else:
                replace_text(child,source_columns[source_index])
                for rp in child.findall('.//a:rPr',N): rp.set('sz','800')
                transform=child.find('.//a:xfrm',N)
                if transform is not None:
                    transform.find('a:off',N).set('x',str(round((.8 if source_index==0 else 5.1)*914400)))
                    transform.find('a:ext',N).set('cx',str(round(4.1*914400)))
        else: tree.remove(child)
    frame=root.find('.//p:graphicFrame',N)
    table=frame.find('.//a:tbl',N);row=table.findall('a:tr',N)[1];row.set('h','2150000');cells=row.findall('a:tc',N)
    frame.find('p:xfrm/a:ext',N).set('cy',str(2673748))
    for cell in cells: replace_text(cell,'')
    status=t.get('status_record',{})
    requested=[]
    if recurrence_code(t): requested.append(recurrence_code(t))
    requested.extend(status.get('controls',[]))
    if status.get('after_close'): requested.append('after_close')
    if status.get('mitigation'): requested.append(status['mitigation'])
    if status.get('promnenie'): requested.append('promnenie')
    if any(name not in legend for name in requested): raise ValueError('В шаблоне не найдены необходимые значки легенды')
    replace_text(cells[2],'');replace_text(cells[3],'')
    replace_text(cells[4],' '.join(filter(None,[status.get('version',''),status.get('deadline','')])))
    def icon(name,x,y,scale=1):
        pic=copy.deepcopy(legend[name]);cnv=pic.find('.//p:cNvPr',N)
        cnv.set('id',str(1+max((int(e.get('id','0')) for e in root.findall('.//p:cNvPr',N)),default=0)))
        cnv.set('descr',ICON_DESCRIPTIONS[name])
        blip=pic.find('.//a:blip',N);old=blip.get(f'{{{R}}}embed');blip.set(f'{{{R}}}embed','rIdDetailIcon'+old)
        xfrm=pic.find('.//a:xfrm',N);off=xfrm.find('a:off',N);ext=xfrm.find('a:ext',N)
        if scale!=1: ext.set('cx',str(int(int(ext.get('cx'))*scale)));ext.set('cy',str(int(int(ext.get('cy'))*scale)))
        off.set('x',str(int(x-int(ext.get('cx'))/2)));off.set('y',str(int(y-int(ext.get('cy'))/2)));tree.append(pic)
    base=348231+569958+7160404;col=1259368;row_y=1100369+523748+2150000/2
    if recurrence_code(t): icon(recurrence_code(t),base+col/2,row_y)
    controls_icons=list(status.get('controls',[]))+(['after_close'] if status.get('after_close') else [])
    for i,name in enumerate(controls_icons): icon(name,base+col+col*(i+1)/(len(controls_icons)+1),row_y,.65 if len(controls_icons)>3 else 1)
    if status.get('mitigation'): icon(status['mitigation'],base+2*col+140000,row_y)
    if status.get('promnenie'): icon('promnenie',348231+569958/2,1100369+523748+180000)
    body=cells[1].find('a:txBody',N);replace_text(cells[1],t['topic']+' ['+str(t['sheet'])+'!'+str(t['row'])+']')
    for quote in quotes:
        if not quote.get('text'): continue
        paragraph=copy.deepcopy(body.find('a:p',N))
        holder=ET.Element(f'{{{A}}}tc');tmp=ET.SubElement(holder,f'{{{A}}}txBody');tmp.append(paragraph)
        source=str(quote.get('source') or 'Источник не указан')
        replace_text(holder,'«'+quote['text']+'» ['+source+'; '+str(quote['sheet'])+'!'+str(quote['row'])+']')
        for rp in paragraph.findall('.//a:rPr',N): rp.set('b','0');rp.set('sz','1100')
        body.append(paragraph)
    return root

def detail_slide(root,t,quote,number):
    tree=shapes(root)
    for s in list(tree):
        if s.tag!=f'{{{P}}}sp': continue
        txt=content(s); x,y=position(s)
        if txt=='SberESM': product_title(s,t['product'])
        elif txt=='47': replace_text(s,str(number))
        elif txt.startswith('«Хотелось бы'): replace_text(s,quote['text'])
        elif txt.startswith('Неудобный конструктор'): replace_text(s,t['topic'])
        elif txt.startswith('Platform'): replace_text(s,str(quote.get('source') or 'Источник не указан'))
        elif txt.startswith('чат «Новая'): replace_text(s,f"{quote['sheet']}!{quote['row']}")
        elif txt.replace(' ','')=='SberНадежность': replace_text(s,str(t.get('owner') or ''))
        elif txt=='UX' and t['zone']!='UX': replace_text(s,'ТОП')
    return root

def remove_embedded_objects(root):
    """Drop source-template OLE frames; generated slides use native text/tables/images."""
    for parent in list(root.iter()):
        for child in list(parent):
            if child.tag==f'{{{P}}}graphicFrame' and child.find(f'.//{{{P}}}oleObj') is not None:
                parent.remove(child)
    if root.find(f'.//{{{P}}}oleObj') is not None:
        raise ValueError('Не удалось удалить встроенный OLE-объект из шаблона')
    return root


def forbidden_embedded_parts(names):
    prefixes=('ppt/embeddings/','ppt/activex/','ppt/notesslides/','ppt/comments/','customxml/')
    return [name for name in names if name.casefold().startswith(prefixes) or name.casefold()=='ppt/vbaproject.bin']

def rel_path(part):
    parent,name=posixpath.split(part)
    return posixpath.join(parent,'_rels',name+'.rels')

def part_from_target(part,target): return posixpath.normpath(posixpath.join(posixpath.dirname(part),target.lstrip('/'))) if not target.startswith('/') else target.lstrip('/')

def filtered_rels(data, slide=False):
    r=xml(data)
    for e in list(r):
        if e.get('TargetMode')=='External' or (slide and e.get('Type','').endswith(('/notesSlide','/comments','/tags'))): r.remove(e)
    return r

def make_deck(template, output, state):
    with ZipFile(template) as source:
        src={i.filename:source.read(i.filename) for i in source.infolist() if not i.is_dir()}
    for needed in ('ppt/slides/slide1.xml','ppt/slides/slide2.xml','ppt/slides/slide3.xml','ppt/slides/slide4.xml','ppt/slides/slide5.xml','ppt/slides/slide6.xml','ppt/slides/slide12.xml','ppt/slides/slide47.xml'):
        if needed not in src: raise ValueError('Шаблон не содержит ожидаемых слайдов образца: '+needed)
    landmarks={1:'Голос IT',2:'Summary',4:'Методика определения приоритезации',6:'Темы, требующие внимания',12:'Темы с высоким приоритетом',47:'UX'}
    for number,label in landmarks.items():
        slide=xml(src[f'ppt/slides/slide{number}.xml'])
        visible=''.join((x.text or '') for x in slide.findall('.//a:t',N))
        if re.sub(r'\s+','',label).casefold() not in re.sub(r'\s+','',visible).casefold():
            raise ValueError(f'Шаблон не соответствует образцу: слайд {number} не содержит {label!r}')
    legend=collect_legend(xml(src['ppt/slides/slide6.xml']))
    legend_rels=xml(src['ppt/slides/_rels/slide6.xml.rels'])
    legend_targets={r.get('Id'):r.get('Target') for r in legend_rels if r.get('Type','').endswith('/image')}
    groups=defaultdict(list)
    for topic in state['topics']: groups[str(topic.get('owner') or 'Владелец не указан')].append(topic)
    report=state.get('promise_report')
    toc_groups=dict(groups)
    if report: toc_groups[PROMISES_SECTION]=[{'product':'Статистика и переносы'}]
    toc_pages=contents_cards(toc_groups);spans={}
    high_count=len({product_key(t['product']) for t in state['topics'] if t['zone']=='высокое внимание'})
    plans=[(1,lambda r:cover(r,state['period'],bool(report))),(2,lambda r:summary(r,state))]
    for i,cards in enumerate(toc_pages,1):
        plans.append((3,lambda r,cards=cards,i=i:contents_slide(r,cards,spans,i,len(toc_pages),high_count)))
    plans.append((4,lambda r:r))
    for group_index,(owner,items) in enumerate(groups.items(),1):
        group_start=len(plans)+1
        plans.append((5,lambda r,owner=owner,group_index=group_index:separator(r,owner,group_index)))
        for is_ux in (False, True):
            selected=order_topics_by_product([t for t in items if (t['zone']=='UX')==is_ux])
            if not selected: continue
            rows=overview_rows(selected)
            pages=overview_pages(rows)
            for idx,part in enumerate(pages,1):
                plans.append((6,lambda r,owner=owner,part=part,idx=idx,total=len(pages),is_ux=is_ux:table_slide(r,owner,part,idx,total,is_ux)))
            for t in selected:
                if t['zone'] != 'высокое внимание': continue
                # One topic per detail slide; full source rows and texts stay in audit.
                title=str(t['topic'])
                if len(title)>250: title=title[:249].rsplit(' ',1)[0].rstrip()+'…'
                fragment=dict(t,topic=title)
                quotes=t.get('presentation_quotes', t.get('quotes',[])[:2])
                if not quotes: quotes=[{'text':'','source':None,'sheet':t['sheet'],'row':t['row']}]
                plans.append((12,lambda r,t=fragment,quotes=quotes:high_detail_slide(r,t,quotes,legend)))
        spans[owner]=(group_start,len(plans))
    if report:
        for number in (54,55,56):
            if f'ppt/slides/slide{number}.xml' not in src: raise ValueError('Шаблон не содержит слайды мониторинга обещаний')
        group_start=len(plans)+1
        plans.append((54,lambda r:promises_separator(r,len(groups)+1)))
        plans.append((55,lambda r:promises_summary(r,report,state['period'])))
        group_rows=[[g['owner'],str(g['start']),str(g['due']),f"{percent(g['completed'],g['due'])} ({g['completed']})",f"{percent(g['postponed'],g['due'])} ({g['postponed']})"] for g in report['groups']]
        pages=[group_rows[i:i+12] for i in range(0,len(group_rows),12)] or [[]]
        for i,rows in enumerate(pages,1):plans.append((56,lambda r,rows=rows,i=i,n=len(pages):promise_table(r,rows,i,n)))
        detail_rows=promise_detail_rows(report);detail_pages=[detail_rows[i:i+4] for i in range(0,len(detail_rows),4)]
        for i,rows in enumerate(detail_pages,1):plans.append((56,lambda r,rows=rows,i=i,n=len(detail_pages):promise_table(r,rows,i,n,True)))
        spans[PROMISES_SECTION]=(group_start,len(plans))
    # Recreate presentation slide manifest and each slide part, retaining only necessary relationships.
    out={}; pres=xml(src['ppt/presentation.xml']); slideids=pres.find('p:sldIdLst',N);slideids.clear()
    for ex in pres.findall('p:extLst',N): pres.remove(ex)
    prels=xml(src['ppt/_rels/presentation.xml.rels'])
    for e in list(prels):
        if e.get('Type','').endswith('/slide') or e.get('TargetMode')=='External': prels.remove(e)
    used_ids={e.get('Id') for e in prels}
    for i,(template_no,transform) in enumerate(plans,1):
        rid=f'rIdGenerated{i}'
        assert rid not in used_ids
        ET.SubElement(prels,f'{{{PR}}}Relationship',{'Id':rid,'Type':R+'/slide','Target':f'slides/slide{i}.xml'})
        ET.SubElement(slideids,f'{{{P}}}sldId',{'id':str(255+i),f'{{{R}}}id':rid})
        slide=remove_embedded_objects(transform(xml(src[f'ppt/slides/slide{template_no}.xml'])))
        if template_no in (47,54):
            for s in slide.findall('.//p:sp',N):
                if content(s)==('0' if template_no==47 else '54'): replace_text(s,str(i))
        target=f'ppt/slides/slide{i}.xml'; out[target]=serial(slide)
        relsource=f'ppt/slides/_rels/slide{template_no}.xml.rels'
        if relsource in src:
            rels=filtered_rels(src[relsource],True)
            # Transfer references for icons copied from the source legend.
            used={v for el in slide.iter() for k,v in el.attrib.items() if k.startswith('{'+R+'}')}
            for rid in sorted(used):
                if rid.startswith('rIdDetailIcon'):
                    original=rid.removeprefix('rIdDetailIcon')
                    if original not in legend_targets: raise ValueError('Не найдена ссылка на значок в шаблоне')
                    ET.SubElement(rels,f'{{{PR}}}Relationship',{'Id':rid,'Type':R+'/image','Target':legend_targets[original]})
            for e in list(rels):
                if not e.get('Type','').endswith('/slideLayout') and e.get('Id') not in used: rels.remove(e)
            valid={e.get('Id') for e in rels}
            # Discard tags and hyperlinks whose external or source-only relationships were removed.
            for parent in slide.iter():
                for child in list(parent):
                    rid=child.get(f'{{{R}}}id')
                    if rid and rid not in valid and child.tag.split('}')[-1] in ('tags','hlinkClick','hlinkHover'):
                        parent.remove(child)
            out[target]=serial(slide)
            out[rel_path(target)]=serial(rels)
    out['ppt/presentation.xml']=serial(pres);out['ppt/_rels/presentation.xml.rels']=serial(prels)
    # Traverse relationships from presentation only. Old slides and old notes are never copied.
    q=deque(['ppt/presentation.xml']);seen=set()
    while q:
        part=q.popleft()
        if part in seen: continue
        seen.add(part)
        if part not in out:
            if part not in src: raise ValueError(f'В шаблоне отсутствует {part}')
            out[part]=src[part]
        rp=rel_path(part)
        if rp not in out and rp in src:
            rels=filtered_rels(src[rp]);
            # Masters may link to every old slide layout, which is fine; remove any slide or notes reference.
            for e in list(rels):
                if e.get('Type','').endswith(('/slide','/notesSlide','/comments','/tags')): rels.remove(e)
            out[rp]=serial(rels)
        if rp in out:
            for rel in xml(out[rp]):
                if rel.get('TargetMode')!='External': q.append(part_from_target(part,rel.get('Target','')))
    forbidden=forbidden_embedded_parts(out)
    if forbidden: raise ValueError('В шаблоне обнаружены неподдерживаемые встроенные части; сборка остановлена')
    # Minimal package-level relationship; no document properties or thumbnails from source.
    rel=ET.Element(f'{{{PR}}}Relationships');ET.SubElement(rel,f'{{{PR}}}Relationship',{'Id':'rId1','Type':R+'/officeDocument','Target':'ppt/presentation.xml'});out['_rels/.rels']=serial(rel)
    types=xml(src['[Content_Types].xml'])
    for e in list(types):
        pn=e.get('PartName','').lstrip('/')
        if pn and pn not in out: types.remove(e)
    existing={e.get('PartName') for e in types}
    for i in range(1,len(plans)+1):
        name=f'/ppt/slides/slide{i}.xml'
        if name not in existing: ET.SubElement(types,f'{{{CT}}}Override',{'PartName':name,'ContentType':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'})
    out['[Content_Types].xml']=serial(types)
    with ZipFile(output,'w',ZIP_DEFLATED) as dest:
        for name,data in out.items(): dest.writestr(name,data)
    return len(plans)

class TemplatePresentationAgent:
    def __call__(self,state):
        output=Path(state['output']).resolve();output.parent.mkdir(parents=True,exist_ok=True)
        count=make_deck(state['template'],output,state)
        return {'output':str(output),'slide_count':count}
