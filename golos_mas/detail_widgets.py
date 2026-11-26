"""Product detail pages: multiple topics, source widgets and template metrics."""
import copy
import math
import xml.etree.ElementTree as ET
import textwrap


def topic_height(topic):
    quotes=topic.get('presentation_quotes', topic.get('quotes', [])[:2])
    chars=len(str(topic.get('topic') or '')[:250])+sum(len(str(q.get('text') or '')) for q in quotes)
    return min(1800000, 370000 + math.ceil(max(chars, 1)/90)*175000)


def detail_pages(topics):
    pages=[];page=[];used=0
    for topic in topics:
        h=topic_height(topic)
        if page and used+h>3100000:
            pages.append(page);page=[];used=0
        page.append(topic);used+=h
    if page: pages.append(page)
    if pages and len(pages[-1])>1 and sum(topic_height(t) for t in pages[-1])>2250000:
        final=pages[-1].pop();pages.append([final])
    return pages


def source_widget_layout(sources):
    """Fit complete rows in the fixed footer, reserving space for overflow text."""
    names = list(sources) or ['Источник не указан']
    layout = []
    top = 4.92
    for start in range(0, len(names), 3):
        group = names[start:start+3]
        lines = max(len(textwrap.wrap(name, width=26)) or 1 for name in group)
        height = .24 * lines + .20
        more = start + len(group) < len(names)
        if top + height > (6.93 if more else 7.42):
            layout.append((f'Ещё {len(names)-start} · в аудите', .76, top, .44))
            return layout, len(names)-start
        layout.extend((name, .76 + 2.73*col, top, height) for col, name in enumerate(group))
        top += height + .04
    return layout, 0


def render_detail(root, topics, legend, sources, last, product_topics, number_offset=0):
    from .template_render import A, P, R, N, shapes, position, content, replace_text, product_title, source_count_label, recurrence_code, ICON_DESCRIPTIONS, remove_bottom_border
    from .agents import valid_signal, display_signal_count
    from .product_metrics import product_metrics, format_metric

    tree=shapes(root); first=topics[0]
    frame=root.find('.//p:graphicFrame',N);table=frame.find('.//a:tbl',N)
    old=table.findall('a:tr',N); header=old[0]; prototype=old[1]
    for row in old[1:]: table.remove(row)
    heights=[topic_height(t) for t in topics]
    table_height=int(header.get('h'))+sum(heights)
    frame.find('p:xfrm/a:ext',N).set('cy',str(table_height))
    source_template=None;number_template=None
    values=[t.get('mau') for t in product_topics if valid_signal(t.get('mau'))]
    mau=values[0] if values and len(values)==len(product_topics) and len(set(values))==1 else None
    counts=[t.get('count') for t in product_topics]
    metrics=product_metrics(product_topics)
    signal_total=sum(int(x) for x in counts) if counts and all(valid_signal(x) for x in counts) else None
    for child in list(tree):
        x,y=position(child); tag=child.tag
        if tag==f'{{{P}}}pic' and x is not None and (y<6*914400 or (not last and x<914400)):
            tree.remove(child);continue
        if tag==f'{{{P}}}pic' and last and x is not None and x<914400 and y>6*914400:
            off=child.find('.//a:xfrm/a:off',N)
            if off is not None: off.set('y',str(round(4.63*914400)))
        if tag==f'{{{P}}}grpSp':
            if content(child).strip()!='+' or not last: tree.remove(child)
            continue
        if tag==f'{{{P}}}cxnSp':
            if not last and x is not None and x<914400 and y>6*914400: tree.remove(child)
            elif last and x is not None and x<914400 and y>6*914400:
                off=child.find('.//a:xfrm/a:off',N)
                if off is not None: off.set('y',str(round(4.88*914400)))
            elif x is not None and x>8*914400 and y>3*914400: tree.remove(child)
            continue
        if tag!=f'{{{P}}}sp': continue
        txt=content(child).strip()
        if not txt:
            if x is not None and abs(x-.4*914400)<100000 and abs(y-1.2*914400)<100000:
                transform=child.find('.//a:xfrm',N)
                ext=transform.find('a:ext',N) if transform is not None else None
                if ext is not None and int(ext.get('cy'))>914400:
                    off=transform.find('a:off',N); frame_transform=frame.find('p:xfrm',N)
                    frame_off=frame_transform.find('a:off',N);frame_ext=frame_transform.find('a:ext',N)
                    off.set('x',frame_off.get('x'));off.set('y',frame_off.get('y'))
                    ext.set('cx',frame_ext.get('cx'));ext.set('cy',str(table_height))
            elif x is not None and x>8*914400 and y>3*914400:
                if not (last and abs(x-10.65*914400)<50000 and abs(y-4.92*914400)<50000):
                    tree.remove(child)
        elif txt=='Рабочее место SberOS': product_title(child,first['product'])
        elif 'Источников' in txt:
            if last: replace_text(child,source_count_label(len(sources)))
            else: tree.remove(child)
        elif 'MAU' in txt: replace_text(child,f'{int(mau)} MAU' if mau is not None else 'MAU: —')
        elif 'Сигнал' in txt:
            replace_text(child,f"{display_signal_count(signal_total, '?')} Сигналов")
            for props in child.findall('.//a:rPr',N): props.set('sz','850')
        elif txt=='SberUser':
            owner = str(first.get('owner') or '')
            replace_text(child, owner)
            for props in child.findall('.//a:rPr', N):
                props.set('sz', str(max(600, min(800, 7200 // max(1, len(owner))))))
            body = child.find('p:txBody/a:bodyPr', N)
            if body is not None: body.set('wrap', 'none')
        elif txt=='1': number_template=copy.deepcopy(child);tree.remove(child)
        elif txt.startswith('Карточка ОС') or txt.startswith('чат «Почта для SberOS') or txt.startswith('чат «SberOS') or txt.startswith('Открытый диалог'):
            if source_template is None: source_template=copy.deepcopy(child)
            tree.remove(child)
        elif txt=='Источники:':
            if not last: tree.remove(child)
            else:
                off=child.find('.//a:xfrm/a:off',N)
                if off is not None: off.set('y',str(round(4.60*914400)))
        elif 'приложении к дайджесту' in txt:
            if not last: tree.remove(child);continue
            transform=child.find('.//a:xfrm',N)
            if transform is not None: transform.find('a:off',N).set('y',str(round(5.35*914400)))
            for props in child.findall('.//a:rPr',N):
                fills=props.findall('a:solidFill',N)
                index=list(props).index(fills[0]) if fills else 0
                for old_fill in fills: props.remove(old_fill)
                fill=ET.Element(f'{{{A}}}solidFill');ET.SubElement(fill,f'{{{A}}}srgbClr',{'val':'334A68'})
                props.insert(index,fill)
        elif txt in ('Дополнительно:', 'ПроМнение', 'CSAT'):
            if not last: tree.remove(child)
        elif txt in ('4,28','4,4'):
            if not last: tree.remove(child)
            else:
                replace_text(child,format_metric(metrics['opinion_score' if txt=='4,28' else 'csat']))
                tr=child.find('.//a:xfrm',N)
                if tr is not None:
                    tr.find('a:off',N).set('x',str(round(11.65*914400)))
                    tr.find('a:ext',N).set('cx',str(round(.62*914400)))
                body=child.find('p:txBody/a:bodyPr',N)
                if body is not None: body.set('wrap','none')
        elif txt.replace(' ','') in ('+0,48','-0,1'):
            if not last: tree.remove(child)
            else:
                replace_text(child,format_metric(metrics['opinion_delta' if '+' in txt else 'csat_delta'],signed=True))
                tr=child.find('.//a:xfrm',N)
                if tr is not None:
                    tr.find('a:off',N).set('x',str(round(12.32*914400)))
                    tr.find('a:ext',N).set('cx',str(round(.62*914400)))
                body=child.find('p:txBody/a:bodyPr',N)
                if body is not None: body.set('wrap','none')
        else: tree.remove(child)

    def unique_id(element):
        cnv=element.find('.//p:cNvPr',N)
        if cnv is not None: cnv.set('id',str(1+max((int(e.get('id','0')) for e in root.findall('.//p:cNvPr',N)),default=0)))
    if last and source_template is not None:
        layout, _ = source_widget_layout(sources)
        for name, x, y, h in layout:
            shape=copy.deepcopy(source_template);unique_id(shape);replace_text(shape,name)
            tr=shape.find('.//a:xfrm',N);off=tr.find('a:off',N);ext=tr.find('a:ext',N)
            off.set('x',str(round(x*914400)));off.set('y',str(round(y*914400)))
            ext.set('cx',str(round(2.60*914400)));ext.set('cy',str(round(h*914400)))
            body=shape.find('p:txBody/a:bodyPr',N)
            if body is not None:
                body.set('wrap','square')
                body.set('tIns','18000')
                body.set('bIns','18000')
                body.set('anchor','ctr')
                for fit in list(body):
                    if fit.tag in (f'{{{A}}}spAutoFit',f'{{{A}}}normAutofit'): body.remove(fit)
                ET.SubElement(body,f'{{{A}}}noAutofit')
            for paragraph in shape.findall('.//a:pPr', N):
                for tag in ('lnSpc', 'spcBef', 'spcAft'):
                    node = paragraph.find('a:' + tag, N)
                    if node is not None: paragraph.remove(node)
                spacing = ET.Element(f'{{{A}}}lnSpc')
                ET.SubElement(spacing, f'{{{A}}}spcPts', {'val': '1250'})
                paragraph.insert(0, spacing)
            for rp in shape.findall('.//a:rPr',N):
                rp.set('sz','1100');rp.set('b','0')
                latin=rp.find('a:latin',N)
                if latin is None: latin=ET.SubElement(rp,f'{{{A}}}latin')
                latin.set('typeface','SB Sans Display Light')
            tree.append(shape)

    def icon(name,x,y,scale=1):
        if not name: return
        if name not in legend: raise ValueError('В шаблоне не найден значок: '+name)
        pic=copy.deepcopy(legend[name]);unique_id(pic)
        pic.find('.//p:cNvPr',N).set('descr',ICON_DESCRIPTIONS[name])
        blip=pic.find('.//a:blip',N);blip.set(f'{{{R}}}embed','rIdDetailIcon'+blip.get(f'{{{R}}}embed'))
        xfrm=pic.find('.//a:xfrm',N);off=xfrm.find('a:off',N);ext=xfrm.find('a:ext',N)
        if scale!=1: ext.set('cx',str(int(int(ext.get('cx'))*scale)));ext.set('cy',str(int(int(ext.get('cy'))*scale)))
        off.set('x',str(int(x-int(ext.get('cx'))/2)));off.set('y',str(int(y-int(ext.get('cy'))/2)));tree.append(pic)

    for index,(topic,height) in enumerate(zip(topics,heights)):
        row=copy.deepcopy(prototype);row.set('h',str(height));cells=row.findall('a:tc',N)
        for cell in cells:
            replace_text(cell,'')
            props=cell.find('a:tcPr',N)
            if props is None: props=ET.SubElement(cell,f'{{{A}}}tcPr')
            for fill in props.findall('a:solidFill',N): props.remove(fill)
            fill=ET.SubElement(props,f'{{{A}}}solidFill');ET.SubElement(fill,f'{{{A}}}srgbClr',{'val':'FFFFFF'})
            if index==len(topics)-1: remove_bottom_border(cell)
        title=str(topic['topic']);title=title if len(title)<=250 else title[:249].rsplit(' ',1)[0].rstrip()+'…'
        replace_text(cells[1],title)
        title_properties=cells[1].findall('.//a:rPr',N)
        for rp in title_properties:
            rp.set('sz','1200');rp.set('b','1')
            latin=rp.find('a:latin',N)
            if latin is None: latin=ET.SubElement(rp,f'{{{A}}}latin')
            latin.set('typeface','SB Sans Display Semibold')
        status=topic.get('status_record',{})
        replace_text(cells[4],' '.join(filter(None,[status.get('version',''),status.get('deadline','')])))
        body=cells[1].find('a:txBody',N)
        for quote in topic.get('presentation_quotes',topic.get('quotes',[])[:2]):
            if not quote.get('text'): continue
            para=copy.deepcopy(body.find('a:p',N));holder=ET.Element(f'{{{A}}}tc');tmp=ET.SubElement(holder,f'{{{A}}}txBody');tmp.append(para)
            replace_text(holder,'«'+quote['text']+'»')
            for rp in para.findall('.//a:rPr',N):
                rp.set('b','0');rp.set('sz','1100')
                latin=rp.find('a:latin',N)
                if latin is None: latin=ET.SubElement(rp,f'{{{A}}}latin')
                latin.set('typeface','SB Sans Display Light')
            body.append(para)
        if not topic.get('presentation_quotes',topic.get('quotes',[])[:2]):
            para=copy.deepcopy(body.find('a:p',N));holder=ET.Element(f'{{{A}}}tc');tmp=ET.SubElement(holder,f'{{{A}}}txBody');tmp.append(para)
            replace_text(holder,'Комментарий клиента не найден')
            for rp in para.findall('.//a:rPr',N): rp.set('b','0');rp.set('sz','1100')
            body.append(para)
        table.append(row)
        row_y=1100369+int(header.get('h'))+sum(heights[:index])+height/2
        if number_template is not None:
            number=copy.deepcopy(number_template);unique_id(number);replace_text(number,str(number_offset+index+1))
            xfrm=number.find('.//a:xfrm',N);off=xfrm.find('a:off',N);ext=xfrm.find('a:ext',N)
            off.set('y',str(int(row_y-int(ext.get('cy'))/2)));tree.append(number)
        base=348231+569958+7160404;col=1259368
        icon(recurrence_code(topic),base+col/2,row_y)
        controls=list(status.get('controls',[]))+(['after_close'] if status.get('after_close') else [])
        for i,name in enumerate(controls): icon(name,base+col+col*(i+1)/(len(controls)+1),row_y,.65 if len(controls)>3 else 1)
        icon(status.get('mitigation'),base+2*col+140000,row_y)
        if status.get('promnenie'): icon('promnenie',348231+569958/2,row_y)
    return root
