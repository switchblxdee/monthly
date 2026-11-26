"""Product detail pages: multiple topics, source widgets and template metrics."""
import copy
import math
import xml.etree.ElementTree as ET


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
    return pages


SOURCE_HEIGHT = 252000
EMU_PER_INCH = 914400


def source_top(topics):
    if not topics:
        return 2.25
    return max(3.30, (1100369 + 523748 + sum(map(topic_height, topics))) / EMU_PER_INCH + .47)


def source_widget_layout(sources, top=4.92):
    """Pack single-line, seven-millimetre badges without hiding source names."""
    names = list(sources)
    x = .76
    y = top
    height = SOURCE_HEIGHT / EMU_PER_INCH
    layout = []
    for name in names:
        glyph_width = sum(1.0 if c in 'МШЩЖЮW@' else .7 if c.isupper() else .56 for c in name)
        width = min(8.11, max(.85, .42 + glyph_width * 11 / 72))
        if x + width > 8.87 + 1e-6:
            x = .76
            y += height + .045
        if y + height > 7.42:
            break
        layout.append((name, x, y, width))
        x += width + .09
    return layout, len(names) - len(layout)


def detail_source_pages(topics, sources):
    pages = detail_pages(topics)
    pending = list(sources) or ['Источник не указан']
    layout, remaining = source_widget_layout(pending, source_top(pages[-1]))
    result = [(page, []) for page in pages[:-1]] + [(pages[-1], layout)]
    while remaining:
        pending = pending[-remaining:]
        layout, remaining = source_widget_layout(pending, source_top([]))
        result.append(([], layout))
    return result


def render_detail(root, topics, legend, sources, last, product_topics, number_offset=0, source_layout=()):
    from .template_render import A, P, R, N, shapes, position, content, replace_text, product_title, source_count_label, recurrence_code, ICON_DESCRIPTIONS, remove_bottom_border
    from .agents import valid_signal, display_signal_count
    from .product_metrics import product_metrics, format_metric, parse_metric

    tree=shapes(root); first=product_topics[0]
    show_sources = bool(source_layout)
    footer_top = source_top(topics)
    markers = []
    header_flag = None
    frame=root.find('.//p:graphicFrame',N);table=frame.find('.//a:tbl',N)
    old=table.findall('a:tr',N); header=old[0]; prototype=old[1]
    for row in old[1:]: table.remove(row)
    heights=[topic_height(t) for t in topics]
    table_height=int(header.get('h'))+sum(heights)
    frame.find('p:xfrm/a:ext',N).set('cy',str(table_height))
    if not topics: tree.remove(frame)
    source_template=None;number_template=None
    values=[t.get('mau') for t in product_topics if valid_signal(t.get('mau'))]
    mau=values[0] if values and len(values)==len(product_topics) and len(set(values))==1 else None
    counts=[t.get('count') for t in product_topics]
    metrics=product_metrics(product_topics)
    signal_total=sum(int(x) for x in counts) if counts and all(valid_signal(x) for x in counts) else None
    for child in list(tree):
        x,y=position(child); tag=child.tag
        if tag==f'{{{P}}}pic' and x is not None and x < 914400 and 1.25*914400 < y < 1.65*914400 and topics:
            header_flag = child
            continue
        if tag==f'{{{P}}}pic' and x is not None and (y<6*914400 or (not show_sources and x<914400)):
            tree.remove(child);continue
        if tag==f'{{{P}}}pic' and show_sources and x is not None and x<914400 and y>6*914400:
            off=child.find('.//a:xfrm/a:off',N)
            if off is not None: off.set('y',str(round((footer_top-.29)*914400)))
        if tag==f'{{{P}}}grpSp':
            if content(child).strip()!='+' or not last: tree.remove(child)
            continue
        if tag==f'{{{P}}}cxnSp':
            if not show_sources and x is not None and x<914400 and y>6*914400: tree.remove(child)
            elif show_sources and x is not None and x<914400 and y>6*914400:
                off=child.find('.//a:xfrm/a:off',N)
                if off is not None: off.set('y',str(round((footer_top-.06)*914400)))
            elif x is not None and x>8*914400 and y>3*914400: tree.remove(child)
            continue
        if tag!=f'{{{P}}}sp': continue
        txt=content(child).strip()
        # The score and delta boxes occupy fixed positions in slide 12.
        # The contents of those boxes in the template can change each month.
        geometry = child.find('p:spPr/a:prstGeom', N)
        is_triangle = geometry is not None and geometry.get('prst') == 'triangle'
        if not is_triangle and child.find('p:txBody', N) is not None and x is not None and 6.55 < y/914400 < 7.13 and 11.45 < x/914400 < 13:
            if last:
                signed = x/914400 >= 12.30
                opinion = y/914400 < 6.80
                field = ('opinion_delta' if signed else 'opinion_score') if opinion else ('csat_delta' if signed else 'csat')
                value = format_metric(metrics[field], signed=signed)
                replace_text(child, value)
                tr = child.find('.//a:xfrm', N)
                tr.find('a:off', N).set('x', str(round((12.31 if signed else 11.24)*914400)))
                tr.find('a:ext', N).set('cx', str(round((.62 if signed else .83)*914400)))
                if signed:
                    values = [parse_metric(n) for n in metrics[field]['observed_values']]
                    if values and (all(n > 0 for n in values) or all(n < 0 for n in values)):
                        markers.append((y, values[0] > 0))
                body = child.find('p:txBody/a:bodyPr', N)
                if body is not None:
                    body.set('wrap', 'none')
                for props in child.findall('.//a:rPr', N):
                    props.set('sz', str(max(650, min(1250 if not signed else 1050, 9200//max(1,len(value))))))
            else:
                tree.remove(child)
            continue
        if not txt:
            if x is not None and abs(x-.4*914400)<100000 and abs(y-1.2*914400)<100000:
                if not topics:
                    tree.remove(child)
                    continue
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
            if not show_sources: tree.remove(child)
            else:
                off=child.find('.//a:xfrm/a:off',N)
                if off is not None: off.set('y',str(round((footer_top-.31)*914400)))
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
        else: tree.remove(child)

    # Line up the right-hand footer with the source block on short detail pages.
    footer_shift = round(max(0, 5.35 - footer_top) * 914400)
    if last and footer_shift:
        for shape in tree:
            x, y = position(shape)
            if x is None or x < 8.90*914400 or y < 4.85*914400:
                continue
            transform = shape.find('p:spPr/a:xfrm/a:off', N)
            if transform is None:
                transform = shape.find('p:grpSpPr/a:xfrm/a:off', N)
            if transform is not None:
                transform.set('y', str(int(transform.get('y')) - footer_shift))
    for badge in tree.findall('p:sp', N):
        x, y = position(badge)
        if x is None or not (8*914400 < x and .60*914400 < y < 914400):
            continue
        for prop in badge.findall('.//a:rPr', N):
            prop.set('sz', '1200')
        body = badge.find('p:txBody/a:bodyPr', N)
        if body is not None:
            body.attrib.update(wrap='none', anchor='ctr', lIns='36000', rIns='36000', tIns='0', bIns='0')
        for paragraph in badge.findall('.//a:pPr', N):
            paragraph.set('algn', 'ctr')
            for node in list(paragraph):
                if node.tag.endswith(('}spcBef', '}spcAft')):
                    paragraph.remove(node)
        line = badge.find('p:spPr/a:ln', N)
        if line is not None:
            for dash in list(line):
                if dash.tag.endswith(('}prstDash', '}custDash')):
                    line.remove(dash)
    if header_flag is not None:
        # Leave the flag above the table so Office does not cover it with the header fill.
        tree.remove(header_flag)
        tree.append(header_flag)

    def unique_id(element):
        cnv=element.find('.//p:cNvPr',N)
        if cnv is not None: cnv.set('id',str(1+max((int(e.get('id','0')) for e in root.findall('.//p:cNvPr',N)),default=0)))
    for y, positive in markers:
        marker = ET.Element(f'{{{P}}}sp')
        nonvisual = ET.SubElement(marker, f'{{{P}}}nvSpPr')
        ET.SubElement(nonvisual, f'{{{P}}}cNvPr', {'name': 'Рост' if positive else 'Снижение'})
        ET.SubElement(nonvisual, f'{{{P}}}cNvSpPr')
        ET.SubElement(nonvisual, f'{{{P}}}nvPr')
        props = ET.SubElement(marker, f'{{{P}}}spPr')
        transform = ET.SubElement(props, f'{{{A}}}xfrm', {'rot': '0' if positive else '10800000'})
        ET.SubElement(transform, f'{{{A}}}off', {'x': str(round(12.13*914400)), 'y': str(y+41000-footer_shift)})
        ET.SubElement(transform, f'{{{A}}}ext', {'cx': '108000', 'cy': '108000'})
        geom = ET.SubElement(props, f'{{{A}}}prstGeom', {'prst': 'triangle'})
        ET.SubElement(geom, f'{{{A}}}avLst')
        fill = ET.SubElement(props, f'{{{A}}}solidFill')
        ET.SubElement(fill, f'{{{A}}}srgbClr', {'val': '28A745' if positive else 'E74755'})
        line = ET.SubElement(props, f'{{{A}}}ln')
        ET.SubElement(line, f'{{{A}}}noFill')
        unique_id(marker)
        tree.append(marker)

    if show_sources and source_template is not None:
        for name, x, y, width in source_layout:
            shape=copy.deepcopy(source_template);unique_id(shape);replace_text(shape,name)
            tr=shape.find('.//a:xfrm',N);off=tr.find('a:off',N);ext=tr.find('a:ext',N)
            off.set('x',str(round(x*914400)));off.set('y',str(round(y*914400)))
            ext.set('cx',str(round(width*914400)));ext.set('cy',str(SOURCE_HEIGHT))
            body=shape.find('p:txBody/a:bodyPr',N)
            if body is not None:
                body.attrib.update(wrap='none', anchor='ctr', tIns='0', bIns='0', lIns='115000', rIns='115000')
                for fit in list(body):
                    if fit.tag in (f'{{{A}}}spAutoFit',f'{{{A}}}normAutofit',f'{{{A}}}noAutofit'): body.remove(fit)
                ET.SubElement(body,f'{{{A}}}noAutofit')
            for paragraph in shape.findall('.//a:pPr', N):
                for node in list(paragraph):
                    if node.tag.endswith(('}lnSpc','}spcBef','}spcAft')): paragraph.remove(node)
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
