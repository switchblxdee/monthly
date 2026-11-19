from collections import defaultdict
from html import escape
from pathlib import Path
import subprocess
import tempfile
import textwrap
from .agents import display_signal_count

NS='xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"'

def frame(x,y,w,h,content,style='body'):
    text=''.join(f'<text:p text:style-name="{style}">{escape(str(line))}</text:p>' for line in content)
    return f'<draw:frame draw:style-name="frame" svg:x="{x}cm" svg:y="{y}cm" svg:width="{w}cm" svg:height="{h}cm"><draw:text-box>{text}</draw:text-box></draw:frame>'

def page(title, lines, subtitle=''):
    return '<draw:page draw:style-name="page" draw:master-page-name="Default">'+frame(1,0.4,30,1.5,[title],'title')+frame(1,2.0,30,0.8,[subtitle],'sub')+frame(1,3,30,14.2,lines)+'</draw:page>'

def chunks(xs,n):
    for i in range(0,len(xs),n): yield xs[i:i+n]


def split_preserving_text(value,limit):
    text=str(value or '').strip();parts=[]
    while text:
        cut=min(limit,len(text))
        if cut<len(text):
            boundary=text.rfind(' ',0,cut)
            if boundary>limit//2: cut=boundary
        parts.append(text[:cut].strip());text=text[cut:].strip()
    return parts or ['']


def wrap_paragraphs(lines,width=88):
    wrapped=[]
    for value in lines:
        for paragraph in str(value).split('\n'):
            wrapped.extend(textwrap.wrap(paragraph,width=width,break_long_words=True,break_on_hyphens=False) or [''])
    return wrapped


def build_pages(state):
    pages=[page('Голос IT',[state['period'],'Автоматически подготовленный выпуск — требуется проверка'],'Мониторинг обратной связи')]
    topics=state['topics'];total=len(topics)
    zones={z:sum(t['zone']==z for t in topics) for z in ['высокое внимание','потенциальный риск','наблюдение','UX','не определена']}
    summary=[f'Тем в листе приоритетов: {total}',*[f'{k}: {v}' for k,v in zones.items()],state.get('narrative','')]
    for part in chunks(wrap_paragraphs(summary),20): pages.append(page('Summary',part,'Показатели рассчитаны по предоставленному листу'))
    groups=defaultdict(list)
    for t in topics: groups[str(t.get('owner') or 'Владелец не указан')].append(t)
    for owner,items in groups.items():
        pages.append(page(owner,['Темы, требующие внимания']))
        lines=[]
        for t in items:
            lines.extend([f"{t['product']}: {t['zone']}; {t['recurrence']}; сигналов: {display_signal_count(t.get('count'))}",str(t['topic']),''])
        parts=list(chunks(wrap_paragraphs(lines),20))
        for idx,part in enumerate(parts,1): pages.append(page('Темы, требующие внимания',part,f'{owner} · {idx}/{len(parts)}'))
        for t in items:
            if t['zone'] not in ('высокое внимание','UX'): continue
            topic=str(t['topic']);short=topic if len(topic)<=200 else topic[:199]+'…'
            quotes=t['quotes'] or [{'text':'','source':None,'sheet':t['sheet'],'row':t['row']}]
            for q in quotes:
                pieces=split_preserving_text(q['text'],700)
                for index,piece in enumerate(pieces,1):
                    lines=[f'Тема: {short}',f"Приоритет: {t['zone']}; сигналов: {display_signal_count(t.get('count'))}",'']
                    lines.append(f'«{piece}»' if piece else 'Подходящее обращение не найдено')
                    if q['text']: lines.append(f"Источник: {q.get('source') or 'не указан'} ({q['sheet']}!{q['row']})")
                    suffix=f' · фрагмент {index}/{len(pieces)}' if len(pieces)>1 else ''
                    for part in chunks(wrap_paragraphs(lines),20): pages.append(page(str(t['product']),part,owner+suffix))
    return pages

class PresentationAgent:
    def __call__(self,state):
        pages=build_pages(state)
        styles='''<office:automatic-styles><style:style style:name="page" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#F3F5FB"/></style:style><style:style style:name="frame" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="none" fo:padding="0cm"/></style:style><style:style style:name="title" style:family="paragraph"><style:text-properties fo:font-family="Liberation Serif" fo:font-size="25pt" fo:font-weight="bold" fo:color="#17243B"/></style:style><style:style style:name="sub" style:family="paragraph"><style:text-properties fo:font-family="Liberation Sans" fo:font-size="11pt" fo:color="#397DDB"/></style:style><style:style style:name="body" style:family="paragraph"><style:paragraph-properties fo:margin-bottom="0.10cm"/><style:text-properties fo:font-family="Liberation Sans" fo:font-size="12pt" fo:color="#17243B"/></style:style><style:page-layout style:name="layout"><style:page-layout-properties fo:page-width="33.87cm" fo:page-height="19.05cm" style:print-orientation="landscape"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="layout"/></office:master-styles>'''
        doc=f'<?xml version="1.0" encoding="UTF-8"?><office:document {NS} office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.presentation">{styles}<office:body><office:presentation>{"".join(pages)}</office:presentation></office:body></office:document>'
        out=Path(state['output']).resolve(); out.parent.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='golos-') as d:
            source=Path(d)/'presentation.fodp'; source.write_text(doc,encoding='utf-8')
            profile=Path(d)/'profile'
            cmd=['soffice',f'-env:UserInstallation={profile.as_uri()}','--headless','--convert-to','pptx','--outdir',d,str(source)]
            try:
                result=subprocess.run(cmd,capture_output=True,text=True,timeout=120)
            except FileNotFoundError as exc:
                raise RuntimeError('Для сборки без шаблона нужен LibreOffice (soffice в PATH); либо укажите --template') from exc
            generated=Path(d)/'presentation.pptx'
            if result.returncode or not generated.exists(): raise RuntimeError(f'LibreOffice: {result.stderr} {result.stdout}')
            out.write_bytes(generated.read_bytes())
        return {'output':str(out), 'slide_count':len(pages)}
