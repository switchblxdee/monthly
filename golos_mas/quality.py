"""Checks that do not require sending local data to a model."""
from pathlib import Path
import posixpath
from zipfile import ZipFile, BadZipFile
import xml.etree.ElementTree as ET
from .template_render import P, A, R, rel_path, part_from_target, forbidden_embedded_parts, product_key

class QualityAgent:
    def __call__(self,state):
        try:
            return self._inspect(state)
        except BaseException:
            if state.get('staging_target'): Path(state['output']).unlink(missing_ok=True)
            raise

    def _inspect(self,state):
        warnings=list(state.get('warnings',[])); path=Path(state['output'])
        try:
            with ZipFile(path) as deck:
                if deck.testzip(): raise ValueError('PPTX содержит повреждённый файл')
                names=set(deck.namelist())
                if forbidden_embedded_parts(names): raise ValueError('PPTX содержит запрещённые встроенные части')
                for rel_file in (n for n in names if n.endswith('.rels')):
                    if rel_file=='_rels/.rels': owner=''
                    else:
                        parent,filename=posixpath.split(rel_file)
                        if not parent.endswith('/_rels'): continue
                        owner=posixpath.join(parent[:-6],filename[:-5])
                    relationships=list(ET.fromstring(deck.read(rel_file)))
                    ids=[r.get('Id') for r in relationships]
                    if len(ids)!=len(set(ids)): raise ValueError(f'{rel_file}: повтор идентификатора связи')
                    for relation in relationships:
                        if relation.get('TargetMode')=='External': continue
                        target=part_from_target(owner,relation.get('Target',''))
                        if target not in names: raise ValueError(f'{rel_file}: отсутствует связанная часть {target}')
                pres=ET.fromstring(deck.read('ppt/presentation.xml'))
                actual=len(pres.findall(f'.//{{{P}}}sldId'))
                if actual != state['slide_count']: raise ValueError(f'Ожидалось {state["slide_count"]} слайдов, найдено {actual}')
                for i in range(1,actual+1):
                    slide=f'ppt/slides/slide{i}.xml'
                    if slide not in names: raise ValueError(f'Не найден {slide}')
                    root=ET.fromstring(deck.read(slide)); rel=rel_path(slide)
                    for para in root.iter(f'{{{A}}}p'):
                        children=list(para)
                        ends=[i for i,child in enumerate(children) if child.tag==f'{{{A}}}endParaRPr']
                        if ends and (len(ends)>1 or ends[0]!=len(children)-1):
                            raise ValueError(f'{slide}: некорректный порядок элементов абзаца')
                    valid={x.get('Id') for x in ET.fromstring(deck.read(rel))} if rel in names else set()
                    referenced={v for el in root.iter() for k,v in el.attrib.items() if k.startswith('{'+R+'}')}
                    if referenced-valid: raise ValueError(f'{slide}: отсутствуют связи {sorted(referenced-valid)}')
        except (BadZipFile,KeyError,ET.ParseError) as error:
            raise ValueError(f'Невозможно проверить презентацию: {error}') from error
        owners_by_product={}
        for t in state['topics']:
            if t.get('owner'):
                key=product_key(t['product'])
                entry=owners_by_product.setdefault(key,{'owners':set(),'ref':f"{t['sheet']}!{t['row']}"})
                entry['owners'].add(product_key(t['owner']))
        for entry in owners_by_product.values():
            if len(entry['owners'])>1:
                warnings.append(f"{entry['ref']}: одно название продукта встречается у нескольких владельцев; в общем количестве названий продукт учтён один раз")
        sources_by_product={}
        for topic in state['topics']:
            key=(product_key(topic.get('owner')),product_key(topic.get('product')))
            entry=sources_by_product.setdefault(key,{'sources':set(),'high':False,'ref':f"{topic['sheet']}!{topic['row']}"})
            entry['sources'].update(str(source).strip() for source in topic.get('sources',[]) if str(source).strip())
            entry['high'] |= topic.get('zone')=='высокое внимание'
        for entry in sources_by_product.values():
            if entry['high'] and (len(entry['sources'])>6 or any(len(name)>50 for name in entry['sources'])):
                warnings.append(f"{entry['ref']}: много источников продукта или длинные названия; проверьте читаемость виджетов")
        for t in state['topics']:
            if len(str(t.get('topic') or ''))>250:
                warnings.append(f"{t['sheet']}!{t['row']}: длинная тема может не поместиться на слайде; проверьте визуально")
            if len(str(t.get('product') or ''))>45:
                warnings.append(f"{t['sheet']}!{t['row']}: длинное название продукта может не поместиться на слайде")
        for t in state['topics']:
            for quote in t.get('quotes',[]):
                if len(str(quote.get('source') or ''))>50:
                    warnings.append(f"{quote['sheet']}!{quote['row']}: длинное название источника может не поместиться на слайде")
            status=t.get('status_record',{})
            if len(status.get('deadline',''))+len(status.get('version',''))>25:
                warnings.append(f"{t['sheet']}!{t['row']}: длинные срок или версия могут не поместиться на слайде")
        if len(str(state.get('narrative') or ''))>450:
            warnings.append('Сводка модели может не поместиться на слайде; проверьте визуально')
        return {'warnings':warnings}
