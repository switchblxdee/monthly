"""Extract the legend and methodology icons from the supplied PowerPoint template."""
import argparse
import html
import json
import re
import shutil
from pathlib import Path
from zipfile import ZipFile

from .template_render import ICON_DESCRIPTIONS, N, R, collect_legend, content, part_from_target, position, xml

GROUPS={
    'new':'recurrence','two':'recurrence','long':'recurrence','resumed':'recurrence',
    'aibolit':'controls','promise_monitor':'controls','pain_candidate':'controls','backlog':'controls',
    'done':'mitigation','analysis':'mitigation','release':'mitigation','limitations':'mitigation',
    'promnenie':'additional','after_close':'additional',
}
METHODOLOGY={
    'mau_small':('Малая аудитория', 'audience',6.67,1.57),
    'mau_medium':('Средняя аудитория', 'audience',8.62,1.57),
    'mau_large':('Большая аудитория', 'audience',11.19,1.57),
    'zone_observation':('Зона наблюдения', 'zone',6.98,5.86),
    'zone_risk':('Зона потенциального риска', 'zone',8.83,5.87),
    'zone_high':('Зона высокого внимания', 'zone',10.75,5.87),
}
GROUP_TITLES={'audience':'Категория аудитории MAU','zone':'Зона приоритета','recurrence':'Повторяемость темы',
              'controls':'Контроль решения','mitigation':'Статус митигации','additional':'Дополнительная информация'}

GUIDE='''# Значки «Голос IT»

Изображения извлечены из методики на слайде 4 и легенды на слайде 6 исходного образца. Откройте `index.html` для просмотра каталога. `index.json` связывает имя файла, код, описание, группу и слайд. Логотипы и декоративные изображения в набор не входят.

- `audience`: `mau_small`, `mau_medium`, `mau_large` — категории продукта на слайде методики. По подтверждённому правилу 1 000 — средний, 20 000 — большой.
- `zone`: `zone_observation`, `zone_risk`, `zone_high` — значки зон на слайде методики.
- `recurrence`: значок повторяемости в одноимённой колонке. Значения: `new`, `two`, `long`, `resumed`.
- `controls`: значки в области контроля решения. Значения: `aibolit`, `promise_monitor`, `pain_candidate`, `backlog`.
- `mitigation`: значок текущего статуса рядом со сроком. Значения: `done`, `analysis`, `release`, `limitations`.
- `additional`: `promnenie` — отметка у темы; `after_close` — дополнительный значок в области контроля.

В генераторе балл длительности 2 или 4 может задавать значок повторяемости. Явный статус имеет приоритет; другие значки выставляются по проверенному дополнительному файлу статусов. Назначение не следует угадывать из отсутствующих данных. Полные правила и пример файла статусов — в README проекта.
'''


def methodology_icons(root):
    pictures=root.findall('.//p:pic',N);result={}
    for code,(_,_,expected_x,expected_y) in METHODOLOGY.items():
        matches=[]
        for pic in pictures:
            x,y=position(pic)
            if x is not None and y is not None and abs(x/914400-expected_x)<.16 and abs(y/914400-expected_y)<.16:
                matches.append(pic)
        if len(matches)!=1: raise ValueError(f'Шаблон: значок методики {code} не найден однозначно')
        result[code]=matches[0]
    return result


def gallery(entries):
    sections=[]
    for group,title in GROUP_TITLES.items():
        cards=[]
        for item in entries:
            if item['group']!=group: continue
            name=html.escape(item['description']);filename=html.escape(item['file'],quote=True)
            code=html.escape(item['code'])
            cards.append(f'<article><img src="{filename}" alt="{name}"><div><strong>{name}</strong><code>{code}</code><small>Слайд {item["slide"]}</small></div></article>')
        sections.append(f'<section><h2>{html.escape(title)}</h2><div class="grid">{"".join(cards)}</div></section>')
    return '''<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Значки «Голос IT»</title><style>body{font-family:Arial,sans-serif;background:#f4f6fb;color:#1a2940;margin:0}main{max-width:1080px;margin:auto;padding:32px 20px}h1{font-size:30px}h2{font-size:20px;margin-top:34px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px}article{display:flex;align-items:center;gap:14px;background:white;border-radius:12px;padding:16px;min-height:68px;box-shadow:0 2px 8px #172b4d12}img{width:48px;height:48px;object-fit:contain}strong,code,small{display:block}code{color:#346ed7;margin:5px 0}small{color:#667085}p{line-height:1.5}</style></head><body><main><h1>Значки «Голос IT»</h1><p>Изображения из методики и легенды исходного образца. Правила применения и границы MAU описаны в README этого набора.</p>''' + ''.join(sections) + '</main></body></html>\n'


def extract_icons(template, output_dir):
    target=Path(output_dir)
    # A new directory prevents accidental replacement of user files.
    target.mkdir(parents=True,exist_ok=False)
    try:
        with ZipFile(template) as deck:
            slide4=xml(deck.read('ppt/slides/slide4.xml'))
            slide6=xml(deck.read('ppt/slides/slide6.xml'))
            for slide,root,landmark in ((4,slide4,'Методика определения приоритезации'),(6,slide6,'Темы, требующие внимания')):
                if re.sub(r'\s+','',landmark).casefold() not in re.sub(r'\s+','',content(root)).casefold():
                    raise ValueError(f'Шаблон не соответствует образцу: слайд {slide} не содержит нужного заголовка')
            legend=collect_legend(slide6)
            missing=set(ICON_DESCRIPTIONS)-set(legend)
            if missing: raise ValueError('В шаблоне не найдены значки: '+', '.join(sorted(missing)))
            methodology=methodology_icons(slide4)
            entries=[]
            items=[(code,ICON_DESCRIPTIONS[code],GROUPS[code],6,legend[code]) for code in ICON_DESCRIPTIONS]
            items += [(code,info[0],info[1],4,methodology[code]) for code,info in METHODOLOGY.items()]
            relationships={number:{rel.get('Id'):rel for rel in xml(deck.read(f'ppt/slides/_rels/slide{number}.xml.rels'))} for number in (4,6)}
            for code,description,group,slide,picture in items:
                blip=picture.find('.//a:blip',N)
                rid=blip.get(f'{{{R}}}embed') if blip is not None else None
                rel=relationships[slide].get(rid)
                if rel is None or rel.get('TargetMode')=='External' or not rel.get('Type','').endswith('/image'):
                    raise ValueError(f'Значок {code}: отсутствует встроенное изображение')
                part=part_from_target(f'ppt/slides/slide{slide}.xml',rel.get('Target',''))
                ext=Path(part).suffix.lower()
                if not part.startswith('ppt/media/') or ext not in ('.png','.jpg','.jpeg','.svg','.emf','.wmf'):
                    raise ValueError(f'Значок {code}: неподдерживаемый тип изображения')
                filename=code+ext
                (target/filename).write_bytes(deck.read(part))
                entries.append({'code':code,'description':description,'group':group,'slide':slide,'file':filename})
        (target/'index.json').write_text(json.dumps({'icons':entries},ensure_ascii=False,indent=2)+'\n',encoding='utf8')
        (target/'index.html').write_text(gallery(entries),encoding='utf8')
        (target/'README.md').write_text(GUIDE,encoding='utf8')
        return target
    except BaseException:
        shutil.rmtree(target)
        raise


def main():
    parser=argparse.ArgumentParser(description='Извлечь значки легенды и методики из шаблона презентации')
    parser.add_argument('template');parser.add_argument('--output-dir',required=True,help='Новая папка для значков')
    args=parser.parse_args()
    try: result=extract_icons(args.template,args.output_dir)
    except FileExistsError: parser.error('Папка уже существует: выберите новое имя')
    print(f'Значки и index.json сохранены: {result}')


if __name__=='__main__': main()
