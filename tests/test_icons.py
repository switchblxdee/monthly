import json
import os
from pathlib import Path
from zipfile import ZipFile

import pytest

from golos_mas.icons import METHODOLOGY, extract_icons, methodology_icons
from golos_mas.template_render import ICON_DESCRIPTIONS, N, R, collect_legend, part_from_target, xml


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'),reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_extracted_icons_are_original_images_with_labels_and_no_overwrite(tmp_path):
    template=os.environ['GOLOS_TEST_TEMPLATE'];out=extract_icons(template,tmp_path/'icons')
    manifest=json.loads((out/'index.json').read_text(encoding='utf8'))['icons']
    assert {item['code'] for item in manifest}==set(ICON_DESCRIPTIONS)|set(METHODOLOGY)
    gallery=(out/'index.html').read_text(encoding='utf8')
    assert all(f'src="{item["file"]}"' in gallery and item['description'] in gallery for item in manifest)
    with ZipFile(template) as deck:
        legend=collect_legend(xml(deck.read('ppt/slides/slide6.xml')))
        method=methodology_icons(xml(deck.read('ppt/slides/slide4.xml')))
        pictures={4:method,6:legend}
        relationships={slide:{r.get('Id'):r for r in xml(deck.read(f'ppt/slides/_rels/slide{slide}.xml.rels'))} for slide in (4,6)}
        for item in manifest:
            expected=ICON_DESCRIPTIONS[item['code']] if item['slide']==6 else METHODOLOGY[item['code']][0]
            assert item['description']==expected
            rid=pictures[item['slide']][item['code']].find('.//a:blip',N).get(f'{{{R}}}embed')
            part=part_from_target(f"ppt/slides/slide{item['slide']}.xml",relationships[item['slide']][rid].get('Target'))
            assert (out/item['file']).read_bytes()==deck.read(part)
    before=(out/'index.json').read_bytes()
    with pytest.raises(FileExistsError): extract_icons(template,out)
    assert (out/'index.json').read_bytes()==before


def test_invalid_template_does_not_leave_partial_icon_folder(tmp_path):
    bad=tmp_path/'bad.pptx'
    with ZipFile(bad,'w') as deck: deck.writestr('dummy','')
    target=tmp_path/'icons'
    with pytest.raises(KeyError): extract_icons(bad,target)
    assert not target.exists()


@pytest.mark.skipif(not os.environ.get('GOLOS_TEST_TEMPLATE'),reason='Set GOLOS_TEST_TEMPLATE to the original example deck')
def test_unrecognised_methodology_slide_is_not_mislabelled(tmp_path):
    import xml.etree.ElementTree as ET
    source=os.environ['GOLOS_TEST_TEMPLATE'];bad=tmp_path/'wrong.pptx';target=tmp_path/'icons'
    with ZipFile(source) as original, ZipFile(bad,'w') as modified:
        slide=xml(original.read('ppt/slides/slide4.xml'))
        for text in slide.findall('.//a:t',N): text.text='Другая презентация'
        modified.writestr('ppt/slides/slide4.xml',ET.tostring(slide,encoding='utf-8',xml_declaration=True))
        modified.writestr('ppt/slides/slide6.xml',original.read('ppt/slides/slide6.xml'))
    with pytest.raises(ValueError,match='Шаблон не соответствует образцу'):
        extract_icons(bad,target)
    assert not target.exists()
