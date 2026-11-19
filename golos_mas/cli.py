import argparse
import json
import os
import tempfile
from pathlib import Path
from .graph import create_graph
from .verify import file_sha256

def main():
    p=argparse.ArgumentParser(description='Голос IT: сборка отчёта из XLSX')
    p.add_argument('workbook'); p.add_argument('--current-sheet',help='Лист приоритетов; при отсутствии определяется по --report-month')
    previous=p.add_mutually_exclusive_group()
    previous.add_argument('--previous-sheet')
    previous.add_argument('--no-previous',action='store_true',help='Не искать лист предыдущего месяца')
    p.add_argument('--schema-json',help='Необязательное сопоставление столбцов книги')
    p.add_argument('--mau-boundaries',choices=['ambiguous','upper','lower'],default='upper',help='Правило границ 1000 и 20000 при проверке методики')
    p.add_argument('--detail-sheet',action='append',dest='detail_sheets',help='Лист детализации; можно указать несколько раз')
    p.add_argument('--report-month',help='Фильтр дат детализации в формате YYYY-MM')
    p.add_argument('--topic-map-json',help='Явные сопоставления тем приоритетов и детализации')
    p.add_argument('--promises-json',help='Необязательный снимок реестра обещаний')
    p.add_argument('--statuses-json',help='Дополнительный файл статусов и сроков')
    p.add_argument('--template',help='Исходный PPTX образца для сохранения оформления')
    p.add_argument('--period',help='Название периода на слайдах; по умолчанию из --report-month'); p.add_argument('--output',required=True)
    p.add_argument('--gigachat-match',action='store_true',help='Сопоставлять темы через внутренний GigaChat')
    p.add_argument('--gigachat-max-candidates',type=int,default=0,help='Лимит уникальных обращений на тему: 0 — без лимита')
    p.add_argument('--max-quotes-per-topic',type=int,choices=[1,2,3],default=3,help='Максимум выбранных цитат на детальных страницах темы (по умолчанию 3)')
    p.add_argument('--gigachat-quotes',action='store_true',help='Отправлять маскированные обращения во внутренний GigaChat для выбора цитат')
    p.add_argument('--gigachat',action='store_true',help='Разрешить отправку агрегатов во внутренний GigaChat')
    args=p.parse_args()
    if args.gigachat_max_candidates<0: p.error('--gigachat-max-candidates должен быть неотрицательным')
    if args.gigachat_max_candidates and not args.gigachat_quotes: p.error('--gigachat-max-candidates требует --gigachat-quotes')
    if (args.gigachat_quotes or args.gigachat_match) and not args.gigachat: p.error('--gigachat-quotes и --gigachat-match требуют --gigachat')
    if args.report_month:
        from datetime import datetime
        import re
        if not re.fullmatch(r'\d{4}-\d{2}', args.report_month): p.error('--report-month должен быть в формате YYYY-MM')
        try: datetime.strptime(args.report_month, '%Y-%m')
        except ValueError: p.error('--report-month должен быть в формате YYYY-MM')
    if not args.current_sheet and not args.report_month: p.error('Укажите --current-sheet или --report-month')
    if not args.period:
        if not args.report_month: p.error('Укажите --period или --report-month')
        from .periods import format_period
        args.period=format_period(args.report_month)
    output_path=Path(args.output).resolve()
    audit_path=output_path.with_suffix('.audit.json')
    inputs=[Path(value).resolve() for value in (args.workbook,args.template,args.statuses_json,args.promises_json,args.topic_map_json,args.schema_json) if value]
    if any(path in inputs for path in (output_path,audit_path)):
        p.error('Выходная презентация и файл аудита не должны совпадать с исходными файлами')
    llm=None
    if args.gigachat:
        from .gigachat_config import configuration_error, configured_model, ConnectionConfigurationError
        if error:=configuration_error(): p.error(error)
        from langchain_gigachat import GigaChat
        try: llm=configured_model(GigaChat)
        except ConnectionConfigurationError as exc: p.error(str(exc))
        if not callable(getattr(llm, 'bind_tools', None)):
            p.error('Для агентного режима GigaChat должен поддерживать bind_tools')
    result=create_graph(llm, use_model_quotes=args.gigachat_quotes, use_model_matching=args.gigachat_match,max_quote_candidates=args.gigachat_max_candidates).invoke({'workbook':args.workbook,'schema_json':args.schema_json,'mau_boundaries':args.mau_boundaries,'template':args.template,'statuses_json':args.statuses_json,'promises_json':args.promises_json,'topic_map_json':args.topic_map_json,'current_sheet':args.current_sheet,'previous_sheet':args.previous_sheet,'auto_previous':not args.no_previous,'detail_sheets':args.detail_sheets,'period':args.period,'report_month':args.report_month,'max_quotes_per_topic':args.max_quotes_per_topic,'output':args.output})
    report=Path(args.output).with_suffix('.audit.json')
    audit_topics=[]
    for topic in result['topics']:
        entry={k:v for k,v in topic.items() if k not in ('matches','quote_candidates')}
        entry['matched_references']=[{'sheet':r['sheet'],'row':r['row']} for r in topic.get('matches',[])]
        audit_topics.append(entry)
    digest=file_sha256(result['output'])
    payload={'presentation_sha256':digest,'warnings':result['warnings'],'agent_decisions':result.get('agent_decisions',[]),'topics':audit_topics,'slide_count':result['slide_count'],'promise_report':result.get('promise_report'),'coverage':result.get('coverage'),'date_filter':result.get('date_filter'),'skipped_rows':result.get('skipped_rows',[]),'narrative':result.get('narrative',''),'narrative_sources':result.get('narrative_sources',[]),'source_sheets':{'current':result.get('current_sheet'),'previous':result.get('previous_sheet'),'detail':result.get('detail_sheets'),'schema':args.schema_json}}
    temporary=None
    try:
        with tempfile.NamedTemporaryFile(mode='w',encoding='utf8',prefix='.golos-audit-',suffix='.json',dir=report.resolve().parent,delete=False) as handle:
            temporary=Path(handle.name);json.dump(payload,handle,ensure_ascii=False,indent=2,default=str)
        os.replace(temporary,report)
    except BaseException:
        if temporary is not None: temporary.unlink(missing_ok=True)
        raise
    print(f"Готово: {result['output']}; слайдов: {result['slide_count']}; предупреждений: {len(result['warnings'])}; аудит: {report}")
if __name__=='__main__': main()
