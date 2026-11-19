"""Produce a schema/count report without copying worksheet names or cell contents."""
import argparse
import json
import os
import tempfile
from itertools import zip_longest
from pathlib import Path
from openpyxl import load_workbook
from .agents import COLUMN_ALIASES, load_schema_config, resolve_columns, locate, norm

class WorksheetRangeMismatch(ValueError):
    pass


def inspect_workbook(path,max_rows=100000,schema_path=None):
    if max_rows<1: raise ValueError('max_rows должен быть положительным')
    values=load_workbook(path,read_only=True,data_only=True)
    formulas=None
    try:
        formulas=load_workbook(path,read_only=True,data_only=False)
        schema=load_schema_config(schema_path,values.sheetnames)
        reports=[]
        for number,name in enumerate(values.sheetnames,1):
            ws=values[name];raw=formulas[name]
            item={'sheet_index':number,'sheet_alias':f'Sheet_{number}','recognized':False}
            if not hasattr(ws,'iter_rows'):
                item['kind']='unsupported';reports.append(item);continue
            override=schema.get(name)
            try: header,_,columns=resolve_columns(ws,override,strict=False)
            except ValueError:
                if override is not None: raise
                item['kind']='unrecognized';reports.append(item);continue
            item.update({'recognized':True,'header_row':header,'schema_applied':override is not None,
                         'fields':{key:index is not None for key,index in columns.items()}})
            kinds=[]
            if columns['score'] is not None or all(columns[x] is not None for x in ('signals_score','duration','impact')): kinds.append('priority')
            if columns['text'] is not None and columns['signal'] is not None: kinds.append('detail')
            item['kind']='+'.join(kinds) if kinds else 'partial'
            # Output only schema identifiers and counts; neither headers nor cell values are copied.
            duplicate_fields=[]
            header_values=next(ws.iter_rows(min_row=header,max_row=header,values_only=True))
            for key,aliases in COLUMN_ALIASES.items():
                if sum(locate((value,),*aliases) is not None for value in header_values)>1 and (override is None or key not in override['columns']): duplicate_fields.append(key)
            item['duplicate_recognized_fields']=duplicate_fields
            last=raw.max_row
            last_checked=min(last,header+max_rows) if last is not None else header+max_rows+1
            item['truncated']=last is not None and last>last_checked
            item['rows_scanned']=0
            item['rows_with_recognized_values']=0;item['rows_with_product_and_topic']=0;item['rows_with_only_one_key']=0
            item['rows_with_no_keys_but_other_values']=0
            item['rows_with_no_keys_and_uncached_formula']=0
            item['formulas']={key:{'total':0,'missing_cached_value':0,'error_cached_value':0} for key,index in columns.items() if index is not None}
            if last_checked>header:
                max_col=max(index for index in columns.values() if index is not None)+1
                source_rows=raw.iter_rows(min_row=header+1,max_row=last_checked,max_col=max_col)
                value_rows=ws.iter_rows(min_row=header+1,max_row=last_checked,max_col=max_col)
                missing_row=object()
                for number,(source_row,value_row) in enumerate(zip_longest(source_rows,value_rows,fillvalue=missing_row),1):
                    if source_row is missing_row or value_row is missing_row:
                        raise WorksheetRangeMismatch('Диапазоны значений и формул не совпадают')
                    if number>max_rows:
                        item['truncated']=True
                        break
                    item['rows_scanned']+=1
                    def cached(key):
                        index=columns[key]
                        return value_row[index].value if index is not None else None
                    if any(source_row[i].value is not None for i in set(x for x in columns.values() if x is not None)):
                        item['rows_with_recognized_values']+=1
                    product,topic=bool(norm(cached('product'))),bool(norm(cached('topic')))
                    item['rows_with_product_and_topic']+=int(product and topic)
                    item['rows_with_only_one_key']+=int(product != topic)
                    if not product and not topic:
                        values=(cached(key) for key in columns if key not in ('product','topic'))
                        item['rows_with_no_keys_but_other_values']+=int(any(v is not None and (not isinstance(v,str) or v.strip()) for v in values))
                        item['rows_with_no_keys_and_uncached_formula']+=int(any(source_row[index].data_type=='f' and value_row[index].value is None for index in set(x for x in columns.values() if x is not None)))
                    for key,index in columns.items():
                        if index is not None and source_row[index].data_type=='f':
                            stats=item['formulas'][key];stats['total']+=1
                            if value_row[index].value is None: stats['missing_cached_value']+=1
                            elif value_row[index].data_type=='e': stats['error_cached_value']+=1
            reports.append(item)
        return {'schema_version':1,'sheet_count':len(reports),'sheets':reports}
    finally:
        values.close()
        if formulas is not None: formulas.close()


def main(argv=None):
    parser=argparse.ArgumentParser(description='Сформировать отчёт о структуре XLSX без значений ячеек')
    parser.add_argument('workbook');parser.add_argument('--schema-json',help='Необязательная схема столбцов');parser.add_argument('--output',help='JSON-файл; без параметра отчёт выводится в stdout')
    parser.add_argument('--max-rows',type=int,default=100000,help='Максимум строк на распознанном листе, по умолчанию 100000')
    args=parser.parse_args(argv)
    if args.max_rows<1: parser.error('--max-rows должен быть положительным')
    if args.output and Path(args.output).resolve() in [Path(x).resolve() for x in (args.workbook,args.schema_json) if x]: parser.error('Выходной файл совпадает с исходной книгой или схемой')
    try: report=inspect_workbook(args.workbook,args.max_rows,args.schema_json)
    except WorksheetRangeMismatch:
        print('Не удалось проверить книгу: диапазоны значений и формул не совпадают')
        return 1
    except Exception as exc:
        print(f'Не удалось проверить книгу: {type(exc).__name__}')
        return 1
    text=json.dumps(report,ensure_ascii=False,indent=2)
    if not args.output:
        print(text);return 0
    target=Path(args.output).resolve();target.parent.mkdir(parents=True,exist_ok=True);temp=None
    try:
        with tempfile.NamedTemporaryFile(mode='w',encoding='utf8',prefix='.golos-schema-',suffix='.json',dir=target.parent,delete=False) as handle:
            temp=Path(handle.name);handle.write(text+'\n')
        os.replace(temp,target)
    except BaseException:
        if temp is not None: temp.unlink(missing_ok=True)
        raise
    print('Отчёт записан. Перед передачей проверьте, допустимо ли делиться агрегированными показателями.')
    return 0

if __name__=='__main__': raise SystemExit(main())
