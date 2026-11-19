"""Local diagnostics; optional ping sends only a fixed synthetic prompt."""
import argparse
import importlib
import importlib.util
import os
from pathlib import Path
from .gigachat_config import configuration_error, configured_model, ConnectionConfigurationError


def main(argv=None):
    parser=argparse.ArgumentParser(description='Проверка окружения GigaChat без данных Excel')
    parser.add_argument('--ping',action='store_true',help='Отправить фиксированный тестовый запрос во внутренний GigaChat')
    args=parser.parse_args(argv)
    missing=[]
    for module in ('openpyxl','langgraph','langchain_core','langchain_gigachat'):
        if importlib.util.find_spec(module) is None: missing.append(module)
    if missing:
        print('Не найдены модули: '+', '.join(missing))
        return 1
    if error:=configuration_error():
        print(error)
        return 1
    bundle=os.environ.get('GIGACHAT_CA_BUNDLE_FILE')
    if bundle and not Path(bundle).is_file():
        print('Указанный файл GIGACHAT_CA_BUNDLE_FILE не найден')
        return 1
    try:
        gigachat=configured_model(importlib.import_module('langchain_gigachat').GigaChat)
        if args.ping:
            response=gigachat.invoke('Ответь одним словом: ОК')
            if not isinstance(response.content,str) or not response.content.strip(): raise ValueError('Пустой ответ')
    except ConnectionConfigurationError as exc:
        print(str(exc))
        return 1
    except Exception as exc:
        # Do not expose server responses, credentials or endpoint in diagnostics.
        print(f'Проверка не пройдена: {type(exc).__name__}')
        return 1
    print('Тестовый запрос выполнен.' if args.ping else 'Зависимости и конфигурация проверены; сетевой запрос не выполнялся.')
    return 0

if __name__=='__main__':
    raise SystemExit(main())
