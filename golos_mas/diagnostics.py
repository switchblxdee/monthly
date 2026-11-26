"""Verify the GigaChat configuration and optionally send synthetic requests."""
import argparse
import importlib
import importlib.util
import os
from pathlib import Path
from .gigachat_config import configuration_error, configured_model, ConnectionConfigurationError


def check_tools(model):
    from .worker_agents import ReportWorker

    def count_sample(_state):
        return {'signals': [{}, {}]}

    def confirm_sample(state):
        return {'sample_confirmed': len(state['signals']) == 2}

    worker = ReportWorker(model, 'diagnostics', 'проверь вызов двух инструментов на учебном примере', [
        ('count_sample', 'Подсчитать две учебные строки.', count_sample),
        ('confirm_sample', 'Подтвердить результат подсчёта учебных строк.', confirm_sample),
    ])
    result = worker({})
    if not result.get('sample_confirmed'):
        raise ValueError('Проверка результата инструмента не завершена')


def main(argv=None):
    parser=argparse.ArgumentParser(description='Проверка окружения GigaChat без данных Excel')
    mode=parser.add_mutually_exclusive_group()
    mode.add_argument('--ping',action='store_true',help='Отправить фиксированный тестовый запрос во внутренний GigaChat')
    mode.add_argument('--check-tools',action='store_true',help='Проверить вызовы инструментов LangGraph на учебных данных')
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
        elif args.check_tools:
            check_tools(gigachat)
    except ConnectionConfigurationError as exc:
        print(str(exc))
        return 1
    except Exception as exc:
        # Do not expose server responses, credentials or endpoint in diagnostics.
        print(f'Проверка не пройдена: {type(exc).__name__}')
        return 1
    if args.check_tools:
        print('GigaChat выполнил вызовы инструментов LangGraph на учебных данных.')
    elif args.ping:
        print('Тестовый запрос выполнен.')
    else:
        print('Зависимости и конфигурация проверены; сетевой запрос не выполнялся.')
    return 0

if __name__=='__main__':
    raise SystemExit(main())
