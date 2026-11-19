"""Check that an audit file belongs to the accompanying presentation."""
import argparse
import hashlib
import hmac
import json
from pathlib import Path


def file_sha256(path):
    digest=hashlib.sha256()
    with open(path,'rb') as handle:
        for chunk in iter(lambda: handle.read(1024*1024),b''):
            digest.update(chunk)
    return digest.hexdigest()


def verify(deck_path, audit_path=None):
    deck=Path(deck_path)
    audit=Path(audit_path) if audit_path else deck.with_suffix('.audit.json')
    with audit.open(encoding='utf8') as handle:
        data=json.load(handle)
    expected=data.get('presentation_sha256') if isinstance(data,dict) else None
    if not isinstance(expected,str) or len(expected)!=64 or any(c not in '0123456789abcdef' for c in expected):
        raise ValueError('В аудите нет корректного SHA-256 презентации; сформируйте отчёт заново')
    actual=file_sha256(deck)
    return hmac.compare_digest(actual,expected)


def main():
    parser=argparse.ArgumentParser(description='Проверить соответствие презентации и аудита')
    parser.add_argument('presentation')
    parser.add_argument('--audit',help='Путь к аудиту, если он расположен отдельно')
    args=parser.parse_args()
    try:
        matches=verify(args.presentation,args.audit)
    except (OSError,ValueError) as exc:
        parser.exit(2,f'Не удалось проверить файлы: {exc}\n')
    if not matches:
        parser.exit(1,'Презентация и аудит не совпадают.\n')
    print('Презентация соответствует аудиту.')


if __name__=='__main__': main()
