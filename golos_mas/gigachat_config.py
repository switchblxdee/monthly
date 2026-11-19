"""Checks required before using the organisation's GigaChat deployment."""
import os

class ConnectionConfigurationError(ValueError):
    pass


def configuration_error():
    if not os.environ.get('GIGACHAT_BASE_URL','').strip():
        return 'Для работы во внутреннем контуре задайте GIGACHAT_BASE_URL явно'
    if os.environ.get('GIGACHAT_CREDENTIALS','').strip() and not os.environ.get('GIGACHAT_AUTH_URL','').strip():
        return 'При GIGACHAT_CREDENTIALS задайте GIGACHAT_AUTH_URL явно: адрес OAuth независим от адреса модели'
    return None


def connection_kwargs():
    result={'base_url':os.environ['GIGACHAT_BASE_URL'].strip()}
    auth_url=os.environ.get('GIGACHAT_AUTH_URL','').strip()
    if auth_url: result['auth_url']=auth_url
    return result


def configured_model(factory):
    """Require the wrapper to expose the explicitly configured endpoints."""
    kwargs=connection_kwargs()
    model=factory(**kwargs)
    for field,expected in kwargs.items():
        if getattr(model,field,None)!=expected:
            raise ConnectionConfigurationError(f'Библиотека GigaChat не подтвердила параметр {field}; проверьте совместимую версию и настройки')
    return model
