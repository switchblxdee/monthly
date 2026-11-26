from types import SimpleNamespace
from golos_mas import diagnostics


def configure(monkeypatch, factory):
    monkeypatch.setenv('GIGACHAT_BASE_URL','https://internal.example/v1')
    monkeypatch.delenv('GIGACHAT_CA_BUNDLE_FILE',raising=False)
    monkeypatch.delenv('GIGACHAT_CREDENTIALS',raising=False)
    monkeypatch.delenv('GIGACHAT_AUTH_URL',raising=False)
    monkeypatch.delenv('GIGACHAT_MODEL',raising=False)
    monkeypatch.setattr(diagnostics.importlib.util,'find_spec',lambda _: object())
    monkeypatch.setattr(diagnostics.importlib,'import_module',lambda _: SimpleNamespace(GigaChat=factory))


def test_default_diagnostics_does_not_invoke_model(monkeypatch,capsys):
    calls=[]
    class Model:
        def __init__(self,**kwargs): self.__dict__.update(kwargs)
        def invoke(self,text): calls.append(text)
    configure(monkeypatch,Model)
    assert diagnostics.main([])==0 and not calls
    assert 'сетевой запрос не выполнялся' in capsys.readouterr().out


def test_ping_sends_only_fixed_prompt_and_hides_reply(monkeypatch,capsys):
    calls=[]
    class Model:
        def __init__(self,**kwargs): self.__dict__.update(kwargs)
        def invoke(self,text): calls.append(text);return SimpleNamespace(content='private response')
    configure(monkeypatch,Model)
    assert diagnostics.main(['--ping'])==0
    assert calls==['Ответь одним словом: ОК']
    assert 'private response' not in capsys.readouterr().out


def test_ping_error_does_not_print_exception_contents(monkeypatch,capsys):
    class Model:
        def __init__(self,**kwargs): self.__dict__.update(kwargs)
        def invoke(self,text): raise ValueError('secret response or token')
    configure(monkeypatch,Model)
    assert diagnostics.main(['--ping'])==1
    out=capsys.readouterr().out
    assert 'ValueError' in out and 'secret response or token' not in out


def test_oauth_address_required_and_passed_to_model(monkeypatch,capsys):
    calls=[]
    def factory(**kwargs): calls.append(kwargs);return SimpleNamespace(**kwargs)
    configure(monkeypatch,factory)
    monkeypatch.setenv('GIGACHAT_CREDENTIALS','secret')
    assert diagnostics.main([])==1 and not calls
    assert 'GIGACHAT_AUTH_URL' in capsys.readouterr().out
    monkeypatch.setenv('GIGACHAT_AUTH_URL','https://auth.internal/token')
    assert diagnostics.main([])==0
    assert calls==[{'base_url':'https://internal.example/v1','auth_url':'https://auth.internal/token','model':'GigaChat'}]


def test_unconfirmed_endpoint_blocks_ping_without_printing_url(monkeypatch,capsys):
    calls=[]
    class Model:
        def __init__(self,**kwargs): self.base_url=kwargs['base_url'];self.auth_url=None
        def invoke(self,text): calls.append(text)
    configure(monkeypatch,Model)
    monkeypatch.setenv('GIGACHAT_CREDENTIALS','secret')
    monkeypatch.setenv('GIGACHAT_AUTH_URL','https://private.example/token')
    assert diagnostics.main(['--ping'])==1 and not calls
    out=capsys.readouterr().out
    assert 'не подтвердила параметр auth_url' in out and 'private.example' not in out and 'secret' not in out
