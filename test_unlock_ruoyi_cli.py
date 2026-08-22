# -*- coding: utf-8 -*-
import os, types
from unlock_outlook_ruoyi import build_parser, load_accounts, _load_unlock_proxies

def test_defaults():
    a = build_parser().parse_args([])
    assert a.input == "email_abuse.txt"
    assert a.concurrency == 2
    assert a.headless is False
    assert a.limit == 0
    assert a.max_press == 5

def test_limit_slice(tmp_path):
    f = tmp_path / "acc.txt"
    f.write_text("a@x.com----pw1----abuse----ts\nb@x.com----pw2\nc@x.com----pw3\n", encoding="utf-8")
    accs = load_accounts(str(f), limit=2)
    assert len(accs) == 2
    assert accs[0][0] == "a@x.com"
    assert accs[1][0] == "b@x.com"

def test_proxy_priority_explicit_wins(monkeypatch):
    # --proxy 优先于 --proxy-url/--proxy-file
    a = build_parser().parse_args(["--proxy", "socks5://u:p@1.2.3.4:1080",
                                   "--proxy-url", "http://example.com/list"])
    out = _load_unlock_proxies(a)
    assert out == ["socks5://u:p@1.2.3.4:1080"]

def test_proxy_url_used_when_no_explicit(monkeypatch):
    called = {}
    import register_outlook_ruoyi as rr
    monkeypatch.setattr(rr, "fetch_proxy_list_http", lambda url: (called.setdefault('url', url), ["socks5://a:b@5.6.7.8:1080"])[1])
    a = build_parser().parse_args(["--proxy-url", "http://example.com/list"])
    out = _load_unlock_proxies(a)
    assert called['url'] == "http://example.com/list"
    assert out == ["socks5://a:b@5.6.7.8:1080"]

def test_no_proxy_returns_empty():
    a = build_parser().parse_args(["--proxy-url", "", "--proxy-file", "", "--proxy", ""])
    assert _load_unlock_proxies(a) == []
