# -*- coding: utf-8 -*-
from unlock_outlook_ruoyi import classify

def test_logged_in_account_microsoft():
    assert classify("anything", "https://account.microsoft.com/?lang=zh") == "logged_in"

def test_logged_in_proofs():
    assert classify("x", "https://account.live.com/proofs/manage") == "logged_in"

def test_locked_en():
    assert classify("Your account has been locked", "https://login.live.com/something") == "locked"

def test_locked_zh():
    assert classify("帐户已锁定", "https://login.live.com/x") == "locked"

def test_px_challenge():
    assert classify("Let's prove you're human. Press and hold", "https://login.live.com/x") == "px_challenge"

def test_sms_verify():
    assert classify("We texted a code to your phone. Enter the code.", "https://login.live.com/x") == "sms_verify"

def test_fido_setup_url():
    assert classify("setting up", "https://login.live.com/fido/create") == "fido_setup"

def test_login_form():
    assert classify("Enter your password", "https://login.live.com/x") == "login_form"

def test_email_form():
    assert classify("Sign in. Enter your email.", "https://login.live.com/x") == "email_form"

def test_net_error():
    assert classify("", "chrome-error://notallowed/") == "net_error"

def test_unknown():
    assert classify("hello world", "https://login.live.com/x") == "unknown"
