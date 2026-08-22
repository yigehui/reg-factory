"""PX 兼容性对拍脚本:在两台机器上各跑一次,对比输出定位 PX 过不去的差异。

用法: .venv/Scripts/python.exe px_env_diff.py
"""
import os
import sys
import hashlib
import json
import subprocess

ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT)

report = {}


def sec(title):
    print("\n===== %s =====" % title)


# ── 1. ruyipage 库 + 内核 ────────────────────────────────────────────
sec("1. ruyipage 库")
try:
    import ruyipage
    from ruyipage import FirefoxOptions, FirefoxPage
    import importlib.metadata as md
    ver = md.version("ruyipage")
    report["ruyipage"] = ver
    print("ruyipage:", ver)
except Exception as e:
    report["ruyipage"] = "ERROR: %r" % e
    print("import failed:", e)
    sys.exit(1)

sec("2. 浏览器内核解析")
out = subprocess.run([sys.executable, "-m", "ruyipage", "path"],
                     capture_output=True, text=True, timeout=15)
kernel_path = out.stdout.strip().splitlines()[-1] if out.stdout.strip() else ""
report["kernel_path"] = kernel_path
print("ruyipage path ->", kernel_path, "(exists=%s)" % os.path.isfile(kernel_path))

# 项目自身的解析结果(可能与上面不同)
env_val = os.environ.get("RUOYI_FIREFOX_PATH", "")
report["env_RUOYI_FIREFOX_PATH"] = env_val or "(unset)"
print("env RUOYI_FIREFOX_PATH:", env_val or "(unset)")

browsers_root = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ruyipage", "browsers")
if os.path.isdir(browsers_root):
    kernels = sorted(os.listdir(browsers_root))
    report["installed_kernels"] = kernels
    print("installed kernels:", kernels)

# ── 2. OS 环境 ──────────────────────────────────────────────────────
sec("3. OS 环境")
import winreg


def reg(path, name):
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
        return winreg.QueryValueEx(k, name)[0]
    except OSError:
        return None


ps = subprocess.run(["powershell", "-NoProfile", "-Command",
                     "(Get-TimeZone).Id;"
                     "(Get-CimInstance Win32_OperatingSystem).Caption;"
                     "(Get-CimInstance Win32_Processor).Name;"
                     "(Get-CimInstance Win32_VideoController).Name -join ' | ';"
                     "(Get-WinSystemLocale).Name;"
                     "(Get-Culture).Name"],
                    capture_output=True, text=True)
lines = [l for l in ps.stdout.splitlines() if l.strip()]
labels = ["timezone", "os", "cpu", "gpu", "system_locale", "culture"]
for lab, val in zip(labels, lines):
    report[lab] = val
    print("%-14s %s" % (lab + ":", val))
dpi = reg(r"Control Panel\Desktop\WindowMetrics", "AppliedDPI")
report["AppliedDPI"] = dpi
print("%-14s %s (%d%% 缩放)" % ("DPI:", dpi, round(dpi / 96 * 100)))

# ── 3. 时钟偏移 ─────────────────────────────────────────────────────
sec("4. 时钟偏移(NTP)")
try:
    import ntplib  # 可能没装,失败就跳过
    c = ntplib.NTPClient()
    resp = c.request("time.windows.com", version=3, timeout=5)
    report["clock_offset_sec"] = round(resp.offset, 2)
    print("offset: %.2f s" % resp.offset)
except Exception as e:
    report["clock_offset_sec"] = "skip (%s)" % type(e).__name__
    print("skip:", e)

# ── 4. 真实浏览器指纹(canvas/WebGL/audio) ────────────────────────────
sec("5. 浏览器指纹(155 内核实测)")
JS = r"""
(() => {
  const out = {};
  // canvas
  try {
    const c = document.createElement('canvas');
    c.width = 260; c.height = 60;
    const g = c.getContext('2d');
    g.textBaseline = 'top';
    g.font = "14px 'Arial'";
    g.fillStyle = '#f60';
    g.fillRect(125, 1, 62, 20);
    g.fillStyle = '#069';
    g.fillText('CryptoFp,:;<>', 2, 15);
    g.fillStyle = 'rgba(102,204,0,0.7)';
    g.fillText('CryptoFp,:;<>', 4, 17);
    out.canvas = Array.from(c.toDataURL()).filter(ch => ch !== ',').length;
    out.canvas_hash = Array.from(new Uint8Array(atob(c.toDataURL().split(',')[1])))
      .reduce((h, b) => (h * 31 + b) >>> 0, 7).toString(16);
  } catch (e) { out.canvas = 'err ' + e; }
  // webgl
  try {
    const gl = document.createElement('canvas').getContext('webgl');
    const dbg = gl.getExtension('WEBGL_debug_renderer_info');
    out.webgl_vendor = dbg ? gl.getParameter(dbg.UNMASKED_VENDOR_WEBGL) : gl.getParameter(gl.VENDOR);
    out.webgl_renderer = dbg ? gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) : gl.getParameter(gl.RENDERER);
  } catch (e) { out.webgl = 'err ' + e; }
  out.ua = navigator.userAgent;
  out.lang = navigator.language + '/' + navigator.languages.join(',');
  out.tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
  out.hw_concurrency = navigator.hardwareConcurrency;
  out.device_memory = navigator.deviceMemory || 'n/a';
  out.pixel_ratio = window.devicePixelRatio;
  out.screen = screen.width + 'x' + screen.height;
  out.plugins = Array.from(navigator.plugins).map(p => p.name).join(';');
  return out;
})()
"""
try:
    opts = FirefoxOptions()
    if kernel_path and os.path.isfile(kernel_path):
        opts.set_browser_path(kernel_path)
    opts.headless(False)  # 与生产一致:有头
    page = FirefoxPage(opts)
    try:
        page.get("https://example.com")
        fp_raw = page.run_js("return JSON.stringify(" + JS + ")")
        import json as _json
        fp = _json.loads(fp_raw) if isinstance(fp_raw, str) else fp_raw
        report["browser_fp"] = fp
        for k, v in fp.items():
            print("%-18s %s" % (k + ":", v))
    finally:
        page.quit()
except Exception as e:
    report["browser_fp"] = "ERROR: %r" % e
    print("browser fp failed:", repr(e))

# ── 5. 出口 IP(走代理,如果有配置) ──────────────────────────────────
sec("6. 代理出口")
try:
    sys.path.insert(0, ROOT)
    # 只读代理文件,不起浏览器
    pf = os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt")
    if os.path.isfile(pf):
        has_entry = False
        for line in open(pf, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#"):
                has_entry = True
                break
        report["proxy_first_line_prefix"] = "(存在,内容不记录)"
        print("proxy file:", pf, "有配置(凭据不记录)")
    else:
        report["proxy_first_line_prefix"] = "(no proxy file)"
        print("proxy file missing:", pf)
except Exception as e:
    print("proxy check failed:", e)

# ── 汇总 ────────────────────────────────────────────────────────────
out_path = os.path.join(ROOT, "px_env_report.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print("\n报告已写:", out_path)
