# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置: Windows 7 兼容版(老系统/不确定对方系统时用这份)。

为什么要单独一份配置:
  paofen.spec 用的是 Python 3.13, 产物里的 python313.dll 与 onnxruntime.dll 都引用
  api-ms-win-core-path-l1-1-0.dll —— 那是 Windows 8+ 才有的 API 集, Windows 7 上
  会直接弹「无法启动此程序, 因为计算机中丢失 api-ms-win-core-path-l1-1-0.dll」。
  本配置固定用 Python 3.8 + requirements-win7.txt(onnxruntime 停在 1.14.1),
  实测包内二进制只用 Win7 也有的 API 集, 可以在 Windows 7 SP1 上启动。

用哪份配置:
  * 对方是 Windows 8.1/10/11  → paofen.spec + build_exe.ps1(识别效果最好)
  * 对方是 Windows 7 或不确定 → 本配置 + build_exe_win7.ps1

构建:
    .\\build_exe_win7.ps1
    # 或手动:
    <python3.8>\\python.exe -m PyInstaller --noconfirm --clean paofen_win7.spec
产物:
    dist\\跑分绩效汇总自动生成工具_Win7.exe

要点:
  * 打包后必须跑 `python check_exe_compat.py dist\\...Win7.exe` 复核(脚本已自动做),
    只要它报出「非 UCRT 的 API 集」, 这个包在老系统上就一定起不来;
  * 全新安装的 Windows 7 还缺 UCRT: 让目标机装「VC++ 2015-2022 运行库」(KB2999226),
    或把 Windows 10 SDK 的 Redist\\ucrt\\DLLs\\x64\\*.dll 放到 _runtime\\ucrt\\x64\\
    (或 参考文件\\ucrt\\x64\\) 再打包, 本配置会自动把它们一起打进 exe;
  * pdfminer/cryptography 与拍照识别无关(pdfminer 依赖 Rust 版 cryptography,
    它引用 Win8+ 的 api-ms-win-core-synch-l1-2-0.dll), 这里显式排除, 免得被别的
    包间接拉进来污染产物。
"""
import os
import sys

from PyInstaller.utils.hooks import collect_all

if sys.version_info[:2] != (3, 8):
    print('[!] paofen_win7.spec: 当前解释器是 Python %d.%d, Win7 兼容版必须用 Python 3.8 '
          '打包(3.9+ 的运行时本身就要 api-ms-win-core-path-l1-1-0.dll)!'
          % sys.version_info[:2])

datas, binaries, hiddenimports = [], [], []
for _pkg in ('rapidocr_onnxruntime', 'tkinterdnd2', 'onnxruntime'):
    _d, _b, _h = collect_all(_pkg)
    datas += _d
    binaries += _b
    hiddenimports += _h

# ---- UCRT(可选): 找得到就一起打进 exe, 目标机就不用装运行库了 ----
_SPEC_DIR = globals().get('SPECPATH') or os.getcwd()
_ucrt_dirs = [
    os.path.join(_SPEC_DIR, '_runtime', 'ucrt', 'x64'),
    os.path.join(_SPEC_DIR, '参考文件', 'ucrt', 'x64'),
    r'C:\Program Files (x86)\Windows Kits\10\Redist\ucrt\DLLs\x64',
]
_ucrt_from = None
for _dir in _ucrt_dirs:
    if os.path.isdir(_dir) and any(f.lower().endswith('.dll') for f in os.listdir(_dir)):
        _ucrt_from = _dir
        for _fn in os.listdir(_dir):
            if _fn.lower().endswith('.dll'):
                binaries.append((os.path.join(_dir, _fn), '.'))
        break
if _ucrt_from:
    print('[*] 已打入 UCRT 文件(来自 %s)' % _ucrt_from)
else:
    print('[!] 未找到 UCRT 重分发文件, 目标机若是全新 Windows 7:')
    print('    请先安装「VC++ 2015-2022 运行库」(KB2999226), 否则会提示')
    print('    「丢失 api-ms-win-crt-runtime-l1-1-0.dll」; 或者把 Windows 10 SDK 的')
    print(r'    Redist\ucrt\DLLs\x64\*.dll 放到 _runtime\ucrt\x64\ 后重新打包。')

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + ['numpy', 'cv2', 'openpyxl', 'check_exe_compat'],
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib', 'pandas', 'scipy', 'IPython', 'notebook',
              'PyQt5', 'PySide2', 'wx', 'tkinter.test',
              'pdfminer', 'cryptography'],   # 拍照识别用不到, 且会引入 Win8+ API 集
    noarchive=False,
)

# ---- 把 VC++ 运行库复制到「真正需要它的二进制」所在目录 ----
# 为什么: onnxruntime\capi\onnxruntime_pybind11_state.pyd 在子目录里, Windows 加载它时
# 先找"它自己所在目录"、再找 system32, 只放在包根目录的 MSVCP140.dll 根本轮不到 ——
# 于是用目标机的 C:\Windows\System32\MSVCP140.dll。老系统(只装过 VC2015/2017 运行库)
# 那份太旧, 缺 ORT 需要的新符号, 就报
# 「ImportError: DLL load failed while importing onnxruntime_pybind11_state: 找不到指定的程序。」
# (实盘 Win7 事故, 处理函数见 check_exe_compat.add_runtime_beside_binaries)
sys.path.insert(0, _SPEC_DIR)
import check_exe_compat as _cc              # noqa: E402  同目录小工具(纯标准库)

# 优先用项目自带的 _runtime\vc-win7-x64(见该目录说明.md); 要求必须是 Win7 能用的版本,
# 否则直接报错(宁可打包失败, 也别再发出对方起不来的包)。
_cc.add_runtime_beside_binaries(
    a.binaries,
    extra_dirs=[os.path.join(_SPEC_DIR, '_runtime', 'vc-win7-x64')],
    require_win7=True)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='跑分绩效汇总自动生成工具_Win7',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,          # GUI 程序不弹控制台; --selftest 结果写入文件
    disable_windowed_traceback=False,
    icon=None,
)
