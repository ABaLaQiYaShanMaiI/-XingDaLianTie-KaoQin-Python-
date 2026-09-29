# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置: 生成单文件 exe(便于分发)。

构建:
    pip install -r requirements.txt pyinstaller
    python -m PyInstaller --noconfirm --clean paofen.spec
产物:
    dist/跑分绩效汇总自动生成工具.exe

要点:
  * RapidOCR 的 onnx 模型/配置、tkinterdnd2 的 tkdnd 库、onnxruntime 的 DLL
    全部由 collect_all 打进 exe(单文件运行时自动解包, 各库自行定位);
  * 公司模板与样例照片(参考文件/)不打包 —— 分发时把模板 xlsx 放在 exe 同目录即可,
    程序按 exe目录 -> 上级目录 -> exe目录/参考文件 的顺序自动寻找「跑分绩效汇总…xlsx」;
  * 配置/日志/生成结果都写在 exe 所在目录(app.py 里按 sys.frozen 取路径);
  * 打包后可用 `跑分绩效汇总自动生成工具.exe --selftest` 自检(结果写入
    exe 同目录的 selftest_result.txt, 退出码 0=通过)。
"""
from PyInstaller.utils.hooks import collect_all

import os
import sys

datas, binaries, hiddenimports = [], [], []
for _pkg in ('rapidocr_onnxruntime', 'tkinterdnd2', 'onnxruntime'):
    _d, _b, _h = collect_all(_pkg)
    datas += _d
    binaries += _b
    hiddenimports += _h

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + ['numpy', 'cv2', 'openpyxl', 'check_exe_compat'],
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib', 'pandas', 'scipy', 'IPython', 'notebook',
              'PyQt5', 'PySide2', 'wx', 'tkinter.test'],
    noarchive=False,
)

# ---- VC 运行库: 补到「引用了它的子目录」里(见 check_exe_compat 里的说明) ----
# 目标机(尤其老系统)system32 里的 MSVCP140.dll 可能过旧, 会让子目录里的
# onnxruntime\capi\*.pyd 报「DLL load failed ... 找不到指定的程序」。
sys.path.insert(0, globals().get('SPECPATH') or os.getcwd())
import check_exe_compat as _cc                              # noqa: E402

_cc.add_runtime_beside_binaries(a.binaries)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='跑分绩效汇总自动生成工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,          # GUI 程序不弹控制台; --selftest 结果写入文件
    disable_windowed_traceback=False,
    icon=None,
)
