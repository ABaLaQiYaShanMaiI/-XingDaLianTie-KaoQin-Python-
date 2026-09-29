# 跑分绩效汇总自动生成工具

把拍摄的《X月X班考核明细》照片(打印体、手写均可), 用 OCR 解析成「日期 + 姓名 + 事由 + 加减分」,
自动填写《煤库作业区跑分绩效汇总》Excel 考勤表: 当日格记加减分、合计总分、绩效金额
(公式实时计算), 并把明细同步存档到「X班考核明细」工作表。
考勤表的**名单(报名表)可每期导入/替换**(人数可变, 自动增删行并重排公式),
按「姓名 → 行、日期 → 列」定位单元格填数; 也支持直接粘贴文字明细。

> 界面与流程参考《安全生产责任制履职清单考评表自动生成工具》(tkinter + tkinterdnd2,
> 拖拽 / 后台线程 / 一键生成), 本项目为其"跑分考勤"场景的独立实现。

## 快速开始

```bash
pip install -r requirements.txt
python app.py        # 启动图形界面
```

1. 模板文件自动定位「煤库作业区跑分绩效汇总X月（表样）.xlsx」(程序目录 / 参考文件);
2. (可选)「名单(报名表)」页签为各班导入报名表 → 导入即生效并记忆为默认(不导入 = 模板内名单);
3. 把各班「考核明细」照片拖入照片区(每班一张); 也可在「手工录入/粘贴明细」里直接粘文本;
4. 点「识别照片」→ 核对预览结果与警告;
5. 点「一键生成」→ 得到《煤库作业区跑分绩效汇总7月.xlsx》(默认存到**桌面**, 可在「输出目录」改),
   可直接打开/打开所在文件夹。

## 计分规则(已与业务确认)

| 项目 | 规则 |
|------|------|
| 名单(报名表) | 每班可导入报名表/花名册; 不导入则用模板内名单; 导入即记忆为默认, 生成时按报名表重排 B 列名单(人数可变) |
| 定位方式 | 姓名 → 行, 日期 → 列(C~AG 即 1~31 日), 写入该格的加减分 |
| 基础分 | 每人 80 分(界面可改), 体现在合计公式中, 日格不写 80 |
| 日格 | 只写考核明细的加减分(+3/-5…), 同人同日多条自动合并求和 |
| 合计总分 AH | `=80+SUM(C行:AG行)` Excel 公式 |
| 绩效总分 AH末行 | `=SUM(AH3:AH末名单行)` 模板公式, 名单人数变化时程序同步修正范围 |
| 绩效总额 AI末行 | 默认保持模板值(如 10200), 界面可另行指定 |
| 每分价值 AJ末行 | `=AI末行/AH末行` 模板公式 |
| 绩效金额 AI | `=ROUND(AH行*AJ$末行,2)` Excel 公式, 程序不代算数值 |
| 签字确认 AJ | 手签, 程序不填 |

## 目录结构

```
├─ app.py             # 入口 + GUI(基本信息/名单(报名表)/手工录入明细/照片拖拽/识别预览/一键生成)
├─ detail_core.py     # 明细 OCR 解析 + 手工文本明细解析(标题/条目/姓名模糊匹配)
├─ excel_core.py      # 模板名单读取 + 报名表读取 + 填表(名单重排/公式/明细存档)
├─ requirements.txt       # Windows 8.1/10/11 版依赖(最新 Python)
├─ requirements-win7.txt  # Windows 7 版依赖(Python 3.8 + onnxruntime 1.14.1 等)
├─ paofen.spec            # PyInstaller 打包配置(单文件 exe, 最新 Python)
├─ paofen_win7.spec       # PyInstaller 打包配置(Windows 7 兼容版, Python 3.8)
├─ build_exe.ps1          # 一键打包(Windows 10/11 版)
├─ build_exe_win7.ps1     # 一键打包(Windows 7 兼容版, 自动找 Python 3.8)
├─ check_exe_compat.py    # 发布前兼容性自检: 包内二进制是否引用 Win8+ 才有的 API 集
├─ paofen_config.json # 记忆模板/输出目录/基础分/绩效总额/各班报名表(首次运行自动生成)
├─ paofen.log         # 运行日志(排障用)
└─ 参考文件/           # 模板与样例照片(考核明细图片.jpg 等)
```

## 打包为 exe(方便分发)

**先看对方电脑的系统版本, 再选打包脚本** —— 用最新 Python 打的包在 Windows 7 上起不来
(双击就弹「无法启动此程序, 因为计算机中丢失 api-ms-win-core-path-l1-1-0.dll」):

| 对方系统 | 用哪个脚本 | 打包用 Python | 产物 |
|----------|-----------|---------------|------|
| Windows 8.1 / 10 / 11 | `.\build_exe.ps1` | 当前最新(3.13) | `dist\跑分绩效汇总自动生成工具.exe` |
| Windows 7 / 不确定 | `.\build_exe_win7.ps1` | **Python 3.8**(必须) | `dist\跑分绩效汇总自动生成工具_Win7.exe` |

```powershell
.\build_exe.ps1           # Windows 10/11 版(约 110 MB)
.\build_exe_win7.ps1      # Windows 7 兼容版(约 85 MB; 自动找 Python 3.8)
```

- 两个脚本最后一步都会跑 `python check_exe_compat.py <exe>` 做**发布前兼容性自检**:
  展开包内每个 dll/pyd 的 PE 导入表, 列出引用了 Windows 8+ 专有 API 集(`api-ms-win-*`)的二进制,
  并给出「能不能在 Windows 7 启动」的结论 —— 这正是「丢失 xxx.dll」这类事故的根因;
  `build_exe_win7.ps1` 只要自检不通过就**直接失败**, 不会把起不来的包发出去;
- **为什么必须分两个版本**: Python 3.9+ 的运行时 `python3x.dll`、`onnxruntime 1.15+` 的
  `onnxruntime_pybind11_state.pyd` 都硬引用 `api-ms-win-core-path-l1-1-0.dll`(Windows 8+ 才有的
  API 集, Win7 上要额外打 KB2533623), 缺了就直接「无法启动」; Win7 版固定 Python 3.8 +
  `requirements-win7.txt`(`onnxruntime==1.14.1`), 自检确认包内二进制不再引用 Win8+ 的 API 集;
- Win7 版补充: 全新安装的 Windows 7 还缺 UCRT —— 让目标机装一次「VC++ 2015-2022 运行库」
  (KB2999226)即可; 或把 Windows 10 SDK 的 `Redist\ucrt\DLLs\x64\*.dll` 放到 `_runtime\ucrt\x64\`
  再打包, 脚本会自动把它们一起打进 exe(目标机就什么都不用装);
- **VC 运行库会挑 Win7 能用的版本, 并跟着走**: 运行库固化在 `_runtime\vc-win7-x64\`
  (`msvcp140` 14.29 + `vcruntime140` 14.28 + `vcruntime140_1` 14.29, 即 VS2019 版 ——
  最后一个官方支持 Win7 的运行时; 详见该目录 `说明.md`), 打包时由
  `check_exe_compat.add_runtime_beside_binaries` 补到「引用了它的子目录」
  (`onnxruntime\capi\`、`pyclipper\` 等)并校验符号覆盖, 版本 >= 14.40 或覆盖不全就**直接报错**
  (实测本机 System32 的 14.51 那套在 Win7 上就是「找不到指定的程序」);
- **`--probe` 精确定位**: 目标机上执行 `跑分绩效汇总自动生成工具_Win7.exe --probe`, 会按加载器
  顺序逐个核对「包内每个二进制引用的函数在这台机器上是否存在」, 结果写 `deps_probe.txt` ——
  Windows 只会含糊地报「找不到指定的程序」, 这个命令直接告诉你是哪个 DLL 缺哪个函数;
- 仍未自动覆盖的: **全新装的 Win7 缺 UCRT**。让目标机装一次
  `dist\vc_redist_2019_14.29.x64.exe`(VS2019 版, Win7 适用; 也可用 KB2999226);
  或把 Windows 10 SDK 的 `Redist\ucrt\DLLs\x64\*.dll` 放到 `_runtime\ucrt\x64\` 打包时一起带走;
- 产物: **单文件**(内含 OCR 模型、onnxruntime、OpenCV、Python 运行时);
- 分发: 把 exe 拷给同事即可(**无需装 Python**); 把模板 xlsx 放在 exe 同目录,
  程序按 `exe目录 → 上级目录 → exe目录/参考文件` 的顺序自动找到「跑分绩效汇总…xlsx」;
- `参考文件/`(公司模板与样例照片)不打包进 exe;
- 运行后会在 exe 同目录生成: `paofen_config.json`(记忆模板/输出目录/基础分/绩效总额/各班报名表)、
  `paofen.log`; 生成的 Excel 默认存到**桌面**(可在界面改, 记忆上次选择);
- **打包后自检**: `.\dist\跑分绩效汇总自动生成工具.exe --selftest`(Win7 版把文件名换成
  `跑分绩效汇总自动生成工具_Win7.exe`; 退出码 0 = 通过; 结果写入 exe 同目录 `selftest_result.txt`,
  含依赖导入、OCR 模型、模板定位、明细解析与「填表 → 读回校验」; GUI 程序无控制台, 用这种方式排障);
- **对方电脑上「未解析到条目」时**: 在对方电脑执行
  `跑分绩效汇总自动生成工具_Win7.exe --ocr "照片完整路径"`, 会逐阶段(环境版本 / 读图 /
  OCR 引擎初始化 / 模型推理文本框数 / 解析结果)记录到 exe 同目录 `ocr_diag.txt`,
  哪一步是 `[X]`/`[!]` 就是问题所在; 同时 `paofen.log` 里也有识别失败的完整回溯;
- 首次双击若被 Windows SmartScreen/杀软拦(未签名程序), 选「仍要运行」;
  单文件程序首启需解包, 约 2~15 秒(之后启动更快)。

## 测试

```bash
python _test_pipeline.py   # 核心端到端: 手写照片->OCR->填表->逐格校验
python _test_handwriting.py # 手写体准确率: 15 条标准答案 x 正拍/横拍/倒拍/拍虚/换名单
python _test_printed.py    # 打印体端到端: 渲染两种打印版式(合并式/分列式)->OCR->填表->逐格校验
python _test_roster.py     # 报名表端到端: 名单变动(18/20/15人)+文字明细->名单重排/公式/日格校验
python _test_name_split.py # 姓名槽/事由边界: 「陈金荣接」+「班发现…」自动切分、复姓 4 字名不误切
python _test_gui.py        # GUI 全流程: 建界面->照片识别->名单导入->文字明细->两种场景生成
```

> `_test_printed.py` 需 Pillow(仅用于渲染打印体样例图); 打印体样例图输出在 `_out/`。
