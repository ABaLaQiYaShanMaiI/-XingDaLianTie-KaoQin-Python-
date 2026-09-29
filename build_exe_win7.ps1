# 打包「Windows 7 兼容版」exe: 给老系统/不确定对方系统的同事
#
# 与 build_exe.ps1 的区别: 必须用 Python 3.8(官方最后一个支持 Win7 的版本) + 老依赖;
# 打完自动做兼容性自检, 只要包内还有 Win8+ 才有的 API 集就直接失败, 避免再发出起不来的包。
#
# 用法:
#   .\build_exe_win7.ps1                                   # 自动找 Python 3.8
#   .\build_exe_win7.ps1 -Py 'C:\Python38\python.exe'      # 指定 3.8
#   .\build_exe_win7.ps1 -IndexUrl 'https://pypi.tuna.tsinghua.edu.cn/simple'   # 国内加速
param(
    [string]$Py = '',
    [string]$IndexUrl = ''
)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

# ---- 1) 找 Python 3.8(Win7 版必须用 3.8) ----
function Test-Py38([string]$exe) {
    if (-not $exe) { return $false }
    try {
        $v = & $exe -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
    } catch { return $false }
    return ($LASTEXITCODE -eq 0 -and $v -eq '3.8')
}
$candidates = @()
if ($Py) { $candidates += $Py }
$launcher = @()
try { $launcher = @(& py -3.8 -c "import sys; print(sys.executable)" 2>$null) } catch { }
if ($launcher.Count -gt 0 -and (Test-Path $launcher[0])) { $candidates += $launcher[0] }
$candidates += (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python38\python.exe')
$candidates += 'C:\Python38\python.exe', 'D:\Python38\python.exe'

$python = ''
foreach ($c in $candidates) { if (Test-Py38 $c) { $python = $c; break } }
if (-not $python) {
    throw @"
找不到 Python 3.8 —— Win7 兼容版必须用它打包(3.9+ 的运行时需要 Windows 8+ 才有的 API 集)。
安装(免管理员, 只装到当前用户):
  1) 下载 https://mirrors.huaweicloud.com/python/3.8.10/python-3.8.10-amd64.exe
  2) 运行并勾选 "Add Python 3.8 to PATH"(或自定义安装到默认目录)
  3) 重新执行 .\build_exe_win7.ps1
"@
}
Write-Host "[0/4] 用 Python 3.8 打包: $python" -ForegroundColor Cyan
& $python -c "import sys; print('      ', sys.version.replace(chr(10), ' '))"

# ---- 2) 依赖 & PyInstaller ----
Write-Host '[1/4] 安装 Win7 版依赖(requirements-win7.txt)...'
$pipArgs = @('-m', 'pip', 'install', '-r', 'requirements-win7.txt')
if ($IndexUrl) { $pipArgs += @('-i', $IndexUrl) }
& $python @pipArgs
Write-Host '[2/4] 安装/更新 PyInstaller(3.8 能装的最新版即可)...'
$pipArgs = @('-m', 'pip', 'install', '--upgrade', 'pyinstaller')
if ($IndexUrl) { $pipArgs += @('-i', $IndexUrl) }
& $python @pipArgs

# ---- 3) 打包 ----
Write-Host '[3/4] 开始打包(数分钟, 首次较慢)...'
& $python -m PyInstaller --noconfirm --clean paofen_win7.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller 打包失败(退出码 $LASTEXITCODE), 详见上面的日志(别用旧的 dist 产物顶替)"
}

$exe = Join-Path $PSScriptRoot 'dist\跑分绩效汇总自动生成工具_Win7.exe'
if (-not (Test-Path $exe)) { throw "打包失败: 未生成 $exe" }
$size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host "`n完成: $exe ($size MB)"

# ---- 4) 兼容性自检: 包里只要还有 Win8+ 才有的 API 集就判失败 ----
Write-Host '[4/4] 兼容性自检(能不能在 Windows 7 启动)...'
& $python (Join-Path $PSScriptRoot 'check_exe_compat.py') $exe
if ($LASTEXITCODE -ne 0) {
    throw '兼容性自检未通过: 本包在 Windows 7 上会报「无法启动此程序」, 请先按上面的提示处理。'
}

Write-Host "`n分发(Win7/老系统):" -ForegroundColor Green
Write-Host "  1) 拷这两个文件到对方电脑同一目录: 跑分绩效汇总自动生成工具_Win7.exe + 模板 xlsx"
Write-Host "  2) 若对方是全新 Windows 7 且提示「丢失 api-ms-win-crt-*.dll」, 装一次"
Write-Host "     「VC++ 2015-2022 运行库」(KB2999226) 即可(可把 vc_redist.x64.exe 一起拷过去)"
Write-Host "  3) 对方电脑上先跑一次自检: `"跑分绩效汇总自动生成工具_Win7.exe`" --selftest"
Write-Host "     打开 selftest_result.txt, 全部 [OK] 且「结果: 通过」就说明环境没问题"
