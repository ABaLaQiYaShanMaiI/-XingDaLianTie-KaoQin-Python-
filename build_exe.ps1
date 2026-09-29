# 打包单文件 exe(便于分发): 右键「使用 PowerShell 运行」或命令行 .\build_exe.ps1
#
# 本脚本用当前最新 Python 打包, 产物只保证 Windows 8.1/10/11 可用。
# 对方是 Windows 7 / 不确定系统时请改用 .\build_exe_win7.ps1, 否则会报
# 「无法启动此程序, 因为计算机中丢失 api-ms-win-core-path-l1-1-0.dll」。
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

Write-Host '[0/4] 检查打包环境...'
$pyver = (python -c "import sys; print('%d.%d.%d' % sys.version_info[:3])").Trim()
if ($pyver -notmatch '^\d+\.\d+') { throw '找不到 python, 请先安装 Python 并加入 PATH' }
Write-Host "      打包用 Python $pyver"
if ([version]$pyver -ge [version]'3.9') {
    Write-Host '      [!] 3.9+ 打的包最低要求 Windows 8.1 —— Windows 7 会报「丢失 api-ms-win-core-path-l1-1-0.dll」' -ForegroundColor Yellow
    Write-Host '      [!] 要给 Windows 7 / 老系统用, 请执行: .\build_exe_win7.ps1' -ForegroundColor Yellow
} else {
    Write-Host '      Python 3.8: 产物可在 Windows 7 SP1 上运行(目标机需有 UCRT / VC++ 运行库)'
}

Write-Host '[1/4] 安装运行依赖...'
python -m pip install -r requirements.txt
Write-Host '[2/4] 安装/更新 PyInstaller...'
python -m pip install --upgrade pyinstaller
Write-Host '[3/4] 开始打包(数分钟, 首次较慢)...'
python -m PyInstaller --noconfirm --clean paofen.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller 打包失败(退出码 $LASTEXITCODE), 详见上面的日志(别用旧的 dist 产物顶替)"
}

$exe = Join-Path $PSScriptRoot 'dist\跑分绩效汇总自动生成工具.exe'
if (-not (Test-Path $exe)) { throw "打包失败: 未生成 $exe" }
$size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host "`n完成: $exe ($size MB)"

Write-Host '[4/4] 发布前兼容性自检(这个包能不能在 Windows 7 上启动)...'
python (Join-Path $PSScriptRoot 'check_exe_compat.py') $exe
if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[!] 本包不兼容 Windows 7(上面已列出是哪些二进制、缺哪个 DLL)。" -ForegroundColor Yellow
    Write-Host '    对方是老系统 / 不确定系统时, 请改用: .\build_exe_win7.ps1' -ForegroundColor Yellow
}

Write-Host "`n自检: `"$exe`" --selftest   # 结果见 exe 同目录 selftest_result.txt"
Write-Host '提示: 分发时把模板 xlsx 与 exe 放在同一目录, 程序会自动找到它。'
