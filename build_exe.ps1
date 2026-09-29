# 打包单文件 exe(便于分发): 右键「使用 PowerShell 运行」或命令行 .\build_exe.ps1
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

Write-Host '[1/3] 安装运行依赖...'
python -m pip install -r requirements.txt
Write-Host '[2/3] 安装/更新 PyInstaller...'
python -m pip install --upgrade pyinstaller
Write-Host '[3/3] 开始打包(数分钟, 首次较慢)...'
python -m PyInstaller --noconfirm --clean paofen.spec

$exe = Join-Path $PSScriptRoot 'dist\跑分绩效汇总自动生成工具.exe'
if (-not (Test-Path $exe)) { throw "打包失败: 未生成 $exe" }
$size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host "`n完成: $exe ($size MB)"
Write-Host "自检: `"$exe`" --selftest   # 结果见 exe 同目录 selftest_result.txt"
Write-Host '提示: 分发时把模板 xlsx 与 exe 放在同一目录, 程序会自动找到它。'
