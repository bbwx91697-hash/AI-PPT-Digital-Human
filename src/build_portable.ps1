$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$pkgName = 'AI-PPT-Digital-Human'
$distRoot = Join-Path $root 'dist'
$stage = Join-Path $distRoot $pkgName
$zipPath = Join-Path $distRoot 'AI-PPT-Digital-Human-Win64-Portable.zip'

if (Test-Path $stage) { Remove-Item -Recurse -Force $stage }
if (Test-Path $zipPath) { Remove-Item -Force $zipPath }
New-Item -ItemType Directory -Force -Path $stage, (Join-Path $stage 'bin'), (Join-Path $stage 'runtime') | Out-Null

# 1. 复制核心源码、默认数字人素材、启动与安装脚本
foreach ($dir in @('src', 'assets')) {
    $srcDir = Join-Path $root $dir
    $dstDir = Join-Path $stage $dir
    $null = robocopy $srcDir $dstDir /E /XD __pycache__ /XF *.pyc *.pyo /NFL /NDL /NJH /NJS /NC /NS /NP
    if ($LASTEXITCODE -ge 8) { throw "复制目录失败: $dir ($LASTEXITCODE)" }
}
foreach ($file in @('README.md', 'requirements.txt', '启动数字人PPT助手.bat', '一键安装并启动.bat', '一键安装本地Qwen3高音质模型.bat')) {
    $fp = Join-Path $root $file
    if (Test-Path $fp) { Copy-Item $fp (Join-Path $stage $file) -Force }
}

# 2. 打包 FFmpeg
$ffmpegSrc = 'D:\ffmpeg\bin\ffmpeg.exe'
if (-not (Test-Path $ffmpegSrc)) {
    $cmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($cmd) { $ffmpegSrc = $cmd.Source }
}
if (-not (Test-Path $ffmpegSrc)) { throw '未找到 ffmpeg.exe，无法构建自带 FFmpeg 的便携包。' }
Copy-Item $ffmpegSrc (Join-Path $stage 'bin\ffmpeg.exe') -Force

# 3. 构建独立精简 Python 3.12 运行时（含 tkinter 与 venv/ensurepip，不依赖本机全局 Python）
$pyRoot = Split-Path -Parent (Get-Command python).Source
$rtDir = Join-Path $stage 'runtime'
foreach ($topFile in @('python.exe', 'pythonw.exe', 'python3.dll', 'python312.dll', 'vcruntime140.dll', 'vcruntime140_1.dll', 'LICENSE.txt')) {
    $p = Join-Path $pyRoot $topFile
    if (Test-Path $p) { Copy-Item $p (Join-Path $rtDir $topFile) -Force }
}
foreach ($sub in @('DLLs', 'tcl')) {
    $null = robocopy (Join-Path $pyRoot $sub) (Join-Path $rtDir $sub) /E /NFL /NDL /NJH /NJS /NC /NS /NP
    if ($LASTEXITCODE -ge 8) { throw "复制运行时目录失败: $sub ($LASTEXITCODE)" }
}
$null = robocopy (Join-Path $pyRoot 'Lib') (Join-Path $rtDir 'Lib') /E /XD test idlelib site-packages __pycache__ /XF *.pyc *.pyo /NFL /NDL /NJH /NJS /NC /NS /NP
if ($LASTEXITCODE -ge 8) { throw "复制标准库失败 ($LASTEXITCODE)" }
New-Item -ItemType Directory -Force -Path (Join-Path $rtDir 'Lib\site-packages') | Out-Null

$rtPy = Join-Path $rtDir 'python.exe'
$env:PYTHONPATH = ''
$env:PYTHONNOUSERSITE = '1'
$env:PIP_CACHE_DIR = Join-Path $root '.cache\pip'
& $rtPy -m ensurepip --upgrade
if ($LASTEXITCODE -ne 0) { throw 'ensurepip 初始化失败' }
& $rtPy -m pip install pillow edge-tts -i https://pypi.tuna.tsinghua.edu.cn/simple --no-warn-script-location
if ($LASTEXITCODE -ne 0) { throw '安装便携运行依赖失败' }

# 清理 __pycache__ 以减小压缩包体积
Get-ChildItem $stage -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "正在压缩便携包: $zipPath ..." -ForegroundColor Cyan
Compress-Archive -Path $stage -DestinationPath $zipPath -CompressionLevel Optimal
$sizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host "构建完成: $zipPath ($sizeMB MB)" -ForegroundColor Green
