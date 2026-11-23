$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$runtimePyW = Join-Path $root 'runtime\pythonw.exe'
if ([IO.File]::Exists($runtimePyW)) {
    Start-Process -FilePath $runtimePyW -ArgumentList ('"' + (Join-Path $root 'src\app.py') + '"')
    exit 0
}

$venvPy = Join-Path $root '.venv\Scripts\python.exe'
$venvPyW = Join-Path $root '.venv\Scripts\pythonw.exe'

if (-not [IO.File]::Exists($venvPy)) {
    $sysPy = Get-Command python -ErrorAction SilentlyContinue
    if (-not $sysPy) {
        Write-Host '[1/3] 未检测到 Python，尝试通过 winget 安装 Python 3.12...' -ForegroundColor Yellow
        $wg = Get-Command winget -ErrorAction SilentlyContinue
        if ($wg) {
            winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
        }
        $sysPy = Get-Command python -ErrorAction SilentlyContinue
        if (-not $sysPy) {
            throw '未能自动找到 Python 3.10~3.12，请先从官网或微软商店安装 Python 3.12 后重试。'
        }
    }
    Write-Host '[1/3] 正在创建独立运行环境 .venv ...' -ForegroundColor Cyan
    & python -m venv (Join-Path $root '.venv')
}

Write-Host '[2/3] 正在通过清华镜像安装基础依赖 (Pillow / edge-tts / imageio-ffmpeg)...' -ForegroundColor Cyan
& $venvPy -m pip install -r (Join-Path $root 'requirements.txt') -i https://pypi.tuna.tsinghua.edu.cn/simple
if ($LASTEXITCODE -ne 0) { throw '基础依赖安装失败，请检查网络连接。' }

Write-Host '[3/3] 安装完成，正在启动 AI PPT 数字人教学助手...' -ForegroundColor Green
Start-Process -FilePath $venvPyW -ArgumentList ('"' + (Join-Path $root 'src\app.py') + '"')
