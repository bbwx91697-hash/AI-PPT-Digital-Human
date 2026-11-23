$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$cacheRoot = Join-Path $root '.cache'
$env:PIP_CACHE_DIR = Join-Path $cacheRoot 'pip'
$env:TMP = Join-Path $cacheRoot 'tmp'
$env:TEMP = $env:TMP
$env:MODELSCOPE_CACHE = Join-Path $cacheRoot 'modelscope'
New-Item -ItemType Directory -Force -Path $env:PIP_CACHE_DIR, $env:TMP, $env:MODELSCOPE_CACHE | Out-Null

# 选择基础 Python（优先便携包 runtime，其次 .venv，最后系统 python）
$basePy = $null
foreach ($cand in @(
    (Join-Path $root 'runtime\python.exe'),
    (Join-Path $root '.venv\Scripts\python.exe')
)) {
    if ([IO.File]::Exists($cand)) { $basePy = $cand; break }
}
if (-not $basePy -and (Get-Command python -ErrorAction SilentlyContinue)) {
    $basePy = 'python'
}
if (-not $basePy) {
    throw '未找到可用 Python，请先下载便携版或运行「一键安装并启动.bat」。'
}

$qwenVenv = Join-Path $root 'venv_qwen_tts'
$qwenPy = Join-Path $qwenVenv 'Scripts\python.exe'
if (-not [IO.File]::Exists($qwenPy)) {
    Write-Host '[1/4] 正在创建 Qwen3-TTS 专用隔离环境 venv_qwen_tts ...' -ForegroundColor Cyan
    & $basePy -m venv $qwenVenv
    if ($LASTEXITCODE -ne 0 -or -not [IO.File]::Exists($qwenPy)) {
        throw '创建虚拟环境失败，请确认本机已安装完整 Python 3.12。'
    }
}

$hasNvidia = $false
if (Get-Command nvidia-smi -ErrorAction SilentlyContinue) {
    $null = & nvidia-smi 2>$null
    if ($LASTEXITCODE -eq 0) { $hasNvidia = $true }
}

if ($hasNvidia) {
    Write-Host '[2/4] 检测到 NVIDIA 显卡，将安装 CUDA 加速版 PyTorch（约 2.7 GB，使用国内阿里云镜像）...' -ForegroundColor Cyan
    & $qwenPy -m pip install torch==2.9.1+cu128 torchaudio==2.9.1+cu128 -f https://mirrors.aliyun.com/pytorch-wheels/cu128/ --no-warn-script-location
} else {
    Write-Host '[2/4] 未检测到 NVIDIA 独立显卡，将安装轻量化 CPU 版 PyTorch（约 180 MB，可用但生成速度约为音频时长的 3.4 倍）...' -ForegroundColor Yellow
    & $qwenPy -m pip install torch==2.9.1 torchaudio==2.9.1 -i https://pypi.tuna.tsinghua.edu.cn/simple --no-warn-script-location
}
if ($LASTEXITCODE -ne 0) { throw 'PyTorch 安装失败，请检查网络连接。' }

Write-Host '[3/4] 正在安装 Qwen3-TTS 推理依赖（跳过 gradio 冲突包）...' -ForegroundColor Cyan
& $qwenPy -m pip install transformers==4.57.3 accelerate==1.12.0 librosa==1.0.1 soundfile==0.14.0 onnxruntime einops modelscope sox -i https://pypi.tuna.tsinghua.edu.cn/simple --no-warn-script-location
if ($LASTEXITCODE -ne 0) { throw '依赖安装失败，请检查网络连接。' }

& $qwenPy -m pip install qwen-tts==0.1.1 --no-deps -i https://pypi.tuna.tsinghua.edu.cn/simple --no-warn-script-location
if ($LASTEXITCODE -ne 0) { throw 'qwen-tts 安装失败。' }

$modelDir = Join-Path $root 'models\qwen3_tts\Qwen3-TTS-12Hz-0.6B-CustomVoice'
New-Item -ItemType Directory -Force -Path $modelDir | Out-Null
Write-Host '[4/4] 正在通过 ModelScope 国内高速通道下载 Qwen3-TTS-12Hz-0.6B-CustomVoice（约 2.4 GB）...' -ForegroundColor Cyan
@'
import os
from modelscope import snapshot_download
target = os.environ["QWEN3_MODEL_TARGET"]
snapshot_download("Qwen/Qwen3-TTS-12Hz-0.6B-CustomVoice", local_dir=target)
print("MODEL_READY:", target)
'@ | Set-Content -Encoding utf8 (Join-Path $cacheRoot 'dl_model.py')
$env:QWEN3_MODEL_TARGET = $modelDir
& $qwenPy (Join-Path $cacheRoot 'dl_model.py')
if ($LASTEXITCODE -ne 0) { throw 'ModelScope 模型下载失败，请重试。' }

Write-Host '正在验证 Qwen3-TTS 运行状态并清理安装缓存...' -ForegroundColor Cyan
& $qwenPy -c "import torch, qwen_tts; print('CUDA:', torch.cuda.is_available())"
Remove-Item -Recurse -Force $cacheRoot -ErrorAction SilentlyContinue

Write-Host '完成！现在启动软件即可默认使用 Qwen3-TTS 本地高音质模型。' -ForegroundColor Green
Read-Host '按回车键关闭窗口'
