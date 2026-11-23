param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Destination,
    [string]$Avatar = '',
    [string]$IdleAvatar = '',
    [Parameter(Mandatory)][string]$PagesJson,
    [ValidateSet('auto','manual')][string]$Mode = 'manual',
    [switch]$ExperimentalIdle,
    [switch]$PreserveAnimations,
    [string]$PythonExe = ''
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not $Avatar) { $Avatar = Join-Path $projectRoot 'assets/avatar/default_avatar.gif' }
if (-not $IdleAvatar) { $IdleAvatar = Join-Path $projectRoot 'assets/avatar/default_avatar_idle.png' }
$sourcePath = (Resolve-Path -LiteralPath $Source).Path
$avatarPath = (Resolve-Path -LiteralPath $Avatar).Path
$destinationPath = [IO.Path]::GetFullPath($Destination)
if ($sourcePath -eq $destinationPath) { throw '输出路径必须不同于源文件。' }
if ([IO.File]::Exists($destinationPath)) { throw '输出文件已存在，请选择新文件名。' }
$parsed = Get-Content -Raw -Encoding utf8 -LiteralPath $PagesJson | ConvertFrom-Json
$pages = if ($parsed -is [System.Array]) { $parsed } else { @($parsed) }
if ($pages.Count -eq 0) { throw '请选择至少一页。' }
$pageNumbers = @{}
foreach ($page in $pages) {
    $number = [int]$page.slide
    if ($number -lt 1 -or $pageNumbers.ContainsKey($number)) { throw '页码无效或重复。' }
    $pageNumbers[$number] = $true
    $page.audio = (Resolve-Path -LiteralPath $page.audio).Path
}
$tempPath = Join-Path ([IO.Path]::GetDirectoryName($destinationPath)) ('.' + [guid]::NewGuid().ToString() + '.pptx')
$app = New-Object -ComObject PowerPoint.Application
$deck = $null
try {
    [IO.File]::Copy($sourcePath, $tempPath)
    $deck = $app.Presentations.Open($tempPath, 0, 0, 0)
    foreach ($page in $pages) {
        if ([int]$page.slide -gt $deck.Slides.Count) { throw "页码超出范围: $($page.slide)" }
    }
    $width = $deck.PageSetup.SlideWidth
    $height = $deck.PageSetup.SlideHeight
    foreach ($page in $pages) {
        $slide = $deck.Slides.Item([int]$page.slide)
        foreach ($name in @('AI Teaching Audio', 'AI Teaching Avatar', 'AI Teaching Idle')) {
            for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
                if ($slide.Shapes.Item($i).Name -eq $name) { $slide.Shapes.Item($i).Delete() }
            }
        }
        $left = if ($null -ne $page.left) { [double]$page.left } else { $width * 0.81 }
        $top = if ($null -ne $page.top) { [double]$page.top } else { $height * 0.60 }
        $gifWidth = if ($null -ne $page.width) { [double]$page.width } else { $width * 0.16 }
        $gifHeight = if ($null -ne $page.height) { [double]$page.height } else { $height * 0.35 }
        if ($left -lt 0 -or $top -lt 0 -or $gifWidth -le 0 -or $gifHeight -le 0 -or
            $left + $gifWidth -gt $width -or $top + $gifHeight -gt $height) {
            throw "第 $($page.slide) 页人物位置超出幻灯片范围。"
        }
        $gif = $slide.Shapes.AddPicture($avatarPath, 0, -1, $left, $top, $gifWidth, $gifHeight)
        $gif.Name = 'AI Teaching Avatar'
        $audio = $slide.Shapes.AddMediaObject2($page.audio, 0, -1, 0, 0, 1, 1)
        $audio.Name = 'AI Teaching Audio'
        $audio.AnimationSettings.PlaySettings.HideWhileNotPlaying = -1
        $audio.AnimationSettings.PlaySettings.StopAfterSlides = 1
        $audio.AnimationSettings.PlaySettings.RewindMovie = -1
        $main = $slide.TimeLine.MainSequence
        for ($i = $main.Count; $i -ge 1; $i--) {
            if ($main.Item($i).Shape.Id -eq $audio.Id) { $main.Item($i).Delete() }
        }
        if ($ExperimentalIdle) {
            $idlePath = (Resolve-Path -LiteralPath $IdleAvatar).Path
            $idle = $slide.Shapes.AddPicture($idlePath, 0, -1, $left, $top, $gifWidth, $gifHeight)
            $idle.Name = 'AI Teaching Idle'
        }
        if ($Mode -eq 'auto') {
            $null = $main.AddEffect($audio, 83, 0, 2)
            if ($ExperimentalIdle) {
                $null = $main.AddEffect($idle, 1, 0, 3)
            }
        } else {
            $sequence = $slide.TimeLine.InteractiveSequences.Add(1)
            $null = $sequence.AddTriggerEffect($audio, 83, 4, $gif, '', 0)
            if ($ExperimentalIdle) {
                $null = $sequence.AddEffect($idle, 1, 0, 3)
                $replay = $slide.TimeLine.InteractiveSequences.Add(2)
                $null = $replay.AddTriggerEffect($audio, 83, 4, $idle, '', 0)
                $hideIdle = $replay.AddEffect($idle, 1, 0, 2)
                $hideIdle.Exit = -1
                $null = $replay.AddEffect($idle, 1, 0, 3)
            }
        }
    }
    $deck.Save()
    $deck.Close()
    $deck = $null
    $mergeArgs = @((Join-Path $PSScriptRoot 'openxml_merge.py'), $sourcePath, $tempPath, $destinationPath, $PagesJson, '--mode', $Mode)
    if ($PreserveAnimations) { $mergeArgs += '--preserve-animations' }
    $pyCmd = $PythonExe
    if (-not $pyCmd) {
        $bundledPy = Join-Path $projectRoot 'runtime\python.exe'
        $venvPy = Join-Path $projectRoot '.venv\Scripts\python.exe'
        if ([IO.File]::Exists($bundledPy)) { $pyCmd = $bundledPy }
        elseif ([IO.File]::Exists($venvPy)) { $pyCmd = $venvPy }
        else { $pyCmd = 'python' }
    }
    & $pyCmd @mergeArgs
    if ($LASTEXITCODE -ne 0) { throw 'Open XML 定点迁移失败，未完成输出。' }
    Write-Output "保存完成: $destinationPath"
} finally {
    if ($null -ne $deck) { $deck.Close() }
    if ([IO.File]::Exists($tempPath)) { [IO.File]::Delete($tempPath) }
}
