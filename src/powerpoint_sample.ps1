param(
    [Parameter(Mandatory)][ValidateSet('alpha','auto','manual')][string]$Mode,
    [switch]$Existing
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$source = Join-Path $root 'input/说课.pptx'
$preview = Join-Path $root "generated/previews/$Mode"
$videos = Join-Path $root 'generated/videos'
New-Item -ItemType Directory -Force -Path $preview | Out-Null

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ShowWindow {
    [StructLayout(LayoutKind.Sequential)]
    public struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out Rect r);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, IntPtr name);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint x, uint y, uint data, UIntPtr extra);
}
'@
[ShowWindow]::SetProcessDPIAware() | Out-Null

function Get-ShowBounds {
    $window = [ShowWindow]::FindWindow('screenClass', [IntPtr]::Zero)
    $bounds = [ShowWindow+Rect]::new()
    if (-not [ShowWindow]::GetWindowRect($window, [ref]$bounds)) {
        throw 'Unable to get the PowerPoint slide show window bounds.'
    }
    return $bounds
}

function Save-Capture($Show, [string]$Name) {
    $b = Get-ShowBounds
    $width = $b.Right - $b.Left
    $height = $b.Bottom - $b.Top
    $bitmap = [System.Drawing.Bitmap]::new($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($b.Left, $b.Top, 0, 0, $bitmap.Size)
        $bitmap.Save((Join-Path $preview "$Name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Click-Video($Show) {
    $b = Get-ShowBounds
    [ShowWindow]::SetForegroundWindow([ShowWindow]::FindWindow('screenClass', [IntPtr]::Zero)) | Out-Null
    [ShowWindow]::SetCursorPos([int](($b.Left + $b.Right) / 2), [int](($b.Top + $b.Bottom) / 2)) | Out-Null
    Start-Sleep -Milliseconds 250
    [ShowWindow]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero)
    [ShowWindow]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
}

function Measure-Playback($Show, [int]$ShapeId, [string]$Label, [int]$Count = 2, [int]$Interval = 600) {
    $player = $Show.View.Player($ShapeId)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $entries = @(1..$Count | ForEach-Object {
        Start-Sleep -Milliseconds $Interval
        [ordered]@{
            slide = $Show.View.Slide.SlideIndex
            state = [int]$player.State
            position_ms = [int]$player.CurrentPosition
            elapsed_ms = [int]$timer.ElapsedMilliseconds
        }
    })
    [ordered]@{ action = $Label; samples = $entries }
}

function Set-VideoTrigger($Slide, $Media, [bool]$Automatic) {
    $main = $Slide.TimeLine.MainSequence
    for ($i = $main.Count; $i -ge 1; $i--) { $main.Item($i).Delete() }
    if ($Automatic) {
        $null = $main.AddEffect($Media, 83, 0, 2)
    } else {
        $null = $Slide.TimeLine.InteractiveSequences.Add(1).AddTriggerEffect($Media, 83, 4, $Media, '', 0)
    }
}

$app = New-Object -ComObject PowerPoint.Application
$deck = $null
$show = $null
try {
    if ($Mode -eq 'alpha') {
        $deck = $app.Presentations.Add(-1)
        $deck.PageSetup.SlideWidth = 960
        $deck.PageSetup.SlideHeight = 540
        $slide = $deck.Slides.Add(1, 12)
        $background = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
        $background.Fill.ForeColor.RGB = 14083539
        $background.Line.Visible = 0
        $bar = $slide.Shapes.AddShape(1, 0, 215, 960, 45)
        $bar.Fill.ForeColor.RGB = 5383059
        $bar.Line.Visible = 0
        $clip = Join-Path $videos 'avatar-alpha.webm'
        Write-Host 'PowerPoint: inserting transparent WebM.'
        $media = $slide.Shapes.AddMediaObject2($clip, 0, -1, 380, 145, 200, 250)
        Write-Host 'PowerPoint: starting alpha slideshow.'
        Set-VideoTrigger $slide $media $true
        $show = $deck.SlideShowSettings.Run()
        $rows = Measure-Playback $show $media.Id 'alpha_auto_entry'
        $rows | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $root 'reports/playback-alpha.json')
        $rows | ConvertTo-Json -Depth 5
        Save-Capture $show 'powerpoint-alpha'
        return
    }

    $suffix = if ($Mode -eq 'manual') { '_手动' } else { '' }
    $destination = Join-Path $root "output/说课_数字人提问样片$suffix.pptx"
    if (-not $Existing) {
        $stage = Join-Path $root "generated/.work/sample-$Mode.pptx"
        [IO.File]::Copy($source, $stage, $true)
        $deck = $app.Presentations.Open($stage, 0, 0, 0)
        if ($deck.Slides.Count -ne 16) { throw 'Unexpected input slide count.' }
        $slide = $deck.Slides.Add(7, 12)
        $slide.DisplayMasterShapes = 0
        $slide.SlideShowTransition.AdvanceOnTime = 0
        $slide.SlideShowTransition.AdvanceOnClick = 0
        $clip = Join-Path $videos 'slide07-composite.mp4'
        $media = $slide.Shapes.AddMediaObject2($clip, 0, -1, 0, 0, 960, 540)
        $media.Name = 'Digital avatar playback test'
        $media.AlternativeText = '静音数字人占位视频，声音待接入'
        $media.MediaFormat.SetDisplayPicture(0)
        $play = $media.AnimationSettings.PlaySettings
        $play.LoopUntilStopped = 0
        $play.RewindMovie = -1
        Set-VideoTrigger $slide $media ($Mode -eq 'auto')
        $deck.Save()
        $deck.Close()
        python (Join-Path $PSScriptRoot 'preserve_original.py') $source $stage $destination
        if ($LASTEXITCODE -ne 0) { throw 'Could not merge the new sample slide into the original package.' }
        [IO.File]::Delete($stage)
    }
    $deck = $app.Presentations.Open($destination, -1, 0, -1)
    foreach ($page in $deck.Slides) {
        $page.Export((Join-Path $preview ('slide-{0:D2}.png' -f $page.SlideIndex)), 'PNG', 1920, 1080)
    }
    $settings = $deck.SlideShowSettings
    $settings.ShowType = 1
    $settings.RangeType = 2
    $settings.StartingSlide = 6
    $settings.EndingSlide = 8
    $show = $settings.Run()
    $showWindow = [ShowWindow]::FindWindow('screenClass', [IntPtr]::Zero)
    [ShowWindow]::SetForegroundWindow($showWindow) | Out-Null
    Start-Sleep -Milliseconds 1400
    [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}')
    Start-Sleep -Milliseconds 300
    if ($show.View.Slide.SlideIndex -ne 7) { throw 'Playback did not reach the sample slide.' }
    $mediaId = $deck.Slides.Item(7).Shapes.Item(1).Id
    $rows = if ($Mode -eq 'auto') {
        @(Measure-Playback $show $mediaId 'entry' 13 300)
    } else {
        @(Measure-Playback $show $mediaId 'entry')
    }
    $rows = @($rows)
    if ($Mode -eq 'manual') {
        try { Save-Capture $show 'powerpoint-before-click' } catch { Write-Warning $_.Exception.Message }
        try {
            Click-Video $show
            $rows += Measure-Playback $show $mediaId 'clicked_video'
        } catch { Write-Warning $_.Exception.Message }
    }
    try { Save-Capture $show 'powerpoint-playing' } catch { Write-Warning $_.Exception.Message }
    Start-Sleep -Seconds 6
    $rows += Measure-Playback $show $mediaId 'after_end'
    try {
        $window = [ShowWindow]::FindWindow('screenClass', [IntPtr]::Zero)
        [ShowWindow]::SetForegroundWindow($window) | Out-Null
        $bounds = Get-ShowBounds
        [ShowWindow]::SetCursorPos(
            [int]($bounds.Left + 0.8 * ($bounds.Right - $bounds.Left)),
            [int]($bounds.Top + 0.8 * ($bounds.Bottom - $bounds.Top))
        ) | Out-Null
        Start-Sleep -Milliseconds 350
        [System.Windows.Forms.SendKeys]::SendWait('%p')
        $replay = Measure-Playback $show $mediaId 'replay_alt_p'
        $rows += $replay
        if ($replay.samples[1].state -ne 0 -and $Mode -eq 'auto') {
            Click-Video $show
            $rows += Measure-Playback $show $mediaId 'replay_clicked_video'
        }
    } catch { Write-Warning $_.Exception.Message }
    $rows | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $root "reports/playback-$Mode.json")
    $rows | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $show) {
        try { $show.View.Exit() } catch { Write-Warning $_.Exception.Message }
    }
    if ($null -ne $deck) {
        try {
            if ($Mode -eq 'alpha') { $deck.Saved = -1 }
            $deck.Close()
        } catch { Write-Warning $_.Exception.Message }
    }
}
