param(
    [Parameter(Mandatory)][string]$Presentation,
    [int]$StartSlide = 1,
    [string]$TracePath
)

# Smart Slide Show launcher: opens read-only and resets visited/backward slides cleanly.
$ErrorActionPreference = 'Stop'
$path = (Resolve-Path -LiteralPath $Presentation).Path
$app = New-Object -ComObject PowerPoint.Application
$deck = $app.Presentations.Open($path, -1, 0, -1)
$show = $null

function Trace([string]$Message) {
    if ($TracePath) {
        Add-Content -Encoding utf8 -LiteralPath $TracePath -Value "$(Get-Date -Format o) $Message"
    }
}

try {
    $slideCount = [int]$deck.Slides.Count
    if ($StartSlide -lt 1 -or $StartSlide -gt $slideCount) { $StartSlide = 1 }
    $deck.SlideShowSettings.RangeType = 1
    $show = $deck.SlideShowSettings.Run()
    if ($StartSlide -gt 1) {
        Start-Sleep -Milliseconds 300
        $show.View.GotoSlide($StartSlide, -1)
    }
    $visited = @{}
    $lastPage = -1
    $prevPage = -1

    while ($true) {
        $page = 0
        try {
            $page = [int]$show.View.Slide.SlideIndex
        } catch {
            break
        }
        if ($page -lt 1 -or $page -gt $slideCount) {
            Start-Sleep -Milliseconds 100
            continue
        }
        if ($page -ne $lastPage) {
            $prevPage = $lastPage
            $lastPage = $page
            $isBackwardOrRepeat = $visited.ContainsKey($page) -or ($prevPage -gt 0 -and $page -lt $prevPage)
            $visited[$page] = $true
            Trace "enter page=$page prev=$prevPage reset=$isBackwardOrRepeat"

            $slide = $deck.Slides.Item($page)
            $audio = $null
            foreach ($shape in $slide.Shapes) {
                if ($shape.Name -eq 'AI Teaching Audio') { $audio = $shape; break }
            }
            if ($null -ne $audio) {
                try {
                    if ($isBackwardOrRepeat) {
                        Start-Sleep -Milliseconds 150
                        $show.View.GotoSlide($page, -1)
                        Trace "page=$page reset via GotoSlide(-1)"
                    } else {
                        Start-Sleep -Milliseconds 450
                        $player = $show.View.Player($audio.Id)
                        $pState = [int]$player.State
                        $pPos = [int]$player.CurrentPosition
                        if ($pState -ne 0 -or $pPos -le 0) {
                            $show.View.GotoSlide($page, -1)
                            Start-Sleep -Milliseconds 350
                            $player = $show.View.Player($audio.Id)
                            if ([int]$player.State -ne 0) {
                                $player.CurrentPosition = 0
                                $player.Play()
                                Trace "page=$page fallback Player.Play()"
                            } else {
                                Trace "page=$page recovered via GotoSlide(-1)"
                            }
                        } else {
                            Trace "page=$page auto-started naturally pos=$pPos"
                        }
                    }
                } catch {
                    Trace "page=$page error=$($_.Exception.Message)"
                }
            }
        }
        Start-Sleep -Milliseconds 120
    }
} finally {
    if ($null -ne $show) { try { $show.View.Exit() } catch {} }
    if ($null -ne $deck) { try { $deck.Close() } catch {} }
}
