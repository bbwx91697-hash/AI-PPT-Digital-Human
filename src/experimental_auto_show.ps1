param(
    [Parameter(Mandatory)][string]$Presentation,
    [string]$TracePath
)

# Experimental controller. It opens the deck read-only and never modifies it.
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
    $show = $deck.SlideShowSettings.Run()
    $lastPage = -1
    $started = $false
    while ($true) {
        try {
            $page = [int]$show.View.Slide.SlideIndex
        } catch {
            break
        }
        if ($page -ne $lastPage) {
            $lastPage = $page
            $started = $false
            Trace "enter page=$page"
        }
        if (-not $started) {
            $slide = $deck.Slides.Item($page)
            $audio = $null
            foreach ($shape in $slide.Shapes) {
                if ($shape.Name -eq 'AI Teaching Audio') { $audio = $shape; break }
            }
            if ($null -eq $audio) {
                $started = $true
            } else {
                try {
                    $player = $show.View.Player($audio.Id)
                    $state = [int]$player.State
                    if ($state -eq 0) {
                        $started = $true
                        Trace "page=$page already playing"
                    } elseif ($state -eq 2) {
                        $player.CurrentPosition = 0
                        $player.Play()
                        $started = $true
                        Trace "page=$page commanded start"
                    } else {
                        Trace "page=$page state=$state awaiting"
                    }
                } catch {
                    Trace "page=$page error=$($_.Exception.Message)"
                }
            }
        }
        Start-Sleep -Milliseconds 150
    }
} finally {
    if ($null -ne $show) {
        try { $show.View.Exit() } catch {}
    }
    if ($null -ne $deck) { try { $deck.Close() } catch {} }
}
