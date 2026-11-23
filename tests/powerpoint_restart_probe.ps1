param(
    [Parameter(Mandatory)][string]$Presentation,
    [Parameter(Mandatory)][string]$LogPath
)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$deck = $app.Presentations.Open((Resolve-Path -LiteralPath $Presentation).Path, -1, 0, -1)
$show = $null
try {
    $show = $deck.SlideShowSettings.Run()
    $log = [Collections.Generic.List[object]]::new()
    function Sample([string]$label, [int]$expected) {
        Start-Sleep -Milliseconds 1800
        $page = $show.View.Slide.SlideIndex
        if ($page -ne $expected) { throw "$label reached page $page rather than $expected" }
        $player = $show.View.Player($deck.Slides.Item($page).Shapes.Item('AI Teaching Audio').Id)
        $state1 = [int]$player.State
        $pos1 = [int]$player.CurrentPosition
        Start-Sleep -Milliseconds 900
        $state2 = [int]$player.State
        $pos2 = [int]$player.CurrentPosition
        $log.Add([ordered]@{action=$label;slide=$page;state1=$state1;position1=$pos1;state2=$state2;position2=$pos2})
        $log | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $LogPath
        Write-Output "$label slide=$page state=$state1,$state2 position=$pos1,$pos2"
    }
    $show.View.GotoSlide(7, -1)
    Sample 'initial explicit reset to 7' 7
    Start-Sleep -Seconds 7
    $show.View.GotoSlide(8, -1)
    Sample 'explicit reset to 8' 8
    Start-Sleep -Seconds 7
    $show.View.GotoSlide(9, -1)
    $show.View.Previous()
    Sample '9 previous to 8' 8
    $show.View.Previous()
    if ($show.View.Slide.SlideIndex -eq 8) {
        Write-Output 'Previous first stepped within the slide animation.'
        $show.View.Previous()
    }
    Sample '8 previous to 7' 7
    $show.View.Next()
    Sample '7 next to 8' 8
    $log | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $LogPath
    $log | ConvertTo-Json -Depth 4
} finally {
    if ($null -ne $show) { $show.View.Exit() }
    $deck.Close()
}
