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
    $routes = @(
        @(6, 7, 8), @(9, 8), @(8, 7, 8),
        @(7, 8, 9, 8), @(8, 7, 8, 7, 8), @(7, 8, 7)
    )
    $results = [Collections.Generic.List[object]]::new()
    foreach ($route in $routes) {
        foreach ($page in $route) {
            $show.View.GotoSlide($page, 0)
            if ($show.View.Slide.SlideIndex -ne $page) { throw "Navigation failed on page $page" }
            if ($page -notin 7,8) { continue }
            $id = $deck.Slides.Item($page).Shapes.Item('AI Teaching Audio').Id
            $player = $show.View.Player($id)
            $player.CurrentPosition = 0
            $player.Play()
            Start-Sleep -Milliseconds 600
            $first = [int]$player.CurrentPosition
            $firstState = [int]$player.State
            Start-Sleep -Milliseconds 600
            $second = [int]$player.CurrentPosition
            $secondState = [int]$player.State
            $results.Add([ordered]@{
                route = ($route -join '→')
                slide = $page
                state1 = $firstState
                position1 = $first
                state2 = $secondState
                position2 = $second
                advanced = ($firstState -eq 0 -and $secondState -eq 0 -and $second -gt $first)
            })
            $results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $LogPath
        }
    }
    $results | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $show) { $show.View.Exit() }
    $deck.Close()
}
