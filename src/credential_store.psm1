function Save-GeminiKey {
    param(
        [Parameter(Mandatory)][Security.SecureString]$Key,
        [Parameter(Mandatory)][string]$Path
    )
    $directory = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Path))
    if (-not [IO.Directory]::Exists($directory)) {
        [IO.Directory]::CreateDirectory($directory) | Out-Null
    }
    $encoded = ConvertFrom-SecureString -SecureString $Key
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($Path), $encoded)
}

function Get-GeminiKey {
    param([Parameter(Mandatory)][string]$Path)
    $encoded = [IO.File]::ReadAllText([IO.Path]::GetFullPath($Path))
    return ConvertTo-SecureString -String $encoded
}

Export-ModuleMember -Function Save-GeminiKey, Get-GeminiKey
