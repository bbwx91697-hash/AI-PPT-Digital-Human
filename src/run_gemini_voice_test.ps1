param(
    [Parameter(Mandatory)][string]$KeyFile,
    [Parameter(Mandatory)][string]$OutputDirectory,
    [switch]$ConfigureKey
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'credential_store.psm1') -Force
if ($ConfigureKey) {
    $key = Read-Host '请输入 Gemini API Key（输入不会显示）' -AsSecureString
    Save-GeminiKey -Key $key -Path $KeyFile
} else {
    $key = Get-GeminiKey -Path $KeyFile
}
$pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($key)
try {
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    $plain | & python (Join-Path $PSScriptRoot 'test_gemini_voices.py') $OutputDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Gemini 测试未正常结束。' }
} finally {
    $plain = $null
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
}
