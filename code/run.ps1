<#
.SYNOPSIS
  Wrapper to run svg-to-pptx python scripts on this machine.

.DESCRIPTION
  On this machine, the python.exe launcher inside a uv-created venv is blocked by an
  application control policy, so `uv run` does not work directly here.
  See CLAUDE.md ("execution environment") and docs/DECISIONS.md (ADR-017) for details.
  As a workaround, this script invokes the trusted uv-managed base interpreter directly,
  and points PYTHONPATH at the venv's site-packages so the same dependencies are available.

.EXAMPLE
  .\code\run.ps1 code\build_pptx.py --session foo --no-text input\test\x.svg --text-only input\test\x.svg
  .\code\run.ps1 code\qa_capture.py output\foo\foo_native.pptx
#>
param(
    [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)]
    [string[]]$ScriptArgs
)

$ErrorActionPreference = "Stop"

$venv = if ($env:UV_PROJECT_ENVIRONMENT) { $env:UV_PROJECT_ENVIRONMENT } else { "$env:USERPROFILE\.venvs\svg-to-pptx" }
$sitePackages = Join-Path $venv "Lib\site-packages"

if (-not (Test-Path $sitePackages)) {
    Write-Error "venv not found: $sitePackages`nRun first:`n  `$env:UV_PROJECT_ENVIRONMENT = '$venv'`n  uv sync"
}

$base = (uv python find 3.12).Trim()
if (-not $base -or -not (Test-Path $base)) {
    Write-Error "Could not resolve uv base interpreter (uv python find 3.12)."
}

$env:PYTHONPATH = @(
    $sitePackages,
    (Join-Path $sitePackages "win32"),
    (Join-Path $sitePackages "win32\lib"),
    (Join-Path $sitePackages "Pythonwin"),
    (Join-Path $sitePackages "pywin32_system32")
) -join ";"

& $base @ScriptArgs
