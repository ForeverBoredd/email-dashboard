$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot
Write-Host "Starting School Mail Priorities..."
Write-Host "Open http://localhost:8787 in your browser."
Write-Host ""
node server.js
