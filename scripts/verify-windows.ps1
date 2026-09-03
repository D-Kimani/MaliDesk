$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Write-Host "Checking MaliDesk files..." -ForegroundColor Cyan
$required = @(
  "$root\frontend\package.json",
  "$root\frontend\vite.config.js",
  "$root\frontend\index.html",
  "$root\frontend\src\main.jsx",
  "$root\frontend\src\MaliDesk.jsx",
  "$root\frontend\src\index.css",
  "$root\server\package.json",
  "$root\server\index.js",
  "$root\server\schema.sql",
  "$root\server\bootstrap-admin.js"
)
foreach ($file in $required) {
  if (-not (Test-Path $file)) { throw "Missing: $file" }
  Write-Host "OK  $file" -ForegroundColor Green
}
Write-Host "All required files are present." -ForegroundColor Green
