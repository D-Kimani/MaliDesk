$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

Start-Process powershell -ArgumentList '-NoExit','-Command',"Set-Location '$root\server'; npm start"
Start-Sleep -Seconds 2
Start-Process powershell -ArgumentList '-NoExit','-Command',"Set-Location '$root\frontend'; npm run dev"

Write-Host "MaliDesk API: http://127.0.0.1:3001" -ForegroundColor Green
Write-Host "MaliDesk UI:  http://127.0.0.1:5173" -ForegroundColor Green
