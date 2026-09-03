$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

Write-Host "=== MaliDesk Windows setup ===" -ForegroundColor Green

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js is not installed. Install Node.js LTS, then run this script again."
}

Write-Host "Node: $(node --version)"

Write-Host "Installing frontend dependencies..." -ForegroundColor Cyan
Push-Location "$root\frontend"
npm install
Pop-Location

Write-Host "Installing server dependencies..." -ForegroundColor Cyan
Push-Location "$root\server"
npm install
Pop-Location

if (-not (Test-Path "$root\server\.env")) {
  Copy-Item "$root\server\.env.example" "$root\server\.env"
  Write-Host "Created server\.env. Review DATABASE_URL and SESSION_SECRET before starting the API." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Setup complete." -ForegroundColor Green
Write-Host "Next:" -ForegroundColor Cyan
Write-Host "  1. Start PostgreSQL (Docker Desktop: docker compose up -d postgres)"
Write-Host "  2. Apply server\schema.sql to the database"
Write-Host "  3. Run server\bootstrap-admin.js once to create the first Administrator"
Write-Host "  4. Start API: cd server; npm start"
Write-Host "  5. Start UI:  cd frontend; npm run dev"
