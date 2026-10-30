# Recreate spps_db database
# This script drops and recreates the database

Write-Host "Recreating spps_db Database" -ForegroundColor Cyan
Write-Host "===========================" -ForegroundColor Cyan
Write-Host ""

Write-Host "IMPORTANT: Make sure MySQL is running in XAMPP Control Panel!" -ForegroundColor Yellow
Write-Host "Press any key after MySQL is started, or Ctrl+C to cancel..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

$xamppPath = "C:\xampp\mysql"

Write-Host ""
Write-Host "[1/3] Stopping MySQL processes (if any)..." -ForegroundColor Green
Get-Process | Where-Object {$_.ProcessName -like "*mysql*"} | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Write-Host "Done." -ForegroundColor Green
Write-Host ""

Write-Host "[2/3] Removing database directory..." -ForegroundColor Green
$dbPath = "$xamppPath\data\spps_db"
if (Test-Path $dbPath) {
    Remove-Item -Path $dbPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Removed: $dbPath" -ForegroundColor Green
} else {
    Write-Host "Database directory doesn't exist." -ForegroundColor Yellow
}
Write-Host ""

Write-Host "[3/3] Creating fresh database..." -ForegroundColor Green
Write-Host "Please start MySQL in XAMPP Control Panel now, then press any key..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

Start-Sleep -Seconds 2

$mysqlPath = "$xamppPath\bin\mysql.exe"
$result = & $mysqlPath -u root -e "CREATE DATABASE IF NOT EXISTS spps_db;" 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "Database created successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "SUCCESS! Database recreated." -ForegroundColor Green
    Write-Host ""
    Write-Host "Next step: Restart your Flask application (python run.py)" -ForegroundColor Cyan
    Write-Host "The app will automatically create all tables on startup." -ForegroundColor Cyan
} else {
    Write-Host "ERROR creating database:" -ForegroundColor Red
    Write-Host $result -ForegroundColor Red
    Write-Host ""
    Write-Host "Make sure MySQL is running in XAMPP Control Panel!" -ForegroundColor Yellow
}

