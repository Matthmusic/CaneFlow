# Clear Windows icon cache and CaneFlow shortcuts (use as needed)

Write-Host "=== CaneFlow icon cleanup ===" -ForegroundColor Cyan

# Stop CaneFlow if running
Write-Host "`n1. Closing CaneFlow..." -ForegroundColor Yellow
$processes = Get-Process | Where-Object { $_.ProcessName -like "*CaneFlow*" -or $_.ProcessName -like "*caneflow*" }
if ($processes) {
    $processes | ForEach-Object {
        Write-Host "  - Stopping $($_.ProcessName) (PID: $($_.Id))"
        Stop-Process -Id $_.Id -Force
    }
} else {
    Write-Host "  - No CaneFlow process found"
}

# Remove shortcuts
Write-Host "`n2. Removing shortcuts..." -ForegroundColor Yellow
$shortcuts = @(
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\CaneFlow.lnk",
    "$env:PUBLIC\Desktop\CaneFlow.lnk",
    "$env:USERPROFILE\Desktop\CaneFlow.lnk"
)

foreach ($shortcut in $shortcuts) {
    if (Test-Path $shortcut) {
        Remove-Item $shortcut -Force
        Write-Host "  - Removed: $shortcut" -ForegroundColor Green
    }
}

# Clear icon cache
Write-Host "`n3. Clearing icon cache..." -ForegroundColor Yellow
Write-Host "  - Stopping explorer..."
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$iconCachePaths = @(
    "$env:LOCALAPPDATA\IconCache.db",
    "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache_*.db",
    "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache_*.db"
)

foreach ($pathPattern in $iconCachePaths) {
    Get-ChildItem -Path $pathPattern -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Remove-Item $_.FullName -Force -ErrorAction Stop
            Write-Host "  - Deleted: $($_.Name)" -ForegroundColor Green
        } catch {
            Write-Host "  - Failed: $($_.Name) (locked)" -ForegroundColor Red
        }
    }
}

Write-Host "  - Restarting explorer..."
Start-Process explorer.exe

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Reinstall the app and recreate the desktop shortcut."
