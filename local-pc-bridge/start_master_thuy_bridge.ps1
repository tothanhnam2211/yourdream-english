# =============================================================================
# MASTER THÚY — ENGLISH STUDIO v2.0 (1-CLICK STARTUP & 00_PRINT_QUEUE BRIDGE)
# =============================================================================
$env:PYTHONIOENCODING = "utf-8"
$ProjectRoot = "D:\YOURDREAM.VN\master-thuy-english-studio"
$PythonExe = "D:\YOURDREAM.VN\venv\Scripts\python.exe"

Write-Host "Dang khoi dong MASTER THUY - ENGLISH STUDIO v2.0 & 00_PRINT_QUEUE Bridge tren cong 8765..." -ForegroundColor Green
Start-Process "http://localhost:8765/teacher/index.html"
& $PythonExe "$ProjectRoot\local-pc-bridge\server_and_print_bridge.py"
