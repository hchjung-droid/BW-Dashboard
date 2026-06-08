@echo off
chcp 65001 >nul
echo ============================================
echo   SCM Dashboard - Local Server
echo ============================================
echo.

:: Try Python first
where python >nul 2>&1
if %ERRORLEVEL%==0 (
    echo [Python 서버 시작]
    echo.
    echo   http://localhost:8080/test_connection.html  (연결 테스트)
    echo   http://localhost:8080/index.html            (대시보드)
    echo.
    echo   종료: Ctrl+C 또는 이 창 닫기
    echo ============================================
    python -m http.server 8080
    goto :end
)

:: Try Node.js npx
where npx >nul 2>&1
if %ERRORLEVEL%==0 (
    echo [Node.js 서버 시작]
    echo.
    echo   http://localhost:8080/test_connection.html  (연결 테스트)
    echo   http://localhost:8080/index.html            (대시보드)
    echo.
    echo   종료: Ctrl+C 또는 이 창 닫기
    echo ============================================
    npx -y serve -l 8080
    goto :end
)

:: Fallback: PowerShell HTTP listener
echo [PowerShell 서버 시작]
echo.
echo   http://localhost:8080/test_connection.html  (연결 테스트)
echo   http://localhost:8080/index.html            (대시보드)
echo.
echo   종료: Ctrl+C 또는 이 창 닫기
echo ============================================
powershell -ExecutionPolicy Bypass -Command ^
  "$listener = New-Object System.Net.HttpListener; $listener.Prefixes.Add('http://localhost:8080/'); $listener.Start(); Write-Host 'Server running on http://localhost:8080/'; $root = (Get-Location).Path; while ($listener.IsListening) { $ctx = $listener.GetContext(); $path = $ctx.Request.Url.LocalPath; if ($path -eq '/') { $path = '/index.html' }; $file = Join-Path $root $path.TrimStart('/'); if (Test-Path $file) { $bytes = [IO.File]::ReadAllBytes($file); $ext = [IO.Path]::GetExtension($file); $mime = @{'.html'='text/html;charset=utf-8';'.js'='application/javascript;charset=utf-8';'.css'='text/css;charset=utf-8';'.json'='application/json';'.png'='image/png';'.jpg'='image/jpeg';'.svg'='image/svg+xml'}; $ct = $mime[$ext]; if (-not $ct) { $ct = 'application/octet-stream' }; $ctx.Response.ContentType = $ct; $ctx.Response.ContentLength64 = $bytes.Length; $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length) } else { $ctx.Response.StatusCode = 404; $msg = [Text.Encoding]::UTF8.GetBytes('Not Found'); $ctx.Response.OutputStream.Write($msg, 0, $msg.Length) }; $ctx.Response.Close() }"

:end
pause
