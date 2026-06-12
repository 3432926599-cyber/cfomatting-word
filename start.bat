@echo off
chcp 65001 >nul
title CFomatting — 文档格式自动化
echo ============================================
echo   CFomatting — 文档格式自动化
echo ============================================
echo.
echo [1/3] 启动后端服务...
start "CFomatting-Backend" /min cmd /c "cd /d \"%~dp0backend\" && C:\Users\Administrator\miniconda3\Scripts\uvicorn.exe app.main:app --host 0.0.0.0 --port 8000"

echo [2/3] 启动公网隧道...
echo.
cd /d "%~dp0"

:: 用 PowerShell 启动隧道、抓取链接、更新 GitHub、复制到剪贴板
powershell -NoProfile -ExecutionPolicy Bypass -Command ^"
  Write-Host '  等待隧道连接...' -ForegroundColor Gray; ^
  ^$output = ^& { .\cloudflared.exe tunnel --url http://localhost:8000 2^>^&1 }; ^
  ^$url = ((^$output | Select-String 'trycloudflare\.com').Line -split '\s+')[2]; ^
  if (^$url) { ^
    ^$fullUrl = 'https://cfomatting-word.pages.dev/?backend=' + ^$url; ^
    Set-Clipboard -Value ^$fullUrl; ^
    Write-Host ''; ^
    Write-Host '  ============================================' -ForegroundColor Green; ^
    Write-Host '   隧道已连接！链接已复制到剪贴板' -ForegroundColor Green; ^
    Write-Host '  ============================================' -ForegroundColor Green; ^
    Write-Host ''; ^
    Write-Host '   前端页面: ' -NoNewline; ^
    Write-Host ^$fullUrl -ForegroundColor Cyan; ^
    Write-Host '   后端直连: ' -NoNewline; ^
    Write-Host ^$url -ForegroundColor Gray; ^
    Write-Host ''; ^
    ^
    Write-Host '[3/3] 更新 GitHub 指针...' -ForegroundColor Gray; ^
    try { ^
      ^$settings = Get-Content -Path ^"^$env:USERPROFILE\.claude\settings.json^" -Raw | ConvertFrom-Json; ^
      ^$token = ^$settings.env.GH_TOKEN; ^
      if (-not ^$token) { throw 'Token not found' }; ^
      ^
      ^$headers = @{ ^
        Authorization = ^"Bearer ^$token^"; ^
        Accept = 'application/vnd.github+json' ^
      }; ^
      ^
      ^$apiBase = 'https://api.github.com/repos/3432926599-cyber/cfomatting-word/contents/backend-url.txt'; ^
      ^
      ^$current = Invoke-RestMethod -Uri ^$apiBase -Headers ^$headers -Method Get -Proxy 'http://127.0.0.1:7897' -ErrorAction Stop; ^
      ^$sha = ^$current.sha; ^
      ^
      ^$content = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(^$url)); ^
      ^$body = @{ message = 'Update tunnel URL'; content = ^$content; sha = ^$sha; branch = 'main' } | ConvertTo-Json; ^
      ^
      Invoke-RestMethod -Uri ^$apiBase -Headers ^$headers -Method Put -Body ^$body -ContentType 'application/json' -Proxy 'http://127.0.0.1:7897' -ErrorAction Stop ^| Out-Null; ^
      ^
      Write-Host '   已更新 GitHub 指针 ✓' -ForegroundColor Green; ^
    } catch { ^
      Write-Host ('   更新 GitHub 失败: ' + ^$_.Exception.Message) -ForegroundColor Yellow; ^
      Write-Host '   （不影响使用，链接已复制到剪贴板）' -ForegroundColor Gray; ^
    } ^
    ^
    Write-Host ''; ^
    Write-Host '  直接粘贴到微信发送即可（链接已复制）'; ^
    Write-Host '  个人网站链接: https://cfomatting-word.pages.dev（永久有效）' -ForegroundColor Green; ^
    Write-Host '  按 Ctrl+C 关闭隧道' -ForegroundColor Gray; ^
    Wait-Event; ^
  } else { ^
    Write-Host '获取链接失败，请查看上方输出' -ForegroundColor Red; ^
    pause; ^
  }
"
