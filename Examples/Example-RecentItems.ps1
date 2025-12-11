# PowerClean 使用示例
# 演示如何使用 PowerClean 模块管理最近项目显示

Import-Module ..\PowerClean.psd1 -Force

Write-Host "
========================================" -ForegroundColor Cyan
Write-Host "PowerClean 模块示例" -ForegroundColor Cyan
Write-Host "========================================
" -ForegroundColor Cyan

Write-Host "1. 查看当前最近项目显示状态" -ForegroundColor Yellow
Get-RecentItemsStatus
Read-Host "
按回车键继续..."

Write-Host "
2. 禁用最近项目显示" -ForegroundColor Yellow
Disable-RecentItems
Start-Sleep -Seconds 2

Write-Host "
3. 查看更改后的状态" -ForegroundColor Yellow
Get-RecentItemsStatus
Read-Host "
按回车键继续..."

Write-Host "
4. 重新启用最近项目显示" -ForegroundColor Yellow
Enable-RecentItems
Start-Sleep -Seconds 2

Write-Host "
5. 最终状态" -ForegroundColor Yellow
Get-RecentItemsStatus

Write-Host "
========================================" -ForegroundColor Cyan
Write-Host "示例完成!" -ForegroundColor Green
Write-Host "========================================
" -ForegroundColor Cyan

$restart = Read-Host "是否要立即重启资源管理器以应用更改? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Restart-Explorer
}
else {
    Write-Host "
更改将在重启资源管理器或重新登录后生效" -ForegroundColor Cyan
}
