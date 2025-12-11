# 快速禁用最近项目

Import-Module ..\PowerClean.psd1 -Force

Write-Host "禁用最近项目显示..." -ForegroundColor Yellow
Disable-RecentItems
Get-RecentItemsStatus

$restart = Read-Host "`n是否立即重启资源管理器? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Restart-Explorer
    Write-Host "`n完成! 最近项目显示已禁用。" -ForegroundColor Green
}
else {
    Write-Host "`n完成! 请重启资源管理器或重新登录以应用更改。" -ForegroundColor Cyan
}
