# PowerClean Office Recent Documents Management Example
# Office最近文档列表管理示例

# 导入模块
Import-Module -Path "..\PowerClean.psm1" -Force

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "PowerClean - Office最近文档列表管理示例" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# 示例1: 检查Office最近文档列表状态
Write-Host "示例1: 检查Office最近文档列表状态" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
$status = Get-OfficeRecentDocumentsStatus
Write-Host ""

# 示例2: 禁用Office最近文档列表
Write-Host "示例2: 禁用Office最近文档列表" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
Write-Host "说明: 这将禁用Word、Excel、PowerPoint等所有Office应用的最近文档显示" -ForegroundColor White
Write-Host "是否继续? (Y/N)" -ForegroundColor Yellow
$response = Read-Host
if ($response -eq "Y" -or $response -eq "y") {
    $result = Disable-OfficeRecentDocuments
    if ($result) {
        Write-Host "禁用成功！" -ForegroundColor Green
    }
}
Write-Host ""

# 示例3: 启用Office最近文档列表
Write-Host "示例3: 启用Office最近文档列表" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
Write-Host "说明: 这将启用所有Office应用的最近文档显示" -ForegroundColor White
Write-Host "是否继续? (Y/N)" -ForegroundColor Yellow
$response = Read-Host
if ($response -eq "Y" -or $response -eq "y") {
    $result = Enable-OfficeRecentDocuments
    if ($result) {
        Write-Host "启用成功！" -ForegroundColor Green
    }
}
Write-Host ""

# 示例4: 再次检查状态以确认更改
Write-Host "示例4: 再次检查状态以确认更改" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
$newStatus = Get-OfficeRecentDocumentsStatus
Write-Host ""

# 示例5: 快速切换状态
Write-Host "示例5: 快速切换最近文档列表状态" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
Write-Host "当前状态: $($status.GlobalSetting -eq 1 ? '启用' : '禁用')" -ForegroundColor Cyan
if ($status.GlobalSetting -eq 1) {
    Write-Host "快速禁用: Disable-OfficeRecentDocuments" -ForegroundColor White
}
else {
    Write-Host "快速启用: Enable-OfficeRecentDocuments" -ForegroundColor White
}
Write-Host ""

Write-Host "提示: 所有更改需要重启Office应用后才能生效" -ForegroundColor Yellow
Write-Host ""
