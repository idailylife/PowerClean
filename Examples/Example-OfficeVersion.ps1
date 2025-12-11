# PowerClean Office Version Detection Example
# Office版本检测示例

# 导入模块
Import-Module -Path "..\PowerClean.psm1" -Force

Write-Host "================================" -ForegroundColor Cyan
Write-Host "PowerClean - Office版本检测示例" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# 示例1: 检测Office是否已安装
Write-Host "示例1: 检查Office是否安装" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
$isInstalled = Test-OfficeInstalled
Write-Host ""

# 示例2: 获取所有Office版本信息
Write-Host "示例2: 获取详细的Office版本信息" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
$versions = Get-OfficeVersion
Write-Host ""

# 示例3: 获取Office安装路径
Write-Host "示例3: 获取Office安装路径" -ForegroundColor Magenta
Write-Host "---" -ForegroundColor Gray
$installPath = Get-OfficeInstallPath
Write-Host ""

# 示例4: 将结果保存到文件（可选）
if ($versions) {
    Write-Host "示例4: 将Office版本信息导出到CSV文件" -ForegroundColor Magenta
    Write-Host "---" -ForegroundColor Gray
    $versions | Export-Csv -Path "office_versions.csv" -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "已将Office版本信息导出到 office_versions.csv" -ForegroundColor Green
    Write-Host ""
    
    # 示例5: 检查是否为Microsoft 365
    Write-Host "示例5: 检查是否为Microsoft 365" -ForegroundColor Magenta
    Write-Host "---" -ForegroundColor Gray
    $is365 = $versions | Where-Object { $_.ProductName -like "*Microsoft 365*" }
    if ($is365) {
        Write-Host "检测到Microsoft 365版本: $($is365.Version)" -ForegroundColor Green
    }
    else {
        Write-Host "未检测到Microsoft 365" -ForegroundColor Yellow
    }
}
else {
    Write-Host "未检测到安装的Office" -ForegroundColor Yellow
}
Write-Host ""

