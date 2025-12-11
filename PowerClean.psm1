# PowerClean PowerShell Module
# 用于管理Windows系统界面和功能的工具库

<#
.SYNOPSIS
    PowerClean - Windows系统清理和配置工具库
.DESCRIPTION
    提供用于管理Windows开始菜单、任务栏和其他系统功能的PowerShell函数
#>

# 获取模块根目录
$ModuleRoot = $PSScriptRoot

# 导入所有函数脚本
$FunctionsPath = Join-Path -Path $ModuleRoot -ChildPath "Functions"
if (Test-Path $FunctionsPath) {
    Get-ChildItem -Path $FunctionsPath -Filter "*.ps1" | ForEach-Object {
        . $_.FullName
    }
}

# 导出模块函数
Export-ModuleMember -Function Set-RecentItemsDisplay, Enable-RecentItems, Disable-RecentItems, Get-RecentItemsStatus, Restart-Explorer, Set-QuickAccessRecentFiles, Set-QuickAccessFrequentFolders, Enable-QuickAccessRecentFiles, Disable-QuickAccessRecentFiles, Enable-QuickAccessFrequentFolders, Disable-QuickAccessFrequentFolders, Get-QuickAccessStatus, Get-OfficeVersion, Test-OfficeInstalled, Get-OfficeInstallPath, Disable-OfficeRecentDocuments, Enable-OfficeRecentDocuments, Get-OfficeRecentDocumentsStatus, Set-OneDriveSync, Enable-OneDrivePersonalSync, Disable-OneDrivePersonalSync, Enable-OneDriveBusinessSync, Disable-OneDriveBusinessSync, Show-OneDriveInExplorer, Hide-OneDriveInExplorer, Get-OneDriveStatus, Test-WPSInstalled, Get-WPSUninstallPath, Uninstall-WPSOffice, Get-WPSRemnantFiles, Remove-WPSRemnantFiles
