# PowerClean 主程序
# 二级交互式命令行菜单

# 检查是否以管理员身份运行
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# 导入模块
$modulePath = Join-Path $PSScriptRoot "PowerClean.psd1"
Import-Module $modulePath -Force

# ==================== 一级菜单 ====================
function Show-MainMenu {
    Clear-Host
    Write-Host "" -ForegroundColor Cyan
    Write-Host "     PowerClean 系统管理工具           " -ForegroundColor Cyan
    Write-Host "" -ForegroundColor Cyan
    
    # 显示权限状态
    if ($isAdmin) {
        Write-Host "     [管理员模式]                    " -ForegroundColor Green
    } else {
        Write-Host "     [普通用户模式]                  " -ForegroundColor Yellow
    }
    Write-Host "" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "【功能模块】" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  1. 开始菜单和任务栏 - 最近项目管理" -ForegroundColor White
    Write-Host "  2. 文件资源管理器 - 快速访问管理" -ForegroundColor White
    Write-Host "  3. Microsoft Office - 最近文档管理" -ForegroundColor White
    Write-Host "  4. OneDrive - 云同步管理" -ForegroundColor White
    Write-Host "  5. WPS Office - 云盘管理" -ForegroundColor White
    Write-Host ""
    Write-Host "【系统工具】" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  9. 重启资源管理器" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  0. 退出程序" -ForegroundColor Red
    Write-Host ""
    Write-Host "" -ForegroundColor Cyan
}

# ==================== 二级菜单 - 最近项目 ====================
function Show-RecentItemsMenu {
    param([string]$Action = "")
    
    Clear-Host
    Write-Host "" -ForegroundColor Cyan
    Write-Host "   开始菜单和任务栏 - 最近项目管理    " -ForegroundColor Cyan
    Write-Host "" -ForegroundColor Cyan
    
    # 显示当前状态
    Get-RecentItemsStatus
    
    if ($Action -ne "") {
        Write-Host "" -ForegroundColor Cyan
        Write-Host $Action -ForegroundColor Green
        Write-Host ""
    }
    
    Write-Host "" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1. 启用最近项目显示" -ForegroundColor Green
    Write-Host "  2. 禁用最近项目显示" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  0. 返回主菜单" -ForegroundColor Gray
    Write-Host ""
    Write-Host "" -ForegroundColor Cyan
}

function Handle-RecentItemsMenu {
    do {
        Show-RecentItemsMenu
        $choice = Read-Host "请选择操作"
        
        switch ($choice) {
            '1' {
                Write-Host "`n确认要启用最近项目显示吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Enable-RecentItems
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-RecentItemsMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '2' {
                Write-Host "`n确认要禁用最近项目显示吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Disable-RecentItems
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-RecentItemsMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '0' { return }
            default { 
                Write-Host "`n无效的选择！" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($choice -ne '0')
}

# ==================== 二级菜单 - 快速访问 ====================
function Show-QuickAccessMenu {
    param([string]$Action = "")
    
    Clear-Host
    Write-Host "" -ForegroundColor Cyan
    Write-Host "  文件资源管理器 - 快速访问管理       " -ForegroundColor Cyan
    Write-Host "" -ForegroundColor Cyan
    
    # 显示当前状态
    Get-QuickAccessStatus
    
    if ($Action -ne "") {
        Write-Host "" -ForegroundColor Cyan
        Write-Host $Action -ForegroundColor Green
        Write-Host ""
    }
    
    Write-Host "" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "【最近使用的文件】" -ForegroundColor Yellow
    Write-Host "  1. 启用最近文件显示" -ForegroundColor Green
    Write-Host "  2. 禁用最近文件显示" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "【常用文件夹】" -ForegroundColor Yellow
    Write-Host "  3. 启用常用文件夹显示" -ForegroundColor Green
    Write-Host "  4. 禁用常用文件夹显示" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  0. 返回主菜单" -ForegroundColor Gray
    Write-Host ""
    Write-Host "" -ForegroundColor Cyan
}

function Handle-QuickAccessMenu {
    do {
        Show-QuickAccessMenu
        $choice = Read-Host "请选择操作"
        
        switch ($choice) {
            '1' {
                Write-Host "`n确认要启用快速访问最近文件显示吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Enable-QuickAccessRecentFiles
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-QuickAccessMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '2' {
                Write-Host "`n确认要禁用快速访问最近文件显示吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Disable-QuickAccessRecentFiles
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-QuickAccessMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '3' {
                Write-Host "`n确认要启用快速访问常用文件夹显示吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Enable-QuickAccessFrequentFolders
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-QuickAccessMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '4' {
                Write-Host "`n确认要禁用快速访问常用文件夹显示吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Disable-QuickAccessFrequentFolders
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-QuickAccessMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '0' { return }
            default { 
                Write-Host "`n无效的选择！" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($choice -ne '0')
}

# ==================== 二级菜单 - Office最近文档 ====================
function Show-OfficeRecentDocumentsMenu {
    param([string]$Action = "")
    
    Clear-Host
    Write-Host "" -ForegroundColor Cyan
    Write-Host "   Microsoft Office - 最近文档管理    " -ForegroundColor Cyan
    Write-Host "" -ForegroundColor Cyan
    
    # 显示当前状态
    Get-OfficeRecentDocumentsStatus
    
    if ($Action -ne "") {
        Write-Host "" -ForegroundColor Cyan
        Write-Host $Action -ForegroundColor Green
        Write-Host ""
    }
    
    Write-Host "" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1. 启用最近文档列表" -ForegroundColor Green
    Write-Host "  2. 禁用最近文档列表" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  0. 返回主菜单" -ForegroundColor Gray
    Write-Host ""
    Write-Host "" -ForegroundColor Cyan
}

function Handle-OfficeRecentDocumentsMenu {
    do {
        Show-OfficeRecentDocumentsMenu
        $choice = Read-Host "请选择操作"
        
        switch ($choice) {
            '1' {
                Write-Host "`n确认要启用Office最近文档列表吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Enable-OfficeRecentDocuments
                    
                    Write-Host "`n操作完成！按任意键继续..." -ForegroundColor Green
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    
                    Show-OfficeRecentDocumentsMenu " 操作完成！"
                    Start-Sleep -Seconds 1
                }
            }
            '2' {
                Write-Host "`n确认要禁用Office最近文档列表吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Disable-OfficeRecentDocuments
                    
                    Write-Host "`n操作完成！按任意键继续..." -ForegroundColor Green
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    
                    Show-OfficeRecentDocumentsMenu " 操作完成！"
                    Start-Sleep -Seconds 1
                }
            }
            '0' { return }
            default { 
                Write-Host "`n无效的选择！" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($choice -ne '0')
}

# ==================== 二级菜单 - OneDrive同步 ====================
function Show-OneDriveMenu {
    param([string]$Action = "")
    
    Clear-Host
    Write-Host "" -ForegroundColor Cyan
    Write-Host "       OneDrive - 云同步管理          " -ForegroundColor Cyan
    Write-Host "" -ForegroundColor Cyan
    
    # 显示当前状态
    Get-OneDriveStatus
    
    if ($Action -ne "") {
        Write-Host "" -ForegroundColor Cyan
        Write-Host $Action -ForegroundColor Green
        Write-Host ""
    }
    
    Write-Host "" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "【OneDrive个人版】" -ForegroundColor Yellow
    Write-Host "  1. 启用OneDrive个人版同步" -ForegroundColor Green
    Write-Host "  2. 禁用OneDrive个人版同步" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "【OneDrive商业版】" -ForegroundColor Yellow
    Write-Host "  3. 启用OneDrive商业版同步" -ForegroundColor Green
    Write-Host "  4. 禁用OneDrive商业版同步" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "【文件资源管理器】" -ForegroundColor Yellow
    Write-Host "  5. 在文件资源管理器中显示OneDrive" -ForegroundColor Green
    Write-Host "  6. 在文件资源管理器中隐藏OneDrive" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  0. 返回主菜单" -ForegroundColor Gray
    Write-Host ""
    Write-Host "" -ForegroundColor Cyan
}

function Handle-OneDriveMenu {
    do {
        Show-OneDriveMenu
        $choice = Read-Host "请选择操作"
        
        switch ($choice) {
            '1' {
                Write-Host "`n确认要启用OneDrive个人版同步吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Enable-OneDrivePersonalSync
                    
                    Write-Host "`n操作完成！按任意键继续..." -ForegroundColor Green
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    
                    Show-OneDriveMenu " 操作完成！"
                    Start-Sleep -Seconds 1
                }
            }
            '2' {
                Write-Host "`n确认要禁用OneDrive个人版同步吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Disable-OneDrivePersonalSync
                    
                    Write-Host "`n操作完成！按任意键继续..." -ForegroundColor Green
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    
                    Show-OneDriveMenu " 操作完成！"
                    Start-Sleep -Seconds 1
                }
            }
            '3' {
                Write-Host "`n确认要启用OneDrive商业版同步吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Enable-OneDriveBusinessSync
                    
                    Write-Host "`n操作完成！按任意键继续..." -ForegroundColor Green
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    
                    Show-OneDriveMenu " 操作完成！"
                    Start-Sleep -Seconds 1
                }
            }
            '4' {
                Write-Host "`n确认要禁用OneDrive商业版同步吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Disable-OneDriveBusinessSync
                    
                    Write-Host "`n操作完成！按任意键继续..." -ForegroundColor Green
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    
                    Show-OneDriveMenu " 操作完成！"
                    Start-Sleep -Seconds 1
                }
            }
            '5' {
                Write-Host "`n确认要在文件资源管理器中显示OneDrive吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Show-OneDriveInExplorer
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-OneDriveMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '6' {
                Write-Host "`n确认要在文件资源管理器中隐藏OneDrive吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Hide-OneDriveInExplorer
                    
                    Write-Host "`n是否立即重启资源管理器以应用更改? (Y/N): " -ForegroundColor Yellow -NoNewline
                    $restart = Read-Host
                    if ($restart -eq 'Y' -or $restart -eq 'y') {
                        Restart-Explorer
                    }
                    
                    Show-OneDriveMenu " 操作完成！"
                    Start-Sleep -Seconds 2
                }
            }
            '0' { return }
            default { 
                Write-Host "`n无效的选择！" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($choice -ne '0')
}

# ==================== 二级菜单 - WPS Office云盘 ====================
function Show-WPSOfficeMenu {
    param([string]$Action = "")
    
    Clear-Host
    Write-Host "" -ForegroundColor Cyan
    Write-Host "      WPS Office - 安全管理           " -ForegroundColor Cyan
    Write-Host "" -ForegroundColor Cyan
    
    # 检测WPS是否安装
    $wpsInstalled = Test-WPSInstalled
    
    Write-Host ""
    Write-Host "【当前状态】" -ForegroundColor Cyan
    Write-Host ""
    
    if ($wpsInstalled) {
        Write-Host "  WPS Office: " -NoNewline
        Write-Host "已安装" -ForegroundColor Red
        
        # 获取安装路径
        $wpsBasePath = "HKCU:\Software\kingsoft\Office"
        $versionPaths = Get-ChildItem $wpsBasePath -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '\d+\.\d+' }
        if ($versionPaths) {
            $wpsCommonPath = Join-Path $versionPaths[0].PSPath "Common"
            if (Test-Path $wpsCommonPath) {
                $commonProps = Get-ItemProperty -Path $wpsCommonPath -ErrorAction SilentlyContinue
                if ($commonProps.InstallRoot) {
                    Write-Host "  安装路径: " -NoNewline
                    Write-Host $commonProps.InstallRoot -ForegroundColor Cyan
                }
            }
        }
        
        Write-Host ""
        Write-Host "  ⚠  安全警告：检测到 WPS Office 已安装" -ForegroundColor Red
        Write-Host "     WPS Office 可能存在数据泄露风险" -ForegroundColor Yellow
        Write-Host "     建议企业环境下卸载此软件" -ForegroundColor Yellow
    } else {
        Write-Host "  WPS Office: " -NoNewline
        Write-Host "未安装" -ForegroundColor Green
        
        # 检测遗留文件
        $remnantPaths = @(
            "$env:USERPROFILE\WPS Cloud Files",
            "$env:APPDATA\kingsoft",
            "$env:LOCALAPPDATA\Kingsoft"
        )
        
        $hasRemnants = $false
        foreach ($path in $remnantPaths) {
            if (Test-Path $path) {
                $hasRemnants = $true
                break
            }
        }
        
        if ($hasRemnants) {
            Write-Host "  ⚠  发现 WPS 遗留文件" -ForegroundColor Yellow
        } else {
            Write-Host "  系统安全状态良好" -ForegroundColor Green
        }
    }
    
    if ($Action -ne "") {
        Write-Host "" -ForegroundColor Cyan
        Write-Host $Action -ForegroundColor Green
        Write-Host ""
    }
    
    Write-Host "" -ForegroundColor Cyan
    Write-Host ""
    
    if ($wpsInstalled) {
        Write-Host "【系统管理】" -ForegroundColor Red
        Write-Host "  1. 卸载 WPS Office（推荐）" -ForegroundColor Red
    } else {
        Write-Host "【遗留文件检测】" -ForegroundColor Yellow
        Write-Host "  1. 扫描 WPS 遗留文件" -ForegroundColor Cyan
        Write-Host "  2. 清理 WPS 遗留文件" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "  0. 返回主菜单" -ForegroundColor Gray
    Write-Host ""
    Write-Host "" -ForegroundColor Cyan
}

function Handle-WPSOfficeMenu {
    do {
        Show-WPSOfficeMenu
        $choice = Read-Host "请选择操作"
        
        switch ($choice) {
            '1' {
                if (Test-WPSInstalled) {
                    Write-Host "`n" -ForegroundColor Red
                    Write-Host "⚠  警告：您即将卸载 WPS Office" -ForegroundColor Red
                    Write-Host "请确认您已备份所有重要文档！" -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "确认要卸载WPS Office吗? (Y/N): " -ForegroundColor Red -NoNewline
                    $confirm = Read-Host
                    
                    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                        $uninstallResult = Uninstall-WPSOffice
                        
                        Write-Host "`n按任意键继续..." -ForegroundColor Green
                        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                        
                        # 卸载后显示更新的状态
                        if ($uninstallResult) {
                            Show-WPSOfficeMenu " WPS Office 已成功卸载！"
                            Start-Sleep -Seconds 2
                        }
                    }
                } else {
                    # 扫描遗留文件
                    Get-WPSRemnantFiles
                    
                    Write-Host "`n按任意键返回..." -ForegroundColor Green
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
            }
            '2' {
                if (-not (Test-WPSInstalled)) {
                    Write-Host "`n" -ForegroundColor Yellow
                    Write-Host "⚠  警告：即将清理 WPS Office 遗留文件" -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "清理选项：" -ForegroundColor Cyan
                    Write-Host "  1. 全部清理（包括云文档）" -ForegroundColor Red
                    Write-Host "  2. 仅清理配置和缓存（保留云文档）" -ForegroundColor Yellow
                    Write-Host "  0. 取消" -ForegroundColor Gray
                    Write-Host ""
                    Write-Host "请选择 (1/2/0): " -NoNewline
                    $cleanChoice = Read-Host
                    
                    if ($cleanChoice -eq '1') {
                        Write-Host "`n确认要删除所有 WPS 遗留文件（包括云文档）吗? (Y/N): " -ForegroundColor Red -NoNewline
                        $confirm = Read-Host
                        if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                            Remove-WPSRemnantFiles
                            Write-Host "`n按任意键返回..." -ForegroundColor Green
                            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                        }
                    }
                    elseif ($cleanChoice -eq '2') {
                        Write-Host "`n确认要清理 WPS 配置和缓存吗? (Y/N): " -ForegroundColor Yellow -NoNewline
                        $confirm = Read-Host
                        if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                            Remove-WPSRemnantFiles -KeepCloudFiles
                            Write-Host "`n按任意键返回..." -ForegroundColor Green
                            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                        }
                    }
                }
            }
            '0' { return }
            default { 
                Write-Host "`n无效的选择！" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($choice -ne '0')
}

# ==================== 重启资源管理器 ====================
function Handle-RestartExplorer {
    Clear-Host
    Write-Host "" -ForegroundColor Cyan
    Write-Host "         重启资源管理器                " -ForegroundColor Cyan
    Write-Host "" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "确认要重启资源管理器吗? (Y/N): " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    
    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        Restart-Explorer
        Write-Host "`n操作完成！按任意键返回..." -ForegroundColor Green
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# ==================== 主循环 ====================
do {
    Show-MainMenu
    $choice = Read-Host "请选择功能模块"
    
    switch ($choice) {
        '1' { Handle-RecentItemsMenu }
        '2' { Handle-QuickAccessMenu }
        '3' { Handle-OfficeRecentDocumentsMenu }
        '4' { Handle-OneDriveMenu }
        '5' { Handle-WPSOfficeMenu }
        '9' { Handle-RestartExplorer }
        '0' { 
            Clear-Host
            Write-Host "`n感谢使用 PowerClean！再见！`n" -ForegroundColor Green
            break 
        }
        default { 
            Write-Host "`n无效的选择，请重新选择。" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
} while ($choice -ne '0')
