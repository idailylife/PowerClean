# WPS Office 安全管理功能
# 检测和卸载 WPS Office

<#
.SYNOPSIS
    检测 WPS Office 是否已安装
#>
function Test-WPSInstalled {
    $wpsBasePath = "HKCU:\Software\kingsoft\Office"
    
    # 检查注册表是否存在
    if (-not (Test-Path $wpsBasePath)) {
        return $false
    }
    
    # 检查版本路径
    $versionPaths = Get-ChildItem $wpsBasePath -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '\d+\.\d+' }
    if (-not $versionPaths) {
        return $false
    }
    
    # 检查安装路径是否有效
    $wpsCommonPath = Join-Path $versionPaths[0].PSPath "Common"
    if (Test-Path $wpsCommonPath) {
        $commonProps = Get-ItemProperty -Path $wpsCommonPath -ErrorAction SilentlyContinue
        if ($commonProps.InstallRoot -and (Test-Path $commonProps.InstallRoot -ErrorAction SilentlyContinue)) {
            return $true
        }
    }
    
    # 如果注册表存在但安装路径无效，说明已卸载
    return $false
}

<#
.SYNOPSIS
    获取 WPS Office 卸载程序路径
#>
function Get-WPSUninstallPath {
    # 检查 HKCU 卸载信息
    $uninstallPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    foreach ($path in $uninstallPaths) {
        if (Test-Path $path) {
            $apps = Get-ChildItem $path -ErrorAction SilentlyContinue | Get-ItemProperty -ErrorAction SilentlyContinue
            $wpsApp = $apps | Where-Object { $_.DisplayName -like "*WPS*" -or $_.DisplayName -like "*kingsoft*" }
            
            if ($wpsApp -and $wpsApp.UninstallString) {
                return $wpsApp.UninstallString
            }
        }
    }
    
    return $null
}

<#
.SYNOPSIS
    卸载 WPS Office
#>
function Uninstall-WPSOffice {
    Write-Host "`n正在检查 WPS Office 安装..." -ForegroundColor Cyan
    
    $uninstallPath = Get-WPSUninstallPath
    
    if (-not $uninstallPath) {
        Write-Host " [!] 未找到 WPS Office 卸载程序" -ForegroundColor Yellow
        return $false
    }
    
    Write-Host " [√] 找到 WPS Office 卸载程序" -ForegroundColor Green
    Write-Host "     路径: $uninstallPath" -ForegroundColor Gray
    Write-Host ""
    
    try {
        # 启动卸载程序
        Write-Host " [>] 正在启动 WPS Office 卸载程序..." -ForegroundColor Cyan
        Write-Host " [!] 请在弹出的卸载向导中完成卸载操作" -ForegroundColor Yellow
        Write-Host ""
        
        Start-Process -FilePath $uninstallPath -Wait
        
        # 检查是否卸载成功
        Start-Sleep -Seconds 2
        if (Test-WPSInstalled) {
            Write-Host " [!] WPS Office 仍然存在，可能未完成卸载" -ForegroundColor Yellow
            return $false
        } else {
            Write-Host " [√] WPS Office 已成功卸载" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host " [×] 卸载失败: $_" -ForegroundColor Red
        return $false
    }
}

<#
.SYNOPSIS
    扫描 WPS Office 遗留的文档和数据
#>
function Get-WPSRemnantFiles {
    Write-Host "`n正在扫描 WPS Office 遗留文件..." -ForegroundColor Cyan
    Write-Host ""
    
    $remnantLocations = @{
        "WPS云文档" = "$env:USERPROFILE\WPS Cloud Files"
        "WPS云文档(备选)" = "$env:USERPROFILE\Documents\WPS Cloud Files"
        "WPS配置数据" = "$env:APPDATA\kingsoft"
        "WPS本地数据" = "$env:LOCALAPPDATA\Kingsoft"
    }
    
    $foundItems = @()
    $totalSize = 0
    $totalFiles = 0
    
    foreach ($location in $remnantLocations.GetEnumerator()) {
        if (Test-Path $location.Value) {
            try {
                $items = Get-ChildItem $location.Value -Recurse -File -ErrorAction SilentlyContinue
                $fileCount = ($items | Measure-Object).Count
                $sizeBytes = ($items | Measure-Object -Property Length -Sum).Sum
                $sizeMB = [math]::Round($sizeBytes / 1MB, 2)
                
                if ($fileCount -gt 0) {
                    $foundItems += [PSCustomObject]@{
                        Location = $location.Key
                        Path = $location.Value
                        FileCount = $fileCount
                        SizeMB = $sizeMB
                    }
                    
                    $totalFiles += $fileCount
                    $totalSize += $sizeMB
                    
                    Write-Host "  [$($location.Key)]" -ForegroundColor Yellow
                    Write-Host "    路径: $($location.Value)" -ForegroundColor Gray
                    Write-Host "    文件数: $fileCount 个" -ForegroundColor Cyan
                    Write-Host "    大小: $sizeMB MB" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
            catch {
                Write-Host "  [!] 无法访问: $($location.Value)" -ForegroundColor Red
            }
        }
    }
    
    if ($foundItems.Count -eq 0) {
        Write-Host " [√] 未发现 WPS Office 遗留文件" -ForegroundColor Green
        return $null
    }
    else {
        Write-Host " [!] 汇总:" -ForegroundColor Yellow
        Write-Host "    共发现 $($foundItems.Count) 个位置" -ForegroundColor Yellow
        Write-Host "    总文件数: $totalFiles 个" -ForegroundColor Yellow
        Write-Host "    总大小: $totalSize MB" -ForegroundColor Yellow
        
        return $foundItems
    }
}

<#
.SYNOPSIS
    清理 WPS Office 遗留的文档和数据
#>
function Remove-WPSRemnantFiles {
    param(
        [switch]$KeepCloudFiles
    )
    
    Write-Host "`n正在清理 WPS Office 遗留文件..." -ForegroundColor Cyan
    Write-Host ""
    
    $remnantLocations = @()
    
    # 始终清理配置和缓存
    $remnantLocations += @{
        Name = "WPS配置数据"
        Path = "$env:APPDATA\kingsoft"
    }
    $remnantLocations += @{
        Name = "WPS本地数据"
        Path = "$env:LOCALAPPDATA\Kingsoft"
    }
    
    # 根据参数决定是否清理云文档
    if (-not $KeepCloudFiles) {
        $remnantLocations += @{
            Name = "WPS云文档"
            Path = "$env:USERPROFILE\WPS Cloud Files"
        }
        $remnantLocations += @{
            Name = "WPS云文档(备选)"
            Path = "$env:USERPROFILE\Documents\WPS Cloud Files"
        }
    }
    
    $cleaned = 0
    $failed = 0
    $accessDenied = 0
    
    foreach ($location in $remnantLocations) {
        if (Test-Path $location.Path) {
            try {
                Write-Host "  [>] 正在删除: $($location.Name)" -ForegroundColor Cyan
                Write-Host "      $($location.Path)" -ForegroundColor Gray
                
                # 尝试获取文件夹访问权限
                $acl = Get-Acl $location.Path -ErrorAction Stop
                
                # 尝试删除
                Remove-Item -Path $location.Path -Recurse -Force -ErrorAction Stop
                
                Write-Host "  [√] 已删除" -ForegroundColor Green
                Write-Host ""
                $cleaned++
            }
            catch [System.UnauthorizedAccessException] {
                Write-Host "  [×] 访问被拒绝：需要管理员权限" -ForegroundColor Red
                Write-Host ""
                $failed++
                $accessDenied++
            }
            catch {
                if ($_.Exception.Message -like "*访问*拒绝*" -or $_.Exception.Message -like "*Access*denied*") {
                    Write-Host "  [×] 访问被拒绝：需要管理员权限" -ForegroundColor Red
                    $accessDenied++
                } else {
                    Write-Host "  [×] 删除失败: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""
                $failed++
            }
        }
    }
    
    Write-Host " [!] 清理完成" -ForegroundColor Cyan
    Write-Host "    成功: $cleaned 项" -ForegroundColor Green
    if ($failed -gt 0) {
        Write-Host "    失败: $failed 项" -ForegroundColor Red
        
        if ($accessDenied -gt 0) {
            Write-Host ""
            Write-Host " [!] 提示：部分文件需要管理员权限才能删除" -ForegroundColor Yellow
            Write-Host "     请以管理员身份运行 PowerShell 后重试" -ForegroundColor Yellow
        }
    }
    
    return ($failed -eq 0)
}

# 导出函数
Export-ModuleMember -Function Test-WPSInstalled, Get-WPSUninstallPath, Uninstall-WPSOffice, Get-WPSRemnantFiles, Remove-WPSRemnantFiles
