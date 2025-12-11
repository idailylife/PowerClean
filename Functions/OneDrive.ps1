# OneDrive 同步管理功能
# 控制 OneDrive 个人版和商业版的同步设置

<#
.SYNOPSIS
    设置 OneDrive 同步功能的启用或禁用状态
    
.DESCRIPTION
    通过修改注册表来控制 OneDrive 的个人版和商业版同步功能，以及在文件资源管理器中的显示
    
.PARAMETER PersonalSync
    是否启用 OneDrive 个人版同步
    
.PARAMETER BusinessSync
    是否启用 OneDrive 商业版同步
    
.PARAMETER ShowInExplorer
    是否在文件资源管理器中显示 OneDrive
    
.EXAMPLE
    Set-OneDriveSync -PersonalSync $false
    禁用 OneDrive 个人版同步
#>
function Set-OneDriveSync {
    param(
        [Parameter(Mandatory=$false)]
        [bool]$PersonalSync,
        
        [Parameter(Mandatory=$false)]
        [bool]$BusinessSync,
        
        [Parameter(Mandatory=$false)]
        [bool]$ShowInExplorer
    )
    
    try {
        $regPath = "HKCU:\Software\Microsoft\OneDrive"
        
        # 确保注册表路径存在
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        # 设置个人版同步
        if ($PSBoundParameters.ContainsKey('PersonalSync')) {
            $value = if ($PersonalSync) { 0 } else { 1 }
            Set-ItemProperty -Path $regPath -Name "DisablePersonalSync" -Value $value -Type DWord -Force
            
            if ($PersonalSync) {
                Write-Host " [√] 已启用 OneDrive 个人版同步" -ForegroundColor Green
            } else {
                Write-Host " [√] 已禁用 OneDrive 个人版同步" -ForegroundColor Yellow
            }
        }
        
        # 设置商业版同步
        if ($PSBoundParameters.ContainsKey('BusinessSync')) {
            $value = if ($BusinessSync) { 0 } else { 1 }
            Set-ItemProperty -Path $regPath -Name "DisableBusinessSync" -Value $value -Type DWord -Force
            
            if ($BusinessSync) {
                Write-Host " [√] 已启用 OneDrive 商业版同步" -ForegroundColor Green
            } else {
                Write-Host " [√] 已禁用 OneDrive 商业版同步" -ForegroundColor Yellow
            }
        }
        
        # 设置文件资源管理器中的显示
        if ($PSBoundParameters.ContainsKey('ShowInExplorer')) {
            $clsidPath = "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
            
            # 确保 CLSID 路径存在
            if (-not (Test-Path $clsidPath)) {
                New-Item -Path $clsidPath -Force | Out-Null
            }
            
            $value = if ($ShowInExplorer) { 1 } else { 0 }
            Set-ItemProperty -Path $clsidPath -Name "System.IsPinnedToNameSpaceTree" -Value $value -Type DWord -Force
            
            if ($ShowInExplorer) {
                Write-Host " [√] OneDrive 将在文件资源管理器中显示" -ForegroundColor Green
            } else {
                Write-Host " [√] OneDrive 已从文件资源管理器中隐藏" -ForegroundColor Yellow
            }
        }
        
        return $true
    }
    catch {
        Write-Host " [×] 操作失败: $_" -ForegroundColor Red
        return $false
    }
}

<#
.SYNOPSIS
    启用 OneDrive 个人版同步
#>
function Enable-OneDrivePersonalSync {
    Write-Host "`n正在启用 OneDrive 个人版同步..." -ForegroundColor Cyan
    return Set-OneDriveSync -PersonalSync $true
}

<#
.SYNOPSIS
    禁用 OneDrive 个人版同步
#>
function Disable-OneDrivePersonalSync {
    Write-Host "`n正在禁用 OneDrive 个人版同步..." -ForegroundColor Cyan
    return Set-OneDriveSync -PersonalSync $false
}

<#
.SYNOPSIS
    启用 OneDrive 商业版同步
#>
function Enable-OneDriveBusinessSync {
    Write-Host "`n正在启用 OneDrive 商业版同步..." -ForegroundColor Cyan
    return Set-OneDriveSync -BusinessSync $true
}

<#
.SYNOPSIS
    禁用 OneDrive 商业版同步
#>
function Disable-OneDriveBusinessSync {
    Write-Host "`n正在禁用 OneDrive 商业版同步..." -ForegroundColor Cyan
    return Set-OneDriveSync -BusinessSync $false
}

<#
.SYNOPSIS
    在文件资源管理器中显示 OneDrive
#>
function Show-OneDriveInExplorer {
    Write-Host "`n正在设置 OneDrive 在文件资源管理器中显示..." -ForegroundColor Cyan
    $result = Set-OneDriveSync -ShowInExplorer $true
    if ($result) {
        Write-Host " [!] 需要重启资源管理器以应用更改" -ForegroundColor Yellow
    }
    return $result
}

<#
.SYNOPSIS
    在文件资源管理器中隐藏 OneDrive
#>
function Hide-OneDriveInExplorer {
    Write-Host "`n正在从文件资源管理器中隐藏 OneDrive..." -ForegroundColor Cyan
    $result = Set-OneDriveSync -ShowInExplorer $false
    if ($result) {
        Write-Host " [!] 需要重启资源管理器以应用更改" -ForegroundColor Yellow
    }
    return $result
}

<#
.SYNOPSIS
    获取 OneDrive 当前的同步状态
#>
function Get-OneDriveStatus {
    Write-Host ""
    Write-Host "【当前状态】" -ForegroundColor Cyan
    Write-Host ""
    
    # 检查个人版同步状态
    $regPath = "HKCU:\Software\Microsoft\OneDrive"
    $personalDisabled = $false
    $businessDisabled = $false
    
    if (Test-Path $regPath) {
        $personalValue = Get-ItemProperty -Path $regPath -Name "DisablePersonalSync" -ErrorAction SilentlyContinue
        if ($personalValue -and $personalValue.DisablePersonalSync -eq 1) {
            $personalDisabled = $true
        }
        
        $businessValue = Get-ItemProperty -Path $regPath -Name "DisableBusinessSync" -ErrorAction SilentlyContinue
        if ($businessValue -and $businessValue.DisableBusinessSync -eq 1) {
            $businessDisabled = $true
        }
    }
    
    # 检查文件资源管理器显示状态
    $clsidPath = "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    $hiddenInExplorer = $false
    
    if (Test-Path $clsidPath) {
        $explorerValue = Get-ItemProperty -Path $clsidPath -Name "System.IsPinnedToNameSpaceTree" -ErrorAction SilentlyContinue
        if ($explorerValue -and $explorerValue.'System.IsPinnedToNameSpaceTree' -eq 0) {
            $hiddenInExplorer = $true
        }
    }
    
    # 显示个人版状态
    if ($personalDisabled) {
        Write-Host "  OneDrive 个人版同步: " -NoNewline
        Write-Host "已禁用" -ForegroundColor Yellow
    } else {
        Write-Host "  OneDrive 个人版同步: " -NoNewline
        Write-Host "已启用" -ForegroundColor Green
    }
    
    # 显示商业版状态
    if ($businessDisabled) {
        Write-Host "  OneDrive 商业版同步: " -NoNewline
        Write-Host "已禁用" -ForegroundColor Yellow
    } else {
        Write-Host "  OneDrive 商业版同步: " -NoNewline
        Write-Host "已启用" -ForegroundColor Green
    }
    
    # 显示文件资源管理器状态
    if ($hiddenInExplorer) {
        Write-Host "  文件资源管理器显示: " -NoNewline
        Write-Host "已隐藏" -ForegroundColor Yellow
    } else {
        Write-Host "  文件资源管理器显示: " -NoNewline
        Write-Host "已显示" -ForegroundColor Green
    }
    
    return [PSCustomObject]@{
        PersonalSyncEnabled = -not $personalDisabled
        BusinessSyncEnabled = -not $businessDisabled
        ShowInExplorer = -not $hiddenInExplorer
    }
}

# 导出函数
Export-ModuleMember -Function Set-OneDriveSync, Enable-OneDrivePersonalSync, Disable-OneDrivePersonalSync, Enable-OneDriveBusinessSync, Disable-OneDriveBusinessSync, Show-OneDriveInExplorer, Hide-OneDriveInExplorer, Get-OneDriveStatus
