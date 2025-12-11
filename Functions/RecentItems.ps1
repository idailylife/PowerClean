# PowerClean - RecentItems Module
# UTF-8 with BOM encoding for Chinese support

function Set-RecentItemsDisplay {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [bool]$Enabled,
        [Parameter(Mandatory=$false)]
        [switch]$RestartExplorer
    )
    
    try {
        $explorerAdvancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        
        if (-not (Test-Path $explorerAdvancedPath)) {
            New-Item -Path $explorerAdvancedPath -Force | Out-Null
        }
        
        if ($Enabled) {
            Write-Host "正在启用最近项目显示..." -ForegroundColor Green
            Set-ItemProperty -Path $explorerAdvancedPath -Name "Start_TrackDocs" -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $explorerAdvancedPath -Name "Start_TrackProgs" -Value 1 -Type DWord -Force
            Write-Host "已启用最近项目显示" -ForegroundColor Green
        }
        else {
            Write-Host "正在禁用最近项目显示..." -ForegroundColor Yellow
            Set-ItemProperty -Path $explorerAdvancedPath -Name "Start_TrackDocs" -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $explorerAdvancedPath -Name "Start_TrackProgs" -Value 0 -Type DWord -Force
            Write-Host "已禁用最近项目显示" -ForegroundColor Yellow
        }
        
        if ($RestartExplorer) {
            Write-Host "正在重启资源管理器以应用更改..." -ForegroundColor Cyan
            Restart-Explorer
        }
        else {
            Write-Host "提示: 更改将在重启资源管理器或重新登录后生效" -ForegroundColor Cyan
            Write-Host "可以运行 'Restart-Explorer' 立即应用更改" -ForegroundColor Cyan
        }
        
        return $true
    }
    catch {
        Write-Error "设置失败: $_"
        return $false
    }
}

function Enable-RecentItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [switch]$RestartExplorer
    )
    Set-RecentItemsDisplay -Enabled $true -RestartExplorer:$RestartExplorer
}

function Disable-RecentItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [switch]$RestartExplorer
    )
    Set-RecentItemsDisplay -Enabled $false -RestartExplorer:$RestartExplorer
}

function Get-RecentItemsStatus {
    [CmdletBinding()]
    param()
    
    try {
        $explorerAdvancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        $trackDocs = 1
        $trackProgs = 1
        
        if (Test-Path $explorerAdvancedPath) {
            $trackDocsValue = Get-ItemProperty -Path $explorerAdvancedPath -Name "Start_TrackDocs" -ErrorAction SilentlyContinue
            $trackProgsValue = Get-ItemProperty -Path $explorerAdvancedPath -Name "Start_TrackProgs" -ErrorAction SilentlyContinue
            
            if ($trackDocsValue) { $trackDocs = $trackDocsValue.Start_TrackDocs }
            if ($trackProgsValue) { $trackProgs = $trackProgsValue.Start_TrackProgs }
        }
        
        $isEnabled = ($trackDocs -eq 1) -and ($trackProgs -eq 1)
        $status = [PSCustomObject]@{
            Enabled = $isEnabled
            TrackDocuments = ($trackDocs -eq 1)
            TrackPrograms = ($trackProgs -eq 1)
            Status = if ($isEnabled) { "已启用" } else { "已禁用" }
        }
        
        Write-Host "
最近项目显示状态:" -ForegroundColor Cyan
        Write-Host "==================" -ForegroundColor Cyan
        Write-Host "状态: $($status.Status)" -ForegroundColor $(if ($isEnabled) { "Green" } else { "Yellow" })
        Write-Host "跟踪文档: $($status.TrackDocuments)" -ForegroundColor White
        Write-Host "跟踪程序: $($status.TrackPrograms)" -ForegroundColor White
        Write-Host ""
        
        return $status
    }
    catch {
        Write-Error "获取状态失败: $_"
        return $null
    }
}

function Restart-Explorer {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "正在重启资源管理器..." -ForegroundColor Yellow
        Stop-Process -Name explorer -Force -ErrorAction Stop
        Start-Sleep -Seconds 1
        Start-Process explorer
        Write-Host "资源管理器已重启" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "重启资源管理器失败: $_"
        Start-Process explorer -ErrorAction SilentlyContinue
        return $false
    }
}

Export-ModuleMember -Function Set-RecentItemsDisplay, Enable-RecentItems, Disable-RecentItems, Get-RecentItemsStatus, Restart-Explorer
