# PowerClean - QuickAccess Module

function Set-QuickAccessRecentFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [bool]$Enabled,
        [Parameter(Mandatory=$false)]
        [switch]$RestartExplorer
    )
    
    try {
        $explorerPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"
        
        if (-not (Test-Path $explorerPath)) {
            New-Item -Path $explorerPath -Force | Out-Null
        }
        
        if ($Enabled) {
            Write-Host "正在启用快速访问最近文件显示..." -ForegroundColor Green
            Set-ItemProperty -Path $explorerPath -Name "ShowRecent" -Value 1 -Type DWord -Force
            Write-Host "已启用快速访问最近文件显示" -ForegroundColor Green
        }
        else {
            Write-Host "正在禁用快速访问最近文件显示..." -ForegroundColor Yellow
            Set-ItemProperty -Path $explorerPath -Name "ShowRecent" -Value 0 -Type DWord -Force
            Write-Host "已禁用快速访问最近文件显示" -ForegroundColor Yellow
        }
        
        if ($RestartExplorer) {
            Write-Host "正在重启资源管理器以应用更改..." -ForegroundColor Cyan
            Restart-Explorer
        }
        else {
            Write-Host "提示: 更改将在重启资源管理器或重新登录后生效" -ForegroundColor Cyan
        }
        
        return $true
    }
    catch {
        Write-Error "设置失败: $_"
        return $false
    }
}

function Set-QuickAccessFrequentFolders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [bool]$Enabled,
        [Parameter(Mandatory=$false)]
        [switch]$RestartExplorer
    )
    
    try {
        $explorerPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"
        
        if (-not (Test-Path $explorerPath)) {
            New-Item -Path $explorerPath -Force | Out-Null
        }
        
        if ($Enabled) {
            Write-Host "正在启用快速访问常用文件夹显示..." -ForegroundColor Green
            Set-ItemProperty -Path $explorerPath -Name "ShowFrequent" -Value 1 -Type DWord -Force
            Write-Host "已启用快速访问常用文件夹显示" -ForegroundColor Green
        }
        else {
            Write-Host "正在禁用快速访问常用文件夹显示..." -ForegroundColor Yellow
            Set-ItemProperty -Path $explorerPath -Name "ShowFrequent" -Value 0 -Type DWord -Force
            Write-Host "已禁用快速访问常用文件夹显示" -ForegroundColor Yellow
        }
        
        if ($RestartExplorer) {
            Write-Host "正在重启资源管理器以应用更改..." -ForegroundColor Cyan
            Restart-Explorer
        }
        else {
            Write-Host "提示: 更改将在重启资源管理器或重新登录后生效" -ForegroundColor Cyan
        }
        
        return $true
    }
    catch {
        Write-Error "设置失败: $_"
        return $false
    }
}

function Enable-QuickAccessRecentFiles {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][switch]$RestartExplorer)
    Set-QuickAccessRecentFiles -Enabled $true -RestartExplorer:$RestartExplorer
}

function Disable-QuickAccessRecentFiles {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][switch]$RestartExplorer)
    Set-QuickAccessRecentFiles -Enabled $false -RestartExplorer:$RestartExplorer
}

function Enable-QuickAccessFrequentFolders {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][switch]$RestartExplorer)
    Set-QuickAccessFrequentFolders -Enabled $true -RestartExplorer:$RestartExplorer
}

function Disable-QuickAccessFrequentFolders {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][switch]$RestartExplorer)
    Set-QuickAccessFrequentFolders -Enabled $false -RestartExplorer:$RestartExplorer
}

function Get-QuickAccessStatus {
    [CmdletBinding()]
    param()
    
    try {
        $explorerPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"
        $showRecent = 1
        $showFrequent = 1
        
        if (Test-Path $explorerPath) {
            $recentValue = Get-ItemProperty -Path $explorerPath -Name "ShowRecent" -ErrorAction SilentlyContinue
            $frequentValue = Get-ItemProperty -Path $explorerPath -Name "ShowFrequent" -ErrorAction SilentlyContinue
            
            if ($recentValue) { $showRecent = $recentValue.ShowRecent }
            if ($frequentValue) { $showFrequent = $frequentValue.ShowFrequent }
        }
        
        $recentEnabled = ($showRecent -eq 1)
        $frequentEnabled = ($showFrequent -eq 1)
        
        $status = [PSCustomObject]@{
            RecentFilesEnabled = $recentEnabled
            FrequentFoldersEnabled = $frequentEnabled
            RecentFilesStatus = if ($recentEnabled) { "已启用" } else { "已禁用" }
            FrequentFoldersStatus = if ($frequentEnabled) { "已启用" } else { "已禁用" }
        }
        
        Write-Host "
快速访问设置状态:" -ForegroundColor Cyan
        Write-Host "==================" -ForegroundColor Cyan
        Write-Host "最近使用的文件: $($status.RecentFilesStatus)" -ForegroundColor $(if ($recentEnabled) { "Green" } else { "Yellow" })
        Write-Host "常用文件夹: $($status.FrequentFoldersStatus)" -ForegroundColor $(if ($frequentEnabled) { "Green" } else { "Yellow" })
        Write-Host ""
        
        return $status
    }
    catch {
        Write-Error "获取状态失败: $_"
        return $null
    }
}

Export-ModuleMember -Function Set-QuickAccessRecentFiles, Set-QuickAccessFrequentFolders, Enable-QuickAccessRecentFiles, Disable-QuickAccessRecentFiles, Enable-QuickAccessFrequentFolders, Disable-QuickAccessFrequentFolders, Get-QuickAccessStatus
