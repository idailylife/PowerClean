# PowerClean - Office Version Detection Module
# UTF-8 with BOM encoding for Chinese support

function Get-OfficeVersion {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "正在检测Office版本..." -ForegroundColor Cyan
        
        $officeVersions = @()
        
        # 检查Microsoft 365新安装方式 (通过注册表 HKLM\Software\Microsoft\Office\16.0)
        $registryPaths365 = @(
            "HKLM:\Software\Microsoft\Office\16.0",
            "HKLM:\Software\WOW6432Node\Microsoft\Office\16.0"
        )
        
        foreach ($path in $registryPaths365) {
            if (Test-Path $path) {
                try {
                    $office16Keys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                    foreach ($key in $office16Keys) {
                        $identityPath = Join-Path -Path $key.PSPath -ChildPath "Common\Identity"
                        if (Test-Path $identityPath) {
                            $productName = Get-ItemProperty -Path $identityPath -Name "ProductName" -ErrorAction SilentlyContinue
                            if ($productName) {
                                $officeVersions += [PSCustomObject]@{
                                    ProductName = $productName.ProductName
                                    Version = "Office 16 (2019/2021/365)"
                                    Architecture = if ($path -like "*WOW6432Node*") { "32-bit" } else { "64-bit" }
                                    Type = "Registry"
                                    RegistryPath = $identityPath
                                }
                            }
                        }
                    }
                }
                catch {
                    # 继续检查其他路径
                }
            }
        }
        
        # 检查其他版本
        $registryPathsOther = @(
            "HKLM:\Software\Microsoft\Office\15.0\Common\Identity",
            "HKLM:\Software\Microsoft\Office\14.0\Common\Identity",
            "HKLM:\Software\WOW6432Node\Microsoft\Office\15.0\Common\Identity",
            "HKLM:\Software\WOW6432Node\Microsoft\Office\14.0\Common\Identity"
        )
        
        foreach ($path in $registryPathsOther) {
            if (Test-Path $path) {
                try {
                    $productName = Get-ItemProperty -Path $path -Name "ProductName" -ErrorAction SilentlyContinue
                    if ($productName) {
                        $officeVersions += [PSCustomObject]@{
                            ProductName = $productName.ProductName
                            Version = $path -replace '.*\\(\d+\.\d+)\\.*', 'Office $1'
                            Architecture = if ($path -like "*WOW6432Node*") { "32-bit" } else { "64-bit" }
                            Type = "Registry"
                            RegistryPath = $path
                        }
                    }
                }
                catch {
                    # 继续检查其他路径
                }
            }
        }
        
        # 检查Microsoft 365安装目录 (新的安装路径)
        $microsoftOfficeRootPath = "C:\Program Files\Microsoft Office\root"
        $microsoftOfficeRootPath86 = "C:\Program Files (x86)\Microsoft Office\root"
        
        $officePaths = @($microsoftOfficeRootPath, $microsoftOfficeRootPath86)
        
        foreach ($basePath in $officePaths) {
            if (Test-Path $basePath) {
                try {
                    # 检查Office16目录
                    $office16Path = Join-Path -Path $basePath -ChildPath "Office16"
                    if (Test-Path $office16Path) {
                        # 检查版本文件
                        $versionFile = Get-ChildItem -Path $office16Path -Filter "excelcnv.exe" -ErrorAction SilentlyContinue
                        if ($versionFile) {
                            $officeVersions += [PSCustomObject]@{
                                ProductName = "Microsoft 365"
                                Version = $versionFile.VersionInfo.FileVersion
                                Architecture = if ($basePath -like "*Program Files (x86)*") { "32-bit" } else { "64-bit" }
                                Type = "FileSystem"
                                InstallPath = $basePath
                            }
                        }
                    }
                }
                catch {
                    # 继续检查其他路径
                }
            }
        }
        
        # 检查传统Office安装目录
        $tradionalOfficePaths = @(
            "C:\Program Files\Microsoft Office",
            "C:\Program Files (x86)\Microsoft Office",
            "C:\Program Files\Microsoft Office 16",
            "C:\Program Files (x86)\Microsoft Office 16"
        )
        
        foreach ($path in $tradionalOfficePaths) {
            if (Test-Path $path) {
                try {
                    Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                        if ($_.Name -like "Office*" -or $_.Name -eq "root") {
                            Write-Host "检测到Office安装目录: $($_.FullName)" -ForegroundColor White
                        }
                    }
                }
                catch {
                    # 继续
                }
            }
        }
        
        # 输出检测结果
        if ($officeVersions.Count -gt 0) {
            Write-Host ""
            Write-Host "检测到的Office版本:" -ForegroundColor Green
            Write-Host "==================" -ForegroundColor Green
            
            # 去重
            $uniqueVersions = $officeVersions | Sort-Object -Property ProductName -Unique
            
            foreach ($office in $uniqueVersions) {
                Write-Host "产品名称: $($office.ProductName)" -ForegroundColor Green
                Write-Host "版本: $($office.Version)" -ForegroundColor White
                Write-Host "架构: $($office.Architecture)" -ForegroundColor White
                Write-Host "检测方式: $($office.Type)" -ForegroundColor Cyan
                if ($office.RegistryPath) {
                    Write-Host "注册表路径: $($office.RegistryPath)" -ForegroundColor Gray
                }
                if ($office.InstallPath) {
                    Write-Host "安装路径: $($office.InstallPath)" -ForegroundColor Gray
                }
                Write-Host "---" -ForegroundColor Gray
            }
            
            return $uniqueVersions
        }
        else {
            Write-Host ""
            Write-Host "未检测到安装的Office版本" -ForegroundColor Yellow
            Write-Host "请确保您已安装Microsoft Office" -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        Write-Error "检测Office版本失败: $_"
        return $null
    }
}

function Test-OfficeInstalled {
    [CmdletBinding()]
    param()
    
    try {
        $isInstalled = $false
        
        # 检查Microsoft 365安装目录
        $microsoftOfficeRootPath = "C:\Program Files\Microsoft Office\root"
        $microsoftOfficeRootPath86 = "C:\Program Files (x86)\Microsoft Office\root"
        
        if ((Test-Path $microsoftOfficeRootPath) -or (Test-Path $microsoftOfficeRootPath86)) {
            $isInstalled = $true
        }
        
        # 检查注册表路径
        if (-not $isInstalled) {
            $registryPaths = @(
                "HKLM:\Software\Microsoft\Office\16.0\Common\Identity",
                "HKLM:\Software\Microsoft\Office\15.0\Common\Identity",
                "HKLM:\Software\Microsoft\Office\14.0\Common\Identity",
                "HKLM:\Software\WOW6432Node\Microsoft\Office\16.0\Common\Identity",
                "HKLM:\Software\WOW6432Node\Microsoft\Office\15.0\Common\Identity",
                "HKLM:\Software\WOW6432Node\Microsoft\Office\14.0\Common\Identity"
            )
            
            foreach ($path in $registryPaths) {
                if (Test-Path $path) {
                    $productName = Get-ItemProperty -Path $path -Name "ProductName" -ErrorAction SilentlyContinue
                    if ($productName) {
                        $isInstalled = $true
                        break
                    }
                }
            }
        }
        
        if ($isInstalled) {
            Write-Host "已检测到安装的Office" -ForegroundColor Green
        }
        else {
            Write-Host "未检测到安装的Office" -ForegroundColor Yellow
        }
        
        return $isInstalled
    }
    catch {
        Write-Error "检测失败: $_"
        return $false
    }
}

function Get-OfficeInstallPath {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "正在查询Office安装路径..." -ForegroundColor Cyan
        
        $installPaths = @()
        
        # 首先检查Microsoft 365新安装路径
        $microsoftOfficeRootPath = "C:\Program Files\Microsoft Office\root"
        $microsoftOfficeRootPath86 = "C:\Program Files (x86)\Microsoft Office\root"
        
        if (Test-Path $microsoftOfficeRootPath) {
            Write-Host "Office安装路径 (64-bit): $microsoftOfficeRootPath" -ForegroundColor Green
            $installPaths += $microsoftOfficeRootPath
        }
        
        if (Test-Path $microsoftOfficeRootPath86) {
            Write-Host "Office安装路径 (32-bit): $microsoftOfficeRootPath86" -ForegroundColor Green
            $installPaths += $microsoftOfficeRootPath86
        }
        
        # 检查传统的注册表路径
        $registryPaths = @(
            "HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot",
            "HKLM:\Software\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot"
        )
        
        foreach ($registryPath in $registryPaths) {
            if (Test-Path $registryPath) {
                try {
                    $pathInfo = Get-ItemProperty -Path $registryPath -Name "Path" -ErrorAction SilentlyContinue
                    if ($pathInfo -and $pathInfo.Path) {
                        if ($installPaths -notcontains $pathInfo.Path) {
                            Write-Host "Office安装路径: $($pathInfo.Path)" -ForegroundColor Green
                            $installPaths += $pathInfo.Path
                        }
                    }
                }
                catch {
                    # 继续
                }
            }
        }
        
        if ($installPaths.Count -eq 0) {
            Write-Host "未找到Office安装路径" -ForegroundColor Yellow
            return $null
        }
        
        return $installPaths
    }
    catch {
        Write-Error "查询安装路径失败: $_"
        return $null
    }
}

function Disable-OfficeRecentDocuments {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "正在禁用Office最近文档列表..." -ForegroundColor Yellow
        
        # Office所有应用程序的注册表位置
        $officeApps = @(
            "Word",
            "Excel",
            "PowerPoint",
            "Access",
            "Publisher",
            "Outlook",
            "OneNote",
            "Visio",
            "Project"
        )
        
        $disabledCount = 0
        
        # 禁用每个Office应用的最近文档功能
        foreach ($app in $officeApps) {
            try {
                # 注册表路径: HKCU\Software\Microsoft\Office\16.0\{AppName}\Options
                $registryPath = "HKCU:\Software\Microsoft\Office\16.0\$app\Options"
                
                # 创建路径如果不存在
                if (-not (Test-Path $registryPath)) {
                    New-Item -Path $registryPath -Force | Out-Null
                }
                
                # 关键设置1: 在File菜单隐藏最近文档数量
                Set-ItemProperty -Path $registryPath -Name "RecentFilesListSize" -Value 0 -Type DWord -Force
                
                # 关键设置2: 在Backstage视图隐藏最近位置数量
                Set-ItemProperty -Path $registryPath -Name "RecentPlacesListSize" -Value 0 -Type DWord -Force
                
                # 设置3: 禁用MRU (Most Recently Used)
                Set-ItemProperty -Path $registryPath -Name "MRUMaxCount" -Value 0 -Type DWord -Force
                
                # 设置4: 不显示最近文档
                Set-ItemProperty -Path $registryPath -Name "DoNotShowRecentDocuments" -Value 1 -Type DWord -Force
                
                # 设置5: 禁用在开始屏幕显示最近使用
                Set-ItemProperty -Path $registryPath -Name "ShowStartupDialog" -Value 0 -Type DWord -Force
                
                Write-Host "✓ 已禁用 $app 的最近文档列表" -ForegroundColor Green
                $disabledCount++
            }
            catch {
                Write-Host "- 无法禁用 $app (可能未安装): $($_.Exception.Message)" -ForegroundColor Gray
            }
        }
        
        # 禁用File菜单中的最近文档列表 - 全局设置
        $fileOptionsPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"
        if (-not (Test-Path $fileOptionsPath)) {
            New-Item -Path $fileOptionsPath -Force | Out-Null
        }
        Set-ItemProperty -Path $fileOptionsPath -Name "ShowRecentDocs" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $fileOptionsPath -Name "RecentFileCount" -Value 0 -Type DWord -Force
        
        # 禁用最近位置
        $openSavePath = "HKCU:\Software\Microsoft\Office\16.0\Common\Open Find"
        if (-not (Test-Path $openSavePath)) {
            New-Item -Path $openSavePath -Force | Out-Null
        }
        Set-ItemProperty -Path $openSavePath -Name "DisableRecentFiles" -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $openSavePath -Name "DisableRecentPlaces" -Value 1 -Type DWord -Force
        
        # 清空现有的最近文档列表（可选）
        Write-Host ""
        Write-Host "正在清除现有的最近文档记录..." -ForegroundColor Yellow
        
        foreach ($app in $officeApps) {
            try {
                $fileMRUPath = "HKCU:\Software\Microsoft\Office\16.0\$app\File MRU"
                $placeMRUPath = "HKCU:\Software\Microsoft\Office\16.0\$app\Place MRU"
                
                if (Test-Path $fileMRUPath) {
                    Remove-Item -Path $fileMRUPath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "  清除了 $app 的文件MRU记录" -ForegroundColor Gray
                }
                
                if (Test-Path $placeMRUPath) {
                    Remove-Item -Path $placeMRUPath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "  清除了 $app 的位置MRU记录" -ForegroundColor Gray
                }
            }
            catch {
                # 继续处理其他应用
            }
        }
        
        Write-Host ""
        Write-Host "已成功禁用 $disabledCount 个Office应用的最近文档列表" -ForegroundColor Green
        Write-Host "已禁用以下功能:" -ForegroundColor Cyan
        Write-Host "  • File菜单的最近文档列表" -ForegroundColor White
        Write-Host "  • 打开页面的最近文档" -ForegroundColor White
        Write-Host "  • 最近位置列表" -ForegroundColor White
        Write-Host "  • 开始屏幕的最近使用" -ForegroundColor White
        Write-Host ""
        Write-Host "重要提示: 必须完全关闭所有Office应用后重新打开才能生效！" -ForegroundColor Yellow
        
        return $true
    }
    catch {
        Write-Error "禁用Office最近文档列表失败: $_"
        return $false
    }
}

function Enable-OfficeRecentDocuments {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "正在启用Office最近文档列表..." -ForegroundColor Cyan
        
        # Office所有应用程序
        $officeApps = @(
            "Word",
            "Excel",
            "PowerPoint",
            "Access",
            "Publisher",
            "Outlook",
            "OneNote",
            "Visio",
            "Project"
        )
        
        $enabledCount = 0
        
        # 为每个Office应用启用最近文档功能
        foreach ($app in $officeApps) {
            try {
                $registryPath = "HKCU:\Software\Microsoft\Office\16.0\$app\Options"
                
                # 创建路径如果不存在
                if (-not (Test-Path $registryPath)) {
                    New-Item -Path $registryPath -Force | Out-Null
                }
                
                # 设置1: 在File菜单显示最近文档数量（25个）
                Set-ItemProperty -Path $registryPath -Name "RecentFilesListSize" -Value 25 -Type DWord -Force
                
                # 设置2: 在Backstage视图显示最近位置数量（5个）
                Set-ItemProperty -Path $registryPath -Name "RecentPlacesListSize" -Value 5 -Type DWord -Force
                
                # 设置3: 启用MRU，显示20个文档
                Set-ItemProperty -Path $registryPath -Name "MRUMaxCount" -Value 20 -Type DWord -Force
                
                # 设置4: 显示最近文档
                Set-ItemProperty -Path $registryPath -Name "DoNotShowRecentDocuments" -Value 0 -Type DWord -Force
                
                # 设置5: 启用开始屏幕
                Set-ItemProperty -Path $registryPath -Name "ShowStartupDialog" -Value 1 -Type DWord -Force
                
                Write-Host "✓ 已启用 $app 的最近文档列表" -ForegroundColor Green
                $enabledCount++
            }
            catch {
                Write-Host "- 无法启用 $app (可能未安装): $($_.Exception.Message)" -ForegroundColor Gray
            }
        }
        
        # 启用File菜单中的最近文档列表 - 全局设置
        $fileOptionsPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"
        if (-not (Test-Path $fileOptionsPath)) {
            New-Item -Path $fileOptionsPath -Force | Out-Null
        }
        Set-ItemProperty -Path $fileOptionsPath -Name "ShowRecentDocs" -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $fileOptionsPath -Name "RecentFileCount" -Value 25 -Type DWord -Force
        
        # 启用最近位置
        $openSavePath = "HKCU:\Software\Microsoft\Office\16.0\Common\Open Find"
        if (-not (Test-Path $openSavePath)) {
            New-Item -Path $openSavePath -Force | Out-Null
        }
        Set-ItemProperty -Path $openSavePath -Name "DisableRecentFiles" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $openSavePath -Name "DisableRecentPlaces" -Value 0 -Type DWord -Force
        
        Write-Host ""
        Write-Host "已成功启用 $enabledCount 个Office应用的最近文档列表" -ForegroundColor Green
        Write-Host "已启用以下功能:" -ForegroundColor Cyan
        Write-Host "  • File菜单的最近文档列表（25个）" -ForegroundColor White
        Write-Host "  • 打开页面的最近文档" -ForegroundColor White
        Write-Host "  • 最近位置列表（5个）" -ForegroundColor White
        Write-Host "  • 开始屏幕的最近使用" -ForegroundColor White
        Write-Host ""
        Write-Host "提示: 更改将在重启Office应用后生效" -ForegroundColor Cyan
        
        return $true
    }
    catch {
        Write-Error "启用Office最近文档列表失败: $_"
        return $false
    }
}

function Get-OfficeRecentDocumentsStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "正在检查Office最近文档列表状态..." -ForegroundColor Cyan
        Write-Host ""
        
        # 检查主要设置
        $fileOptionsPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"
        $showRecentDocs = 1  # 默认启用
        
        if (Test-Path $fileOptionsPath) {
            $setting = Get-ItemProperty -Path $fileOptionsPath -Name "ShowRecentDocs" -ErrorAction SilentlyContinue
            if ($setting) {
                $showRecentDocs = $setting.ShowRecentDocs
            }
        }
        
        # 检查各个应用的设置
        $officeApps = @("Word", "Excel", "PowerPoint", "Access", "Publisher", "Outlook", "OneNote", "Visio", "Project")
        $appStatus = @()
        
        foreach ($app in $officeApps) {
            $registryPath = "HKCU:\Software\Microsoft\Office\16.0\$app\Options"
            $mruCount = 20  # 默认值
            $doNotShow = 0
            
            if (Test-Path $registryPath) {
                $mruSetting = Get-ItemProperty -Path $registryPath -Name "MRUMaxCount" -ErrorAction SilentlyContinue
                $doNotShowSetting = Get-ItemProperty -Path $registryPath -Name "DoNotShowRecentDocuments" -ErrorAction SilentlyContinue
                
                if ($mruSetting) { $mruCount = $mruSetting.MRUMaxCount }
                if ($doNotShowSetting) { $doNotShow = $doNotShowSetting.DoNotShowRecentDocuments }
            }
            
            $isEnabled = ($mruCount -gt 0) -and ($doNotShow -eq 0)
            $appStatus += [PSCustomObject]@{
                Application = $app
                Enabled = $isEnabled
                MRUCount = if ($isEnabled) { $mruCount } else { 0 }
                Status = if ($isEnabled) { "已启用" } else { "已禁用" }
            }
        }
        
        # 显示状态
        Write-Host "Office最近文档列表状态:" -ForegroundColor Cyan
        Write-Host "========================" -ForegroundColor Cyan
        Write-Host "全局设置 (File菜单显示最近文档): $(if ($showRecentDocs -eq 1) { '启用' } else { '禁用' })" -ForegroundColor White
        Write-Host ""
        Write-Host "各应用程序状态:" -ForegroundColor Cyan
        Write-Host "---" -ForegroundColor Gray
        
        foreach ($status in $appStatus) {
            $statusColor = if ($status.Enabled) { "Green" } else { "Yellow" }
            Write-Host "$($status.Application): $($status.Status) (MRU数量: $($status.MRUCount))" -ForegroundColor $statusColor
        }
        
        Write-Host ""
        
        return [PSCustomObject]@{
            GlobalSetting = $showRecentDocs
            ApplicationStatus = $appStatus
            Summary = "Office最近文档列表状态已检查"
        }
    }
    catch {
        Write-Error "检查状态失败: $_"
        return $null
    }
}
