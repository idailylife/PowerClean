# PowerClean - PowerShell系统管理工具库

一个用于管理Windows系统界面、隐私和安全的PowerShell模块，专为中文用户和企业环境设计。

## ✨ 功能特性

### 📋 开始菜单和任务栏 - 最近项目管理
- 控制开始菜单和任务栏右键菜单中"最近"列表的显示
- 支持开启/关闭最近使用的项目显示
- 查询当前状态
- 可选择是否自动重启资源管理器以立即应用更改

### 📁 文件资源管理器 - 快速访问管理
- 控制"最近使用的文件"显示
- 控制"常用文件夹"显示
- 独立管理两个功能的开关状态
- 实时查看当前配置

### 📄 Microsoft Office - 最近文档管理
- 自动检测已安装的Office版本
- 启用/禁用Office最近文档列表
- 支持多个Office版本
- 实时状态查询

### ☁️ OneDrive - 云同步管理
- 启用/禁用OneDrive个人版同步
- 启用/禁用OneDrive商业版同步
- 在文件资源管理器中显示/隐藏OneDrive
- 状态实时监控

### 🔒 WPS Office - 安全管理（企业功能）
- **安全检测**：自动检测WPS Office安装状态
- **安全警告**：显示潜在的数据泄露风险
- **一键卸载**：快速卸载WPS Office
- **遗留文件扫描**：扫描WPS云文档和配置文件
- **智能清理**：
  - 全部清理（包括云文档）
  - 保留云文档，仅清理配置和缓存
- **权限管理**：自动检测管理员权限，提示权限不足时的解决方案

## 🚀 快速开始

### 方法一：使用交互式菜单（推荐）⭐

**最简单的方式 - 双击启动：**
```
双击 "启动PowerClean.bat" 文件
```

**或在PowerShell中运行：**
```powershell
.\PowerClean.ps1
```

**以管理员身份运行（推荐）：**
```powershell
# 右键 PowerShell → 以管理员身份运行
.\PowerClean.ps1
```

交互式菜单提供：
- 🔍 实时状态查看
- 📋 最近项目管理
- 📁 快速访问管理
- 📄 Office文档管理
- ☁️ OneDrive同步管理
- 🔒 WPS安全管理
- 🔄 系统工具（重启资源管理器）
- 所有操作都有确认提示，安全便捷
- 管理员模式标识

详细使用说明请查看 [使用指南.md](使用指南.md)

### 方法二：直接使用命令

1. 导入模块：
```powershell
Import-Module .\PowerClean.psd1
```

2. 使用命令（见下方详细说明）

### 企业环境部署建议

对于企业IT管理员：
1. 以管理员身份运行PowerShell
2. 运行 `.\PowerClean.ps1`
3. 选择"5. WPS Office - 安全管理"
4. 检测并卸载WPS Office
5. 扫描和清理遗留文件

## 📖 使用方法

### 最近项目管理
```powershell
# 禁用最近项目显示
Disable-RecentItems

# 启用最近项目显示
Enable-RecentItems

# 查询当前状态
Get-RecentItemsStatus
```

### 快速访问管理
```powershell
# 禁用最近文件
Disable-QuickAccessRecentFiles

# 禁用常用文件夹
Disable-QuickAccessFrequentFolders

# 查询状态
Get-QuickAccessStatus
```

### Office文档管理
```powershell
# 禁用Office最近文档
Disable-OfficeRecentDocuments

# 启用Office最近文档
Enable-OfficeRecentDocuments

# 查询状态
Get-OfficeRecentDocumentsStatus
```

### OneDrive管理
```powershell
# 禁用OneDrive个人版同步
Disable-OneDrivePersonalSync

# 在文件资源管理器中隐藏OneDrive
Hide-OneDriveInExplorer

# 查询状态
Get-OneDriveStatus
```

### WPS Office管理
```powershell
# 检测WPS是否已安装
Test-WPSInstalled

# 扫描WPS遗留文件
Get-WPSRemnantFiles

# 卸载WPS Office
Uninstall-WPSOffice

# 清理WPS遗留文件（保留云文档）
Remove-WPSRemnantFiles -KeepCloudFiles
```

### 系统工具
```powershell
# 重启资源管理器
Restart-Explorer
```

## 快速使用脚本

### 主程序（交互式菜单）
```powershell
.\PowerClean.ps1
```
提供完整的交互式菜单界面，是最推荐的使用方式。

### 其他示例脚本

项目中还包含了快速脚本：

- **Quick-Disable.ps1** - 快速禁用最近项目显示
- **Quick-Enable.ps1** - 快速启用最近项目显示
- **Example-RecentItems.ps1** - 完整的交互式示例

直接运行这些脚本：
```powershell
.\Examples\Quick-Disable.ps1
```

## 📂 项目结构

```
power_clean/
├── PowerClean.ps1                  # 主程序入口（交互式菜单）
├── PowerClean.psm1                 # 主模块文件
├── PowerClean.psd1                 # 模块清单
├── 启动PowerClean.bat              # 一键启动脚本
├── README.md                       # 说明文档
├── 使用指南.md                     # 详细使用指南
├── Functions/                      # 功能函数目录
│   ├── RecentItems.ps1            # 最近项目管理
│   ├── QuickAccess.ps1            # 快速访问管理
│   ├── OfficeVersion.ps1          # Office检测和管理
│   ├── OneDrive.ps1               # OneDrive同步管理
│   └── WPSOffice.ps1              # WPS Office安全管理
└── Examples/                       # 示例脚本目录
    ├── Quick-Disable.ps1          # 快速禁用脚本
    ├── Quick-Enable.ps1           # 快速启用脚本
    ├── Example-RecentItems.ps1    # 最近项目示例
    ├── Example-OfficeVersion.ps1  # Office检测示例
    └── Example-OfficeRecentDocuments.ps1  # Office文档示例
```

## 💻 系统要求

- Windows 10/11
- PowerShell 5.1 或更高版本
- **推荐以管理员权限运行**（部分功能需要）
- 支持UTF-8编码以正确显示中文

## ⚙️ 工作原理

本模块通过修改Windows注册表来控制系统行为：

### 最近项目管理
- **Start_TrackDocs**: 控制开始菜单中最近打开的文档
- **Start_TrackProgs**: 控制开始菜单中最近使用的程序
- 路径: `HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced`

### 快速访问管理
- **ShowRecent**: 控制最近使用的文件显示
- **ShowFrequent**: 控制常用文件夹显示
- 路径: `HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer`

### Office管理
- 自动检测Office安装路径和版本
- 修改Office注册表配置
- 路径: `HKCU:\Software\Microsoft\Office\<Version>\Common`

### OneDrive管理
- **DisablePersonalSync**: 控制个人版同步
- **DisableBusinessSync**: 控制商业版同步
- **System.IsPinnedToNameSpaceTree**: 控制资源管理器显示
- 路径: `HKCU:\Software\Microsoft\OneDrive`

### WPS Office管理
- 检测WPS安装状态（验证安装路径）
- 扫描遗留文件位置：
  - `%USERPROFILE%\WPS Cloud Files` - 云文档
  - `%APPDATA%\kingsoft` - 配置数据
  - `%LOCALAPPDATA%\Kingsoft` - 本地数据
- 调用官方卸载程序确保干净卸载

## ⚠️ 注意事项

- 修改注册表设置后，某些功能需要重启资源管理器或重新登录才能生效
- **WPS清理功能强烈建议以管理员身份运行**
- 清理前请确保已备份重要文档
- 建议在企业环境中统一部署和配置
- 所有脚本文件使用UTF-8 with BOM编码以正确显示中文

## 🔐 安全性

- 所有操作都有确认提示
- 显示操作前后的状态对比
- WPS检测包含安全警告
- 遗留文件清理提供两种模式（保护用户数据）
- 自动检测管理员权限，提示权限不足

## 🎯 适用场景

### 个人用户
- 隐私保护：清理最近使用记录
- 界面优化：精简文件资源管理器
- 云服务管理：控制OneDrive同步

### 企业IT管理
- **安全审计**：检测不合规软件（如WPS Office）
- **批量部署**：统一配置用户电脑
- **数据保护**：防止数据通过云盘泄露
- **合规管理**：清理不安全的第三方软件

## 🌟 特色功能

### WPS Office企业安全管理
PowerClean专为企业环境设计了WPS Office管理功能：

1. **自动检测**：启动时显示安全警告
2. **风险提示**：明确标注数据泄露风险
3. **一键卸载**：调用官方卸载程序
4. **深度清理**：
   - 扫描云文档位置
   - 统计文件数量和大小
   - 提供保护性清理选项
5. **权限管理**：智能提示管理员权限需求

## 🔧 故障排除

### 权限问题
**问题**：清理WPS文件时提示"访问被拒绝"  
**解决**：右键PowerShell → 以管理员身份运行

### 更改未生效
**问题**：修改设置后未看到效果  
**解决**：使用程序内的"重启资源管理器"功能，或重新登录

### WPS仍然可以登录
**问题**：禁用选项无效  
**解决**：WPS个人版可能不完全支持这些设置，建议直接卸载

## 📚 相关文档

- [使用指南.md](使用指南.md) - 详细的操作指南
- [更新说明_v1.1.0.md](更新说明_v1.1.0.md) - 版本更新记录

## 🤝 贡献

欢迎提交问题和拉取请求！

## 📄 许可证

MIT License

## 📋 版本历史

- **v1.2.0** (2025-12-11)
  - ✨ 新增WPS Office安全管理功能
  - ✨ 新增WPS遗留文件扫描和清理
  - ✨ 新增管理员权限检测
  - ✨ 改进权限错误处理
  - 📝 更新文档和示例

- **v1.1.0** (2025-12-10)
  - ✨ 新增OneDrive同步管理
  - ✨ 新增Office最近文档管理
  - ✨ 新增快速访问管理
  - 🎨 优化交互式菜单界面
  - 📝 完善中文支持

- **v1.0.0** (2025-12-09)
  - 🎉 初始版本
  - ✅ 支持最近项目显示管理
  - ✅ 完整的中文支持
