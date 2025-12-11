# PowerClean - PowerShell系统管理工具库

一个用于管理Windows系统界面和功能的PowerShell模块，专为中文用户设计。

## 功能特性

### 最近项目显示管理
- 控制开始菜单和任务栏右键菜单中"最近"列表的显示
- 支持开启/关闭最近使用的项目显示
- 查询当前状态
- 可选择是否自动重启资源管理器以立即应用更改

## 快速开始

### 方法一：使用交互式菜单（推荐）⭐

**最简单的方式 - 双击启动：**
```
双击 "启动PowerClean.bat" 文件
```

**或在PowerShell中运行：**
```powershell
.\PowerClean.ps1
```

交互式菜单提供：
- 🔍 查看当前状态
- ✅ 启用最近项目显示
- ❌ 禁用最近项目显示
- 🔄 重启资源管理器
- 所有操作都有确认提示，安全便捷

详细使用说明请查看 [使用指南.md](使用指南.md)

### 方法二：直接使用命令

1. 导入模块：
```powershell
Import-Module .\PowerClean.psd1
```

2. 使用命令（见下方详细说明）

## 使用方法

### 禁用"最近"项目显示
```powershell
Disable-RecentItems
```

### 启用"最近"项目显示
```powershell
Enable-RecentItems
```

### 查询当前状态
```powershell
Get-RecentItemsStatus
```

### 手动设置状态
```powershell
# 禁用
Set-RecentItemsDisplay -Enabled $false

# 启用
Set-RecentItemsDisplay -Enabled $true

# 启用并立即重启资源管理器
Set-RecentItemsDisplay -Enabled $true -RestartExplorer
```

### 重启资源管理器
```powershell
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

## 项目结构

```
power_clean/
├── PowerClean.ps1            # 主程序入口（交互式菜单）
├── PowerClean.psm1           # 主模块文件
├── PowerClean.psd1           # 模块清单
├── README.md                 # 说明文档
├── Functions/                # 功能函数目录
│   └── RecentItems.ps1      # 最近项目管理功能
└── Examples/                 # 示例脚本目录
    ├── Quick-Disable.ps1    # 快速禁用脚本
    ├── Quick-Enable.ps1     # 快速启用脚本
    └── Example-RecentItems.ps1  # 完整示例
```

## 系统要求

- Windows 10/11
- PowerShell 5.1 或更高版本
- 某些功能可能需要管理员权限

## 工作原理

本模块通过修改Windows注册表来控制系统行为：

- **Start_TrackDocs**: 控制是否在开始菜单中显示最近打开的文档
- **Start_TrackProgs**: 控制是否在开始菜单中跟踪最近使用的程序

注册表路径: `HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced`

## 注意事项

- 修改注册表设置后，需要重启资源管理器或重新登录才能看到效果
- 建议在修改前备份相关注册表设置
- 所有脚本文件使用UTF-8编码以正确显示中文

## 未来计划

- 添加更多系统界面管理功能
- 支持批量配置导入/导出
- 提供图形化界面选项
- 添加系统清理功能

## 许可证

MIT License

## 贡献

欢迎提交问题和拉取请求！

## 版本历史

- **v1.0.0** (2025-12-09)
  - 初始版本
  - 支持最近项目显示管理
  - 完整的中文支持
