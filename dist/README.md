# DakeSCI Clone — PowerPoint 插件 (Windows)

一键安装的 DakeSCI 1.2 行为复刻版。打开 PowerPoint 后顶部会出现 "DakeSCI Clone" 选项卡，包含 7 个按钮，分别对应原版 DakeSCI 1.2 列出的 7 项功能。

## 版本
- v1.2.0
- Release: 2026-07-03（与原版 1.2 同日）

## 功能对照（与原版 DakeSCI 1.2）

| 按钮 | 功能 | 对应原版 |
|------|------|---------|
| Trim PNG | 自动裁除 PNG 空白边 | ✅ |
| Generate Table | CSV / Markdown / 纯文本 → PPT 表格 | ✅ |
| Layer Stack | 层状结构（带下箭头） | ✅ |
| Matrix Offset | 矩阵逐行偏移（alternate / progressive / none） | ✅ |
| Table Images | 多图按比例排入表格（letterbox，不拉伸） | ✅ |
| Add Icon | 插入命名图标 | ✅ |
| About | 显示版本 + 发布日期 | ✅ |

## 安装步骤（一次性）

1. **解压** 本目录到任意位置（如桌面）
2. **启用 VBA 信任访问**（关键，只需做一次）：
   - PowerPoint → 文件 → 选项 → 信任中心 → 信任中心设置
   - 宏设置 → ☑ 信任对 VBA 工程对象模型的访问
3. **关闭 PowerPoint**（如果开着）
4. **右键 `install.ps1` → 用 PowerShell 运行**
   - 如果 PowerShell 执行策略阻止了脚本，先在管理员 PowerShell 跑一次：
     ```powershell
     Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
     ```
5. 看到绿色 `[OK]` 提示即安装成功
6. **打开 PowerPoint** —— 顶部应该出现 "DakeSCI Clone" 选项卡

## 卸载

右键 `uninstall.ps1` → 用 PowerShell 运行。

## 系统要求

- Windows 10 / 11
- PowerPoint 2016 / 2019 / 2021 / 365（Office 2024 也可以）
- .NET Framework 4.5+（PowerShell 自带）

## 文件清单

```
dist/
├── DakeSCI.bas         # VBA 主模块（7 个功能）
├── customUI14.xml      # ribbon 定义
├── install.ps1         # Windows 安装器
├── uninstall.ps1       # 卸载器
└── README.md           # 本文件
```

## 常见问题

**Q: 运行 install.ps1 提示 "PowerPoint blocks VBA project access"**
A: 没启用第 2 步的 VBA 信任访问。回到 PowerPoint 启用后重新跑 `install.ps1`。

**Q: 安装成功但 ribbon 没出现**
A: 完全关闭所有 PowerPoint 进程（任务管理器查 `POWERPNT.EXE`），然后重开。

**Q: 杀软报毒**
A: 误报。脚本是明文，可读可审计。VBA 注入 PowerPoint 是标准 COM 自动化操作。

**Q: 与原版 DakeSCI 1.2 是否冲突**
A: 不冲突。本插件名为 `DakeSCI Clone`，ribbon 标签、注册表项、AddIns 文件名（`DakeSCI.ppam` vs 原 `DakeSCI1.2.vsto`）都不同。

## 与原版的差异

- 本插件用纯 VBA 实现（原版是 VSTO/.NET）
- AI 生图功能不在 1.2 release notes 里，未实现
- 兑换码 / 机器码激活流程未实现（开源，无需激活）
- 同一行为可能 UI 微调（input box 而非任务面板），但功能等价

## 反馈

源码仓库：`/home/yangkai/00-make-money/dakesci-clone/`
Python CLI 版本：`dake --help`（同目录的 .venv 已配置）
