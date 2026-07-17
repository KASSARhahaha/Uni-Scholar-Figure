# Uni-Scholar Figure — PowerPoint 插件 (Windows + Mac)

跨平台 UniScholarFigure 1.2 行为复刻版。Windows 装完直接出 ribbon，Mac 多一步手动启用。

## 版本
- v1.2.4
- Release: 2026-07-17

## 平台支持矩阵

| 功能 | Windows | Mac |
|------|---------|-----|
| Trim PNG | ✅ GDI+ 内置 | ⚠️ 需 Python3 + Pillow（installer 会检测） |
| Generate Table | ✅ | ✅ |
| Layer Stack | ✅ | ✅ |
| Matrix Offset | ✅ | ✅ |
| Table Images | ✅ | ✅ |
| Add Icon | ✅ | ✅ |
| About | ✅ | ✅ |
| Check Updates | ✅ MSXML2.XMLHTTP | ✅ curl via MacScript |

## 预览

> 📸 截图采集进行中。如果你也装好了，欢迎按 `screenshots/CAPTURE-GUIDE.md` 跑一遍贡献回来。

### 1. Windows 一键安装

![Windows install.ps1 三次绿色 OK](screenshots/win-01-install-ps1.png)

### 2. PowerPoint 中的 Ribbon 标签

![Uni-Scholar Figure ribbon with 9 buttons](screenshots/win-02-ribbon-tab.png)

### 3. Generate Demo 一键演示

![Generate Demo 生成的 3 张幻灯片](screenshots/win-04-generate-demo.png)

### 4. Layer Stack 架构图

![5 层堆叠 + 箭头](screenshots/win-06-layer-stack.png)

### 5. Check for Updates

![Check for Updates 对话框](screenshots/win-09-check-updates.png)

---



## 安装 — Windows

1. **解压** 本目录到任意位置
2. **启用 VBA 信任访问**（一次性）：
   PowerPoint → 文件 → 选项 → 信任中心 → 信任中心设置
   → 宏设置 → ☑ 信任对 VBA 工程对象模型的访问
3. 关闭所有 PowerPoint
4. 右键 `install.ps1` → 用 PowerShell 运行
   - 如果 PowerShell 执行策略阻止：
     ```powershell
     Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
     ```
5. 看到三次绿色 `[OK]` 即成功
6. 打开 PowerPoint → 顶部出现 "Uni-Scholar Figure" 选项卡

## 安装 — Mac

### 第 0 步（可选）：装 Python3 让 PNG Trim 在 Mac 上能用

```bash
# 装 Homebrew（如果还没装）
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# 装 Python + Pillow
brew install python
pip3 install Pillow
```

不装也能用，只是 PNG Trim 按钮会提示装一下。其他 6 个功能不受影响。

### 第 1 步：跑 installer

双击 `install-mac.command`（或在终端 `bash install-mac.command`）。

它会：
- 检测 Python3 + Pillow
- 把 `UniScholarFigure.bas` 拷到 `~/Library/Group Containers/UBF8T346G9.Office/UserContent/Add-Ins/`
- 在 Finder 打开这个文件夹方便后续操作

### 第 2 步：在 PowerPoint 里启用（一次性手动）

**情况 A — 你只有 Mac，没有 Windows**：
1. 打开 PowerPoint → 新建空白文稿
2. 文件 → 另存为 → PowerPoint Macro-Enabled 演示文稿（.pptm）
3. 按 `Option+F11` 打开 VBA 编辑器
4. 文件 → 导入文件... → 选刚才 installer 拷过去的 `UniScholarFigure.bas`
5. 关闭 VBA 编辑器，保存 .pptm
6. 以后要用 UniScholarFigure 功能就打开这个 .pptm
   - 注意：.pptm 模式下 ribbon 只在文档开着时显示
   - 想永久 ribbon：需要用 Windows 跑 install.ps1 生成 .ppam，把 .ppam 拷回 Mac，再跑一次 install-mac.command

**情况 B — 你已经在 Windows 装过了，把 UniScholarFigure.ppam 拷到 Mac**：
1. 在 Windows：%APPDATA%\Microsoft\AddIns\UniScholarFigure.ppam
2. 把它拷到这个 dist 目录（与 install-mac.command 同级）
3. 重新跑 install-mac.command → 它会自动把 .ppam 装到 Mac AddIns 目录
4. 打开 PowerPoint → 工具 → PowerPoint 加载项 → 添加 → 选 UniScholarFigure.ppam
5. 重启 PowerPoint，"Uni-Scholar Figure" 选项卡永久出现

## 卸载

- Windows：右键 `uninstall.ps1` → 用 PowerShell 运行
- Mac：删除 `~/Library/Group Containers/UBF8T346G9.Office/UserContent/Add-Ins/UniScholarFigure.*` 即可

## 系统要求

| 平台 | 要求 |
|------|------|
| Windows | Win 10/11 + PowerPoint 2016/2019/2021/365/2024 + .NET Framework 4.5+ |
| Mac | macOS 12+ + PowerPoint 2019/2021/365 + （可选）Python 3.9+ 和 Pillow |

## 文件清单

```
dist/
├── UniScholarFigure.bas            # VBA 主模块（跨平台，含 #If Mac Then 分支）
├── customUI14.xml         # ribbon 定义
├── install.ps1            # Windows 一键安装器
├── install-mac.command    # Mac 一键安装器
├── uninstall.ps1          # Windows 卸载器
├── README.md              # 本文件
└── QUICK-TEST.md          # 内部测试说明
```

## 常见问题

**Q: 运行 install.ps1 提示 "PowerPoint blocks VBA project access"**
A: 没启用 VBA 信任访问。见 Windows 安装第 2 步。

**Q: Mac 上 PNG Trim 报 "Setup Required"**
A: 装 Python3 + Pillow。一行命令：`pip3 install Pillow`（前提是已 `brew install python`）

**Q: 安装成功但 ribbon 没出现**
A: 完全关闭 PowerPoint（任务管理器查 POWERPNT.EXE / Mac 活动监视器查 Microsoft PowerPoint），重开。

**Q: 与原版 UniScholarFigure 1.2 是否冲突**
A: 不冲突。ribbon 标签、注册表项、AddIns 文件名都不同。

**Q: 杀软报毒 / Mac Gatekeeper 拦截**
A: 误报。脚本明文可读可审计。Mac 首次运行 `.command` 文件可能要右键 → 打开，绕过 Gatekeeper。

## 与原版的差异

- 纯 VBA 实现（原版 VSTO/.NET）
- AI 生图功能不在 1.2 release notes 里，未实现
- 兑换码 / 机器码激活流程未实现（开源，无需激活）
- 同一行为可能 UI 微调（input box 而非任务面板），但功能等价

## 反馈

源码：`/home/yangkai/00-make-money/Uni-Scholar-Figure/`
Python CLI 同源实现：`unisfig --help`（同目录 .venv 已配置）
