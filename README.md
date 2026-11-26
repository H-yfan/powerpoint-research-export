# 科研图片高清导出 · PowerPoint Research Export

Windows 桌面 Microsoft PowerPoint 加载项，中文功能区，安装一次后供不同演示文稿使用。当前版本 **v1.2.0**，未签名。

## 功能

- 当前幻灯片、选中对象或组合导出 PNG。
- 300 / 600 / 900 / 1200 DPI，自定义 36–2400 DPI。
- 按实际物理尺寸计算像素，使用 `Slide.Export` 显式像素尺寸，不依赖 `ExportBitmapResolution` 注册表。
- 写入 PNG pHYs 分辨率元数据，检查实际像素与 PNG CRC。
- **定尺寸裁剪框**：输入毫米宽高，移动框后导出区域内全部可见内容，无须拆开组合。
- 透明选区为实验功能，尺寸或 alpha 校验失败时拒绝保存。

55 × 50 mm、600 DPI 对应 **1299 × 1181 像素**；pHYs 为23622 px/m（约599.9988 DPI）。高DPI不会恢复低分辨率原图的细节。

## 安装

1. 下载完整项目并解压。需要桌面 PowerPoint、VBA 组件、Windows PowerShell 5.1。
2. 在解压文件夹打开 PowerShell，执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Addin.ps1
```

3. 保存并完全退出 PowerPoint，再重新打开，查看 **科研高清导出** 选项卡。

脚本将PPAM和PNG处理组件安装到 `%APPDATA%\ResearchExport`，并启用自动加载。普通安装不需要开启“信任对VBA工程对象模型的访问”。未签名宏可能受组织政策限制；不要全局启用所有宏。进程级 ExecutionPolicy Bypass 不改变系统策略，组织策略限制时请联系管理员。

## 使用裁剪框

1. 点击 **插入裁剪框…**，输入宽高（毫米）。
2. 移动红色虚线框罩住所需区域；不要将它与组图组合或旋转。
3. 单独选中框，选择DPI，点击 **按裁剪框导出**。

新功能输出框内所有可见内容和幻灯片背景，自动排除所有顶层科研裁剪框。框必须完整位于幻灯片内。裁剪框保留在原稿中，用完可删除；旧的当前页/选区导出仍会把它当作普通形状。

## 修改与重新构建

完整源码位于 `ResearchExport.bas`（**GBK编码**，构建脚本按GBK读取），功能区在 `customUI14.xml`。不要直接把BAS改成UTF-8而不更新构建器。

备份并移走 `build/ResearchExport.ppam`，在PowerPoint中卸载同名旧加载项。临时开启“信任对VBA工程对象模型的访问”，保存设置后运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-Addin.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Addin.ps1
```

完成后关闭VBA工程访问。详细安装、安全和卸载步骤见 [使用说明](docs/使用说明.txt)。

## 实测与限制

在 Windows 桌面 PowerPoint 16.0 实际生成PPAM、编译运行VBA；当前页、选区、带偏移的裁剪框输出均核对像素与pHYs。嵌套组合测试检查了布局及取样像素。自动化测试跳过输入/保存对话框；没有完整的交互点击测试、跨版本兼容性保证或数字签名。

透明导出、复杂字体、阴影、旋转形状、OLE仍需检查。白底选区和裁剪框先渲染临时页面再按整数像素裁切，不做插值；大页面高DPI受32767边长、一亿总像素上限限制。当前版本不提供TIFF/SVG。

详见 [验证报告](docs/验证报告.txt)。

## 项目文件

| 文件 | 用途 |
|---|---|
| `build/ResearchExport.ppam` | 本机生成的可安装加载项，未签名 |
| `Install-Addin.ps1` | 固定路径安装并启用自动加载 |
| `Build-Addin.ps1` | 从VBA源码生成PPAM并注入Ribbon |
| `ResearchExport.bas` | 完整VBA，GBK |
| `customUI14.xml` | 中文功能区 |
| `Finalize-Png.ps1` | PNG裁切、尺寸/CRC校验、DPI写入 |

本仓库仅包含工具代码与通用文档，不含科研原稿、论文图片或临时测试演示文稿。
