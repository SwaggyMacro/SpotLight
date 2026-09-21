# SpotLight

[English](README.md) | **简体中文**

SpotLight 是一款为 Microsoft Excel 提供 WPS 风格选中单元格行列高亮、可开关并支持自定义颜色的轻量级 VSTO 加载项。

## 功能

- 自动高亮当前选中单元格所在的行和列。
- 在 Excel 的“视图”选项卡中开启或关闭高亮。
## 使用方法

1. 打开 Excel 的“视图”选项卡。
2. 在“选区高亮”组中勾选或取消勾选“高亮行列”。
3. 点击“颜色”选择高亮颜色。

## 安装

1. 关闭所有 Excel 窗口。
2. 运行单文件安装程序 `dist\SpotLightSetup.exe`。
3. 按照安装提示完成安装，然后重新打开 Excel。

安装程序会先卸载当前用户已经安装的 SpotLight VSTO 版本，因此即使安装包来自不同目录，也可以正常更新。

如果直接使用 Visual Studio 的 Publish 输出安装，必须让 `SpotLight.vsto` 与完整的 `Application Files` 目录保持在一起。`.vsto` 只是部署清单，不能单独分发。

## 构建要求

- Windows
- Microsoft Excel
- Visual Studio，且已安装 Microsoft Office 开发人员工具
- .NET Framework 4.7.2

## 构建单文件安装程序

1. 在 Visual Studio 中发布 VSTO 项目。
2. 运行：

   ```powershell
   .\Installer\build-setup.ps1 -PublishDirectory .\bin\Release\app.publish
   ```

生成的安装程序位于 `dist\SpotLightSetup.exe`。

## 自动发布

`SpotLight.csproj` 是版本号的唯一来源：

```xml
<Version>1.0.0.1</Version>
```

程序集版本、文件版本、VSTO 部署版本、安装程序文件名和 GitHub Release 版本都由该值生成。

需要先在 GitHub 仓库中配置以下 Secrets：

- `VSTO_PFX_BASE64`：用于签署 VSTO 清单和安装程序的代码签名 PFX 文件的 Base64 内容。
- `VSTO_PFX_PASSWORD`：PFX 密码；如果 PFX 没有密码，可以不创建该 Secret。

可以通过 PowerShell 生成 Base64 内容：

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes('.\SpotLight_TemporaryKey.pfx')) | Set-Clipboard
```

发布时修改 `<Version>`、提交更改，然后推送与其一致的四段版本 Tag：

```powershell
git tag v1.0.0.1
git push origin v1.0.0.1
```

工作流会自动创建 GitHub Release，并上传单文件安装程序、完整 VSTO Publish 压缩包和 SHA-256 校验文件。
