# SpotLight

**English** | [简体中文](README.zh-CN.md)

SpotLight is a lightweight VSTO add-in for Microsoft Excel that adds WPS-style active row and column highlighting with an on/off toggle and customizable colors.

## Features

- Highlights the row and column of the active cell automatically.
- Enables or disables highlighting from Excel's **View** tab.

## Usage

1. Open the **View** tab in Excel.
2. In the **Selection Highlight** group, turn **Highlight Row and Column** on or off.
3. Select a highlight color from **Color**.

## Installation

1. Close all Excel windows.
2. Run the single-file installer `dist\SpotLightSetup.exe`.
3. Complete the installation prompts, then reopen Excel.

The installer removes an existing per-user SpotLight VSTO installation before installing the new version, so updates also work when the package is launched from a different folder.

To install directly from Visual Studio's Publish output instead, keep `SpotLight.vsto` together with the entire `Application Files` directory. A `.vsto` file is a deployment manifest and cannot be distributed by itself.

## Build Requirements

- Windows
- Microsoft Excel
- Visual Studio with Microsoft Office Developer Tools
- .NET Framework 4.7.2

## Build the Single-File Installer

1. Publish the VSTO project in Visual Studio.
2. Run:

   ```powershell
   .\Installer\build-setup.ps1 -PublishDirectory .\bin\Release\app.publish
   ```

The generated installer is written to `dist\SpotLightSetup.exe`.

## Automated Releases

`SpotLight.csproj` is the single source of truth for the version:

```xml
<Version>1.0.0.1</Version>
```

The assembly version, file version, VSTO deployment version, installer filename, and GitHub Release version are derived from this value.

Configure these GitHub repository secrets once:

- `VSTO_PFX_BASE64`: Base64-encoded code-signing PFX used to sign the VSTO manifests and installer.
- `VSTO_PFX_PASSWORD`: PFX password; leave this secret unset when the PFX has no password.

Generate the Base64 value in PowerShell:

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes('.\SpotLight_TemporaryKey.pfx')) | Set-Clipboard
```

To publish a release, update `<Version>`, commit the change, and push a matching four-part version tag:

```powershell
git tag v1.0.0.1
git push origin v1.0.0.1
```

The workflow creates the GitHub Release automatically and attaches the single-file installer, the complete VSTO Publish package, and SHA-256 checksums.
