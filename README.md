# SpotLight

**English** | [简体中文](README.zh-CN.md)

SpotLight is a lightweight VSTO add-in for Microsoft Excel that adds WPS-style active row and column highlighting with an on/off toggle and customizable colors.

## Features

- Highlights the row and column of the active cell automatically.
- Enables or disables highlighting from Excel's **View** tab.
- Includes an Office-style color palette and a custom color option.
- Saves the enabled state and selected color between Excel sessions.
- Temporarily removes its formatting rule while saving so the highlight is not stored in the workbook.

## Usage

1. Open the **View** tab in Excel.
2. In the **选区高亮** group, turn **高亮行列** on or off.
3. Select a highlight color from **颜色**.

## Installation

1. Close all Excel windows.
2. Double-click `SpotLight.vsto` in the build output directory.
3. Complete the installation prompts, then reopen Excel.

## Build Requirements

- Windows
- Microsoft Excel
- Visual Studio with Microsoft Office Developer Tools
- .NET Framework 4.7.2
