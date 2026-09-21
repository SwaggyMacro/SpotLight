using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace SpotLight
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class SpotLightRibbon : Office.IRibbonExtensibility
    {
        private static readonly ColorChoice[] Palette =
        {
            new ColorChoice(0xFFFFFF, "白色"),
            new ColorChoice(0x000000, "黑色"),
            new ColorChoice(0xE7E6E6, "浅灰色"),
            new ColorChoice(0x44546A, "蓝灰色"),
            new ColorChoice(0x5B9BD5, "蓝色"),
            new ColorChoice(0xED7D31, "橙色"),
            new ColorChoice(0xA5A5A5, "灰色"),
            new ColorChoice(0xFFC000, "金色"),
            new ColorChoice(0x4472C4, "深蓝色"),
            new ColorChoice(0x70AD47, "绿色"),

            new ColorChoice(0xF2F2F2, "白色，深色 5%"),
            new ColorChoice(0xD9D9D9, "黑色，浅色 85%"),
            new ColorChoice(0xF4F3F3, "浅灰色，浅色 80%"),
            new ColorChoice(0xD6DCE4, "蓝灰色，浅色 80%"),
            new ColorChoice(0xDEEBF7, "蓝色，浅色 80%"),
            new ColorChoice(0xFCE4D6, "橙色，浅色 80%"),
            new ColorChoice(0xEDEDED, "灰色，浅色 80%"),
            new ColorChoice(0xFFF2CC, "金色，浅色 80%"),
            new ColorChoice(0xD9E2F3, "深蓝色，浅色 80%"),
            new ColorChoice(0xE2F0D9, "绿色，浅色 80%"),

            new ColorChoice(0xD9D9D9, "白色，深色 15%"),
            new ColorChoice(0xBFBFBF, "黑色，浅色 75%"),
            new ColorChoice(0xD0CECE, "浅灰色，深色 10%"),
            new ColorChoice(0xACB9CA, "蓝灰色，浅色 60%"),
            new ColorChoice(0xBDD7EE, "蓝色，浅色 60%"),
            new ColorChoice(0xF8CBAD, "橙色，浅色 60%"),
            new ColorChoice(0xDBDBDB, "灰色，浅色 60%"),
            new ColorChoice(0xFFE699, "金色，浅色 60%"),
            new ColorChoice(0xB4C6E7, "深蓝色，浅色 60%"),
            new ColorChoice(0xC6E0B4, "绿色，浅色 60%"),

            new ColorChoice(0xBFBFBF, "白色，深色 25%"),
            new ColorChoice(0x808080, "黑色，浅色 50%"),
            new ColorChoice(0xAEAAAA, "浅灰色，深色 25%"),
            new ColorChoice(0x8497B0, "蓝灰色，浅色 40%"),
            new ColorChoice(0x9DC3E6, "蓝色，浅色 40%"),
            new ColorChoice(0xF4B183, "橙色，浅色 40%"),
            new ColorChoice(0xC9C9C9, "灰色，浅色 40%"),
            new ColorChoice(0xFFD966, "金色，浅色 40%"),
            new ColorChoice(0x8EAADB, "深蓝色，浅色 40%"),
            new ColorChoice(0xA9D18E, "绿色，浅色 40%"),

            new ColorChoice(0xC00000, "深红色"),
            new ColorChoice(0xFF0000, "红色"),
            new ColorChoice(0xFF9900, "深橙色"),
            new ColorChoice(0xFFFF00, "黄色"),
            new ColorChoice(0x92D050, "浅绿色"),
            new ColorChoice(0x00B050, "绿色"),
            new ColorChoice(0x00B0F0, "浅蓝色"),
            new ColorChoice(0x0070C0, "蓝色"),
            new ColorChoice(0x002060, "深蓝色"),
            new ColorChoice(0x7030A0, "紫色")
        };

        private readonly object[] paletteImages = new object[Palette.Length];
        private int currentColorImageArgb;
        private object currentColorImage;
        private Office.IRibbonUI ribbonUi;

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""OnLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""TabView"">
        <group id=""SpotLightSelectionHighlightGroup"" label=""选区高亮"">
          <checkBox id=""HighlightEnabled"" label=""高亮行列"" getPressed=""GetHighlightEnabled"" onAction=""OnHighlightEnabledChanged"" />
          <menu id=""HighlightColorMenu"" label=""颜色"" size=""large"" getImage=""GetCurrentColorImage"">
            <gallery id=""HighlightColorGallery"" label=""颜色"" columns=""10"" itemWidth=""18"" itemHeight=""18"" showItemLabel=""false"" showItemImage=""true"" getItemCount=""GetHighlightColorCount"" getItemLabel=""GetHighlightColorLabel"" getItemImage=""GetHighlightColorImage"" getSelectedItemIndex=""GetHighlightColorIndex"" onAction=""OnHighlightColorChanged"" />
            <menuSeparator id=""HighlightColorSeparator"" />
            <button id=""HighlightMoreColors"" label=""其他颜色..."" onAction=""OnMoreColors"" />
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnLoad(Office.IRibbonUI ui)
        {
            ribbonUi = ui;
        }

        public bool GetHighlightEnabled(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.HighlightEnabled;
        }

        public int GetHighlightColorCount(Office.IRibbonControl control)
        {
            return Palette.Length;
        }

        public string GetHighlightColorLabel(Office.IRibbonControl control, int index)
        {
            return IsPaletteIndex(index) ? Palette[index].Label : string.Empty;
        }

        public object GetHighlightColorImage(Office.IRibbonControl control, int index)
        {
            if (!IsPaletteIndex(index))
            {
                return null;
            }

            if (paletteImages[index] == null)
            {
                paletteImages[index] = PictureDispConverter.ToPictureDisp(CreateColorSwatch(Palette[index].Color, 18));
            }

            return paletteImages[index];
        }

        public object GetCurrentColorImage(Office.IRibbonControl control)
        {
            int colorArgb = Globals.ThisAddIn.HighlightColorArgb;
            if (currentColorImage == null || currentColorImageArgb != colorArgb)
            {
                currentColorImageArgb = colorArgb;
                currentColorImage = PictureDispConverter.ToPictureDisp(
                    CreateColorSwatch(Color.FromArgb(colorArgb), 32));
            }

            return currentColorImage;
        }

        public int GetHighlightColorIndex(Office.IRibbonControl control)
        {
            int selectedArgb = Globals.ThisAddIn.HighlightColorArgb;
            for (int index = 0; index < Palette.Length; index++)
            {
                if (Palette[index].Color.ToArgb() == selectedArgb)
                {
                    return index;
                }
            }

            return -1;
        }

        public void OnHighlightEnabledChanged(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.SetHighlightEnabled(pressed);
            InvalidateControl("HighlightEnabled");
        }

        public void OnHighlightColorChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (IsPaletteIndex(selectedIndex))
            {
                Globals.ThisAddIn.SetHighlightColor(Palette[selectedIndex].Color.ToArgb());
                InvalidateColorControls();
            }
        }

        public void OnMoreColors(Office.IRibbonControl control)
        {
            using (ColorDialog dialog = new ColorDialog())
            {
                dialog.Color = Color.FromArgb(Globals.ThisAddIn.HighlightColorArgb);
                dialog.FullOpen = true;
                dialog.AnyColor = true;

                IWin32Window owner = new ExcelWindow(new IntPtr(Globals.ThisAddIn.Application.Hwnd));
                if (dialog.ShowDialog(owner) == DialogResult.OK)
                {
                    Globals.ThisAddIn.SetHighlightColor(dialog.Color.ToArgb());
                    InvalidateColorControls();
                }
            }
        }

        private static bool IsPaletteIndex(int index)
        {
            return index >= 0 && index < Palette.Length;
        }

        private static Bitmap CreateColorSwatch(Color color, int size)
        {
            Bitmap bitmap = new Bitmap(size, size);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (SolidBrush brush = new SolidBrush(color))
            using (Pen border = new Pen(Color.FromArgb(112, 112, 112)))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillRectangle(brush, 1, 1, size - 3, size - 3);
                graphics.DrawRectangle(border, 0, 0, size - 2, size - 2);
            }

            return bitmap;
        }

        private void InvalidateColorControls()
        {
            InvalidateControl("HighlightColorMenu");
            InvalidateControl("HighlightColorGallery");
        }

        private void InvalidateControl(string controlId)
        {
            if (ribbonUi != null)
            {
                ribbonUi.InvalidateControl(controlId);
            }
        }

        private sealed class ColorChoice
        {
            internal ColorChoice(int rgb, string label)
            {
                Color = Color.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
                Label = label;
            }

            internal Color Color { get; private set; }

            internal string Label { get; private set; }
        }

        private sealed class ExcelWindow : IWin32Window
        {
            internal ExcelWindow(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; private set; }
        }

        [ComVisible(false)]
        private sealed class PictureDispConverter : AxHost
        {
            private PictureDispConverter()
                : base(string.Empty)
            {
            }

            internal static object ToPictureDisp(Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }
    }
}
