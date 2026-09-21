using System;
using System.Drawing;
using System.Globalization;
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
            new ColorChoice(0xFFFFFF, "白色", "White"),
            new ColorChoice(0x000000, "黑色", "Black"),
            new ColorChoice(0xE7E6E6, "浅灰色", "Light Gray"),
            new ColorChoice(0x44546A, "蓝灰色", "Blue Gray"),
            new ColorChoice(0x5B9BD5, "蓝色", "Blue"),
            new ColorChoice(0xED7D31, "橙色", "Orange"),
            new ColorChoice(0xA5A5A5, "灰色", "Gray"),
            new ColorChoice(0xFFC000, "金色", "Gold"),
            new ColorChoice(0x4472C4, "深蓝色", "Dark Blue"),
            new ColorChoice(0x70AD47, "绿色", "Green"),

            new ColorChoice(0xF2F2F2, "白色，深色 5%", "White, Darker 5%"),
            new ColorChoice(0xD9D9D9, "黑色，浅色 85%", "Black, Lighter 85%"),
            new ColorChoice(0xF4F3F3, "浅灰色，浅色 80%", "Light Gray, Lighter 80%"),
            new ColorChoice(0xD6DCE4, "蓝灰色，浅色 80%", "Blue Gray, Lighter 80%"),
            new ColorChoice(0xDEEBF7, "蓝色，浅色 80%", "Blue, Lighter 80%"),
            new ColorChoice(0xFCE4D6, "橙色，浅色 80%", "Orange, Lighter 80%"),
            new ColorChoice(0xEDEDED, "灰色，浅色 80%", "Gray, Lighter 80%"),
            new ColorChoice(0xFFF2CC, "金色，浅色 80%", "Gold, Lighter 80%"),
            new ColorChoice(0xD9E2F3, "深蓝色，浅色 80%", "Dark Blue, Lighter 80%"),
            new ColorChoice(0xE2F0D9, "绿色，浅色 80%", "Green, Lighter 80%"),

            new ColorChoice(0xD9D9D9, "白色，深色 15%", "White, Darker 15%"),
            new ColorChoice(0xBFBFBF, "黑色，浅色 75%", "Black, Lighter 75%"),
            new ColorChoice(0xD0CECE, "浅灰色，深色 10%", "Light Gray, Darker 10%"),
            new ColorChoice(0xACB9CA, "蓝灰色，浅色 60%", "Blue Gray, Lighter 60%"),
            new ColorChoice(0xBDD7EE, "蓝色，浅色 60%", "Blue, Lighter 60%"),
            new ColorChoice(0xF8CBAD, "橙色，浅色 60%", "Orange, Lighter 60%"),
            new ColorChoice(0xDBDBDB, "灰色，浅色 60%", "Gray, Lighter 60%"),
            new ColorChoice(0xFFE699, "金色，浅色 60%", "Gold, Lighter 60%"),
            new ColorChoice(0xB4C6E7, "深蓝色，浅色 60%", "Dark Blue, Lighter 60%"),
            new ColorChoice(0xC6E0B4, "绿色，浅色 60%", "Green, Lighter 60%"),

            new ColorChoice(0xBFBFBF, "白色，深色 25%", "White, Darker 25%"),
            new ColorChoice(0x808080, "黑色，浅色 50%", "Black, Lighter 50%"),
            new ColorChoice(0xAEAAAA, "浅灰色，深色 25%", "Light Gray, Darker 25%"),
            new ColorChoice(0x8497B0, "蓝灰色，浅色 40%", "Blue Gray, Lighter 40%"),
            new ColorChoice(0x9DC3E6, "蓝色，浅色 40%", "Blue, Lighter 40%"),
            new ColorChoice(0xF4B183, "橙色，浅色 40%", "Orange, Lighter 40%"),
            new ColorChoice(0xC9C9C9, "灰色，浅色 40%", "Gray, Lighter 40%"),
            new ColorChoice(0xFFD966, "金色，浅色 40%", "Gold, Lighter 40%"),
            new ColorChoice(0x8EAADB, "深蓝色，浅色 40%", "Dark Blue, Lighter 40%"),
            new ColorChoice(0xA9D18E, "绿色，浅色 40%", "Green, Lighter 40%"),

            new ColorChoice(0xC00000, "深红色", "Dark Red"),
            new ColorChoice(0xFF0000, "红色", "Red"),
            new ColorChoice(0xFF9900, "深橙色", "Dark Orange"),
            new ColorChoice(0xFFFF00, "黄色", "Yellow"),
            new ColorChoice(0x92D050, "浅绿色", "Light Green"),
            new ColorChoice(0x00B050, "绿色", "Green"),
            new ColorChoice(0x00B0F0, "浅蓝色", "Light Blue"),
            new ColorChoice(0x0070C0, "蓝色", "Blue"),
            new ColorChoice(0x002060, "深蓝色", "Dark Blue"),
            new ColorChoice(0x7030A0, "紫色", "Purple")
        };

        private readonly object[] paletteImages = new object[Palette.Length];
        private int currentColorImageArgb;
        private object currentColorImage;
        private Office.IRibbonUI ribbonUi;
        private bool useChineseUi;

        public string GetCustomUI(string ribbonId)
        {
            useChineseUi = IsChineseOfficeUi();
            return BuildCustomUi(useChineseUi);
        }

        private static string BuildCustomUi(bool useChinese)
        {
            string groupLabel = useChinese ? "选区高亮" : "Selection Highlight";
            string enabledLabel = useChinese ? "高亮行列" : "Highlight Row and Column";
            string colorLabel = useChinese ? "颜色" : "Color";
            string moreColorsLabel = useChinese ? "其他颜色..." : "More Colors...";

            return string.Format(
                CultureInfo.InvariantCulture,
                @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""OnLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""TabView"">
        <group id=""SpotLightSelectionHighlightGroup"" label=""{0}"">
          <checkBox id=""HighlightEnabled"" label=""{1}"" getPressed=""GetHighlightEnabled"" onAction=""OnHighlightEnabledChanged"" />
          <menu id=""HighlightColorMenu"" label=""{2}"" size=""large"" getImage=""GetCurrentColorImage"">
            <gallery id=""HighlightColorGallery"" label=""{2}"" columns=""10"" itemWidth=""18"" itemHeight=""18"" showItemLabel=""false"" showItemImage=""true"" getItemCount=""GetHighlightColorCount"" getItemLabel=""GetHighlightColorLabel"" getItemImage=""GetHighlightColorImage"" getSelectedItemIndex=""GetHighlightColorIndex"" onAction=""OnHighlightColorChanged"" />
            <menuSeparator id=""HighlightColorSeparator"" />
            <button id=""HighlightMoreColors"" label=""{3}"" onAction=""OnMoreColors"" />
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>",
                groupLabel,
                enabledLabel,
                colorLabel,
                moreColorsLabel);
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
            return IsPaletteIndex(index) ? Palette[index].GetLabel(useChineseUi) : string.Empty;
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

        private static bool IsChineseOfficeUi()
        {
            Office.LanguageSettings languageSettings = null;
            try
            {
                languageSettings = Globals.ThisAddIn.Application.LanguageSettings;
                int languageId = languageSettings.get_LanguageID(
                    Office.MsoAppLanguageID.msoLanguageIDUI);

                return IsChineseLanguageId(languageId);
            }
            catch (COMException)
            {
                return false;
            }
            catch (InvalidComObjectException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
            finally
            {
                if (languageSettings != null && Marshal.IsComObject(languageSettings))
                {
                    Marshal.ReleaseComObject(languageSettings);
                }
            }
        }

        private static bool IsChineseLanguageId(int languageId)
        {
            // All Chinese LCIDs share primary language ID 0x04.
            return (languageId & 0x03FF) == 0x04;
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
            internal ColorChoice(int rgb, string chineseLabel, string englishLabel)
            {
                Color = Color.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
                ChineseLabel = chineseLabel;
                EnglishLabel = englishLabel;
            }

            internal Color Color { get; private set; }

            private string ChineseLabel { get; set; }

            private string EnglishLabel { get; set; }

            internal string GetLabel(bool useChinese)
            {
                return useChinese ? ChineseLabel : EnglishLabel;
            }
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
