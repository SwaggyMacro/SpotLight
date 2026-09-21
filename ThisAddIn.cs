using System;
using System.Configuration;
using System.Drawing;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SpotLight
{
    public partial class ThisAddIn
    {
        private const string HighlightFormulaMarker = "SpotLight.SelectionHighlight";
        private Excel.Worksheet highlightedWorksheet;
        private Excel.FormatCondition highlightCondition;
        private bool restoreHighlightAfterSave;
        private bool highlightEnabled = true;
        private Color highlightColor = Color.FromArgb(255, 246, 128);

        internal bool HighlightEnabled
        {
            get { return highlightEnabled; }
        }

        internal int HighlightColorArgb
        {
            get { return highlightColor.ToArgb(); }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            LoadUserSettings();

            Application.SheetSelectionChange += Application_SheetSelectionChange;
            Application.SheetActivate += Application_SheetActivate;
            Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            Application.WorkbookAfterSave += Application_WorkbookAfterSave;

            // Excel may not raise SheetSelectionChange for the initial selection.
            UpdateHighlightForActiveSheet();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            Application.SheetActivate -= Application_SheetActivate;
            Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
            Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
            restoreHighlightAfterSave = false;
            RemoveHighlight();
        }

        private void Application_SheetSelectionChange(object sheet, Excel.Range target)
        {
            UpdateHighlight(sheet as Excel.Worksheet, target);
        }

        private void Application_SheetActivate(object sheet)
        {
            UpdateHighlightForActiveSheet(sheet as Excel.Worksheet);
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook workbook, bool saveAsUi, ref bool cancel)
        {
            if (!IsHighlightedWorkbook(workbook))
            {
                return;
            }

            restoreHighlightAfterSave = true;
            RemoveHighlight();
        }

        private void Application_WorkbookAfterSave(Excel.Workbook workbook, bool success)
        {
            if (!restoreHighlightAfterSave)
            {
                return;
            }

            restoreHighlightAfterSave = false;
            if (SameComObject(Application.ActiveWorkbook, workbook))
            {
                UpdateHighlightForActiveSheet();
            }
        }

        internal void SetHighlightEnabled(bool enabled)
        {
            highlightEnabled = enabled;
            SaveUserSettings();
            if (enabled)
            {
                UpdateHighlightForActiveSheet();
            }
            else
            {
                restoreHighlightAfterSave = false;
                RemoveHighlight();
                Excel.Worksheet worksheet = Application.ActiveSheet as Excel.Worksheet;
                if (worksheet != null)
                {
                    try
                    {
                        RemoveStaleHighlightConditions(worksheet);
                    }
                    catch (COMException)
                    {
                        // A protected or closing worksheet may reject the cleanup.
                    }
                    finally
                    {
                        ReleaseComObject(worksheet);
                    }
                }
            }
        }

        internal void SetHighlightColor(int colorArgb)
        {
            highlightColor = Color.FromArgb(colorArgb);
            SaveUserSettings();
            if (highlightCondition == null || highlightedWorksheet == null)
            {
                return;
            }

            Excel.Workbook workbook = null;
            bool wasSaved = false;
            try
            {
                workbook = highlightedWorksheet.Parent as Excel.Workbook;
                wasSaved = workbook != null && workbook.Saved;
                highlightCondition.Interior.Color = ColorTranslator.ToOle(highlightColor);
            }
            catch (COMException)
            {
                RemoveHighlight();
            }
            finally
            {
                RestoreWorkbookSavedState(workbook, wasSaved);
                ReleaseComObject(workbook);
            }
        }

        private void LoadUserSettings()
        {
            try
            {
                Properties.Settings settings = Properties.Settings.Default;
                if (settings.UpgradeRequired)
                {
                    settings.Upgrade();
                    settings.UpgradeRequired = false;
                    settings.Save();
                }

                highlightEnabled = settings.HighlightEnabled;
                highlightColor = Color.FromArgb(settings.HighlightColorArgb);
            }
            catch (ConfigurationErrorsException)
            {
                // Keep the built-in defaults if the user configuration is corrupt.
            }
        }

        private void SaveUserSettings()
        {
            try
            {
                Properties.Settings settings = Properties.Settings.Default;
                settings.HighlightEnabled = highlightEnabled;
                settings.HighlightColorArgb = highlightColor.ToArgb();
                settings.Save();
            }
            catch (ConfigurationErrorsException)
            {
                // Highlighting should continue to work even if settings cannot be saved.
            }
            catch (UnauthorizedAccessException)
            {
                // Some managed environments make the user settings folder read-only.
            }
        }

        private void UpdateHighlightForActiveSheet(Excel.Worksheet worksheet = null)
        {
            worksheet = worksheet ?? Application.ActiveSheet as Excel.Worksheet;
            if (worksheet == null)
            {
                RemoveHighlight();
                return;
            }

            Excel.Range target = null;
            try
            {
                target = Application.ActiveCell as Excel.Range;
                UpdateHighlight(worksheet, target);
            }
            finally
            {
                ReleaseComObject(target);
            }
        }

        private void UpdateHighlight(Excel.Worksheet worksheet, Excel.Range target)
        {
            if (!highlightEnabled)
            {
                RemoveHighlight();
                return;
            }

            if (worksheet == null || target == null)
            {
                RemoveHighlight();
                return;
            }

            try
            {
                if (highlightedWorksheet != null && !SameComObject(highlightedWorksheet, worksheet))
                {
                    RemoveHighlight();
                }

                if (highlightCondition == null)
                {
                    RemoveStaleHighlightConditions(worksheet);
                    highlightCondition = CreateHighlightCondition(worksheet);
                    highlightedWorksheet = worksheet;
                }

                Excel.Workbook workbook = null;
                bool wasSaved = false;
                try
                {
                    workbook = worksheet.Parent as Excel.Workbook;
                    wasSaved = workbook != null && workbook.Saved;
                    highlightCondition.Modify(
                        Excel.XlFormatConditionType.xlExpression,
                        Type.Missing,
                        BuildHighlightFormula(target.Row, target.Column),
                        Type.Missing);
                }
                finally
                {
                    RestoreWorkbookSavedState(workbook, wasSaved);
                    ReleaseComObject(workbook);
                }
            }
            catch (COMException)
            {
                // A protected sheet or a workbook closing during an event can reject
                // conditional-format changes. Clear our state and try again next time.
                RemoveHighlight();
            }
        }

        private Excel.FormatCondition CreateHighlightCondition(Excel.Worksheet worksheet)
        {
            Excel.Range allCells = null;
            Excel.FormatConditions conditions = null;
            Excel.Workbook workbook = null;
            bool wasSaved = false;
            try
            {
                workbook = worksheet.Parent as Excel.Workbook;
                wasSaved = workbook != null && workbook.Saved;
                allCells = worksheet.Cells;
                conditions = allCells.FormatConditions;
                Excel.FormatCondition condition = (Excel.FormatCondition)conditions.Add(
                    Excel.XlFormatConditionType.xlExpression,
                    Type.Missing,
                    BuildHighlightFormula(1, 1),
                    Type.Missing);
                condition.Interior.Color = ColorTranslator.ToOle(highlightColor);
                condition.StopIfTrue = false;
                return condition;
            }
            finally
            {
                ReleaseComObject(conditions);
                ReleaseComObject(allCells);
                RestoreWorkbookSavedState(workbook, wasSaved);
                ReleaseComObject(workbook);
            }
        }

        private void RemoveStaleHighlightConditions(Excel.Worksheet worksheet)
        {
            Excel.Range allCells = null;
            Excel.FormatConditions conditions = null;
            Excel.Workbook workbook = null;
            bool wasSaved = false;
            try
            {
                workbook = worksheet.Parent as Excel.Workbook;
                wasSaved = workbook != null && workbook.Saved;
                allCells = worksheet.Cells;
                conditions = allCells.FormatConditions;
                for (int index = conditions.Count; index >= 1; index--)
                {
                    Excel.FormatCondition condition = null;
                    try
                    {
                        condition = (Excel.FormatCondition)conditions[index];
                        string formula = condition.Formula1 as string;
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf(HighlightFormulaMarker, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            condition.Delete();
                        }
                    }
                    finally
                    {
                        ReleaseComObject(condition);
                    }
                }
            }
            finally
            {
                ReleaseComObject(conditions);
                ReleaseComObject(allCells);
                RestoreWorkbookSavedState(workbook, wasSaved);
                ReleaseComObject(workbook);
            }
        }

        private void RemoveHighlight()
        {
            Excel.FormatCondition condition = highlightCondition;
            Excel.Worksheet worksheet = highlightedWorksheet;
            highlightCondition = null;
            highlightedWorksheet = null;

            if (condition == null)
            {
                ReleaseComObject(worksheet);
                return;
            }

            Excel.Workbook workbook = null;
            bool wasSaved = false;
            try
            {
                workbook = worksheet == null ? null : worksheet.Parent as Excel.Workbook;
                wasSaved = workbook != null && workbook.Saved;
                condition.Delete();
            }
            catch (COMException)
            {
                // The workbook may already have been closed.
            }
            finally
            {
                RestoreWorkbookSavedState(workbook, wasSaved);
                ReleaseComObject(workbook);
                ReleaseComObject(condition);
                ReleaseComObject(worksheet);
            }
        }

        private bool IsHighlightedWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null || highlightedWorksheet == null)
            {
                return false;
            }

            Excel.Workbook highlightedWorkbook = null;
            try
            {
                highlightedWorkbook = highlightedWorksheet.Parent as Excel.Workbook;
                return SameComObject(highlightedWorkbook, workbook);
            }
            finally
            {
                ReleaseComObject(highlightedWorkbook);
            }
        }

        private static void RestoreWorkbookSavedState(Excel.Workbook workbook, bool wasSaved)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                workbook.Saved = wasSaved;
            }
            catch (COMException)
            {
                // Excel may reject the assignment while a workbook is closing.
            }
        }

        private static string BuildHighlightFormula(int row, int column)
        {
            // Exclude the active cell so its normal fill remains visible inside the
            // selection border, matching the WPS-style crosshair appearance.
            return string.Format(
                "=AND(OR(AND(ROW()={0},COLUMN()<>{1}),AND(COLUMN()={1},ROW()<>{0})),N(\"{2}\")=0)",
                row,
                column,
                HighlightFormulaMarker);
        }

        private static bool SameComObject(object first, object second)
        {
            if (first == null || second == null)
            {
                return false;
            }

            if (ReferenceEquals(first, second))
            {
                return true;
            }

            IntPtr firstPointer = IntPtr.Zero;
            IntPtr secondPointer = IntPtr.Zero;
            try
            {
                firstPointer = Marshal.GetIUnknownForObject(first);
                secondPointer = Marshal.GetIUnknownForObject(second);
                return firstPointer == secondPointer;
            }
            catch (COMException)
            {
                return false;
            }
            catch (ArgumentException)
            {
                return false;
            }
            finally
            {
                if (firstPointer != IntPtr.Zero)
                {
                    Marshal.Release(firstPointer);
                }

                if (secondPointer != IntPtr.Zero)
                {
                    Marshal.Release(secondPointer);
                }
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SpotLightRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
