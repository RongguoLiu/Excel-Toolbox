using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace Excel工具箱
{
    public partial class RibbonUI
    {
        private void RibbonUI_Load(object sender, RibbonUIEventArgs e)
        {
            mergesheets_HeadRowNum.SelectedItemIndex = 1;
            mergesheets_contentRowNum.SelectedItemIndex = 1;
            convert_sourceFormat.SelectedItemIndex = 3;
            convert_targetFormat.SelectedItemIndex = 0;
            //Globals.ThisAddIn.Application.SheetDeactivate += new AppEvents_SheetDeactivateEventHandler(Application_SheetDeactivate);
            Globals.ThisAddIn.Application.WorkbookDeactivate += new Excel.AppEvents_WorkbookDeactivateEventHandler(Application_WorkbookDeactivate);
            Globals.ThisAddIn.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
        }
        //Button Handlers
        private void mergebooks_BeginMerge_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.MergeBooks(mergebooks_RequireNewBook.Checked, mergebooks_MergeAllSheets.Checked);
        }
        private void mergesheets_BeginMerge_Click(object sender, RibbonControlEventArgs e)
        {
            if (mergebooks_AIO.Checked) Globals.ThisAddIn.MergeSheetsInBooks(mergebooks_MergeAllSheets.Checked, mergesheets_HeadRowNum.SelectedItemIndex, mergesheets_contentRowNum.SelectedItemIndex, mergesheets_isFunctionEmbeded.Checked);
            else Globals.ThisAddIn.MergeSheets(mergesheets_HeadRowNum.SelectedItemIndex, mergesheets_contentRowNum.SelectedItemIndex, mergesheets_isFunctionEmbeded.Checked);
        }
        private void convert_Exchange_Click(object sender, RibbonControlEventArgs e)
        {
            int temp = convert_sourceFormat.SelectedItemIndex;
            convert_sourceFormat.SelectedItemIndex = convert_targetFormat.SelectedItemIndex;
            convert_targetFormat.SelectedItemIndex = temp;
        }
        private void convert_BeginConvert_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook WorkbookToConvert;
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: (convert_sourceFormat.SelectedItem.Label + "," + convert_sourceFormat.SelectedItem.Label), MultiSelect: true, Title: "请选择需要转换的工作簿");
            if (FileOpen.GetType() == typeof(bool)) return;
            int ConvertNum = ((System.Collections.IList)FileOpen).Count;
            int TargetFormatCode = Convert.ToInt32(convert_targetFormat.SelectedItem.Tag);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            for (int counter = 1; counter <= ConvertNum; counter++)
            {
                try
                {
                    WorkbookToConvert = Globals.ThisAddIn.Application.Workbooks.Open(Filename: (string)((System.Collections.IList)FileOpen)[counter]);
                    Globals.ThisAddIn.ConvertWorkbookFormat(WorkbookToConvert, TargetFormatCode, convert_targetFormat.SelectedItem.Label);
                    WorkbookToConvert.Close();
                }
                catch
                {
                    MessageBox.Show("出现了错误，文件名："+ (string)((System.Collections.IList)FileOpen)[counter]);
                    continue;
                }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            MessageBox.Show("转换完成");
        }
        private void cellActions_ConvertToValue_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.SelectedRangeCount() == 0)
            {
                MessageBox.Show("操作未执行");
                return;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            double d;
            foreach(Excel.Range range in (Excel.Range)Globals.ThisAddIn.Application.Selection)
            {
                if (range.Text == "") continue;
                if(double.TryParse(range.Text, out d)) range.Value = d;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void cellActions_ConvertToString_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.SelectedRangeCount() == 0)
            {
                MessageBox.Show("操作未执行");
                return;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Excel.Range range in (Excel.Range)Globals.ThisAddIn.Application.Selection)
            {
                try
                {
                    range.Value = range.Value.ToString();
                }
                catch
                {
                    continue;
                }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void rename_RenameWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.ActiveWorkbookExists()) return;
            try
            {
                Globals.ThisAddIn.SheetRenamer.Show();
            }
            catch
            {
                Globals.ThisAddIn.SheetRenamer = new rename_Worksheets_Form();
                Globals.ThisAddIn.SheetRenamer.Show();
            }
        }
        private void rename_SortSheets_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.ActiveWorkbookExists()) return;
            int SheetsCount = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count;
            string[] SheetsNames = new string[SheetsCount];
            for (int counter = 0; counter < SheetsCount; counter++) 
            {
                SheetsNames[counter] = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[counter + 1].Name;
            }
            string[] SortedNames = Globals.ThisAddIn.SortStrings(SheetsNames);
            for (int counter = 0; counter < SheetsCount; counter++)
            {
                ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[SortedNames[counter]]).Move(Before: Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[counter + 1]);
            }
        }
        private void rename_RenameWorkbooks_Click(object sender, RibbonControlEventArgs e)
        {

        }
        private void others_DeleteOtherSheets_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.ActiveWorkbookExists()) return;
            double RowsToReserve;
            try
            {
                RowsToReserve = Globals.ThisAddIn.Application.InputBox(Prompt: "保留几张表？默认1张！", Type: 1);
            }
            catch
            {
                return;
            }
            if (RowsToReserve < 1) RowsToReserve = 1;
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (worksheet.Index <= RowsToReserve) continue;
                worksheet.Delete();
            }
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }
        private void others_ClrClipboard_Click(object sender, RibbonControlEventArgs e)
        {
            Clipboard.Clear();
        }
        private void others_LookForFirstEmptyRow_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.ActiveWorkbookExists()) return;
            Globals.ThisAddIn.Application.ActiveSheet.Rows[Globals.ThisAddIn.FirstEmptyRowOf(Globals.ThisAddIn.Application.ActiveSheet, 10)].Select();
        }
        private void others_ConvertUI_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.UniversalConverter.Show();
                Globals.ThisAddIn.UniversalConverter.Activate();
            }
            catch
            {
                Globals.ThisAddIn.UniversalConverter = new UniversalConvert_Form();
                Globals.ThisAddIn.UniversalConverter.Show();
                Globals.ThisAddIn.UniversalConverter.Activate();
            }
        }
        private void groupConvert_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.UniversalConverter.Show();
                Globals.ThisAddIn.UniversalConverter.Activate();
            }
            catch
            {
                Globals.ThisAddIn.UniversalConverter = new UniversalConvert_Form();
                Globals.ThisAddIn.UniversalConverter.Show();
                Globals.ThisAddIn.UniversalConverter.Activate();
            }
        }
        private void dangerzone_tryFix_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }
        private void help_About_Click(object sender, RibbonControlEventArgs e)
        {
            //try
            //{
            //    Globals.ThisAddIn.aboutBox.Show();
            //}
            //catch
            //{
            //    Globals.ThisAddIn.aboutBox = new AboutBox();
            //    Globals.ThisAddIn.aboutBox.Show();
            //}
        }
        //Checkbox & Dropdowns Handlers
        private void mergebooks_AIO_Click(object sender, RibbonControlEventArgs e)
        {
            if (mergebooks_AIO.Checked)
            {
                mergebooks_RequireNewBook.Checked = true;
                mergebooks_RequireNewBook.Enabled = false;
                mergebooks_BeginMerge.Visible = false;
                mergesheets_BeginMerge.Enabled = true;
            }
            else
            {
                mergebooks_BeginMerge.Visible = true;
                if (Globals.ThisAddIn.ActiveWorkbookExists())
                {
                    mergebooks_RequireNewBook.Enabled = true;
                    mergesheets_BeginMerge.Enabled = true;
                }
                else
                {
                    mergebooks_RequireNewBook.Enabled = false;
                    mergesheets_BeginMerge.Enabled = false;
                }
            }
        }
        private void cellActions_HighlightCurrentRC_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.EnableHighlight = cellActions_HighlightCurrentRC.Checked;
            if (Globals.ThisAddIn.ActiveWorkbookExists() && !cellActions_HighlightCurrentRC.Checked) Globals.ThisAddIn.Application.ActiveSheet.Cells.Interior.ColorIndex = -4142;
            if (Globals.ThisAddIn.ActiveWorkbookExists() && cellActions_HighlightCurrentRC.Checked) Globals.ThisAddIn.HighlightCurrentRC();
        }
        private void updateView_Click(object sender, RibbonControlEventArgs e)
        {
            if (dangerzone_updateView.Checked) Globals.ThisAddIn.Application.ScreenUpdating = true;
            else Globals.ThisAddIn.Application.ScreenUpdating = false;
        }
        private void showAlert_Click(object sender, RibbonControlEventArgs e)
        {
            if (dangerzone_showAlert.Checked) Globals.ThisAddIn.Application.DisplayAlerts = true;
            else Globals.ThisAddIn.Application.DisplayAlerts = false;
        }
        private void convert_sourceFormat_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (convert_sourceFormat.SelectedItemIndex == 5 || convert_targetFormat.SelectedItemIndex == 5) convert_Exchange.Enabled = false;
            else convert_Exchange.Enabled = true;
        }
        private void convert_targetFormat_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (convert_sourceFormat.SelectedItemIndex == 5 || convert_targetFormat.SelectedItemIndex == 5) convert_Exchange.Enabled = false;
            else convert_Exchange.Enabled = true;
        }
        //UI Refresher
        private void Application_WorkbookDeactivate(Excel.Workbook wb)
        {
            mergebooks_RequireNewBook.Checked = true;
            mergebooks_RequireNewBook.Enabled = false;
            mergesheets_BeginMerge.Enabled = false;
            cellActions_ConvertToValue.Enabled = false;
            cellActions_ConvertToString.Enabled = false;
            cellActions_HighlightCurrentRC.Enabled = false;
            rename_RenameWorksheets.Enabled = false;
            rename_SortSheets.Enabled = false;
            others_DeleteOtherSheets.Enabled = false;
            LookForFirstEmptyRow.Enabled = false;
        }
        private void Application_WorkbookActivate(Excel.Workbook wb)
        {
            mergebooks_RequireNewBook.Enabled = true;
            mergesheets_BeginMerge.Enabled = true;
            cellActions_ConvertToValue.Enabled = true;
            cellActions_ConvertToString.Enabled = true;
            cellActions_HighlightCurrentRC.Enabled = true;
            rename_RenameWorksheets.Enabled = true;
            rename_SortSheets.Enabled = true;
            others_DeleteOtherSheets.Enabled = true;
            LookForFirstEmptyRow.Enabled = true;
        }
    }
}