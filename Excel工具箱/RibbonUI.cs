using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace Excel工具箱
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            mergesheets_HeadRowNum.SelectedItemIndex = 1;
            mergesheets_contentRowNum.SelectedItemIndex = 1;
            convert_sourceFormat.SelectedItemIndex = 3;
            convert_targetFormat.SelectedItemIndex = 0;
            SheetRenamer = new rename_Worksheets_Form();
        }
        //Consts
        rename_Worksheets_Form SheetRenamer;
        const string FileFitterForMerge = "Microsoft Excel文件(*.xlsx),*.xlsx,Excel 97-2003 工作簿(*.xls),*xls,CSV(逗号分隔)(*.csv),*.csv";
        //Button Handlers
        private void mergebooks_BeginMerge_Click(object sender, RibbonControlEventArgs e)
        {
            MergeBooks(mergebooks_RequireNewBook.Checked, mergebooks_MergeAllSheets.Checked);
        }
        private void mergesheets_BeginMerge_Click(object sender, RibbonControlEventArgs e)
        {
            if (mergebooks_AIO.Checked) MergeSheetsInBooks(mergebooks_MergeAllSheets.Checked, mergesheets_HeadRowNum.SelectedItemIndex, mergesheets_contentRowNum.SelectedItemIndex, mergesheets_isFunctionEmbeded.Checked);
            else MergeSheets(mergesheets_HeadRowNum.SelectedItemIndex, mergesheets_contentRowNum.SelectedItemIndex, mergesheets_isFunctionEmbeded.Checked);
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
                    ConvertWorkbookFormat(WorkbookToConvert, TargetFormatCode, convert_targetFormat.SelectedItem.Label);
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
        private void rename_RenameWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            SheetRenamer = new rename_Worksheets_Form();
            SheetRenamer.Show();
        }
        private void others_DeleteOtherSheets_Click(object sender, RibbonControlEventArgs e)
        {
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
            Globals.ThisAddIn.Application.ActiveSheet.Rows[FirstEmptyRowOf(Globals.ThisAddIn.Application.ActiveSheet, 10)].Select();
        }
        private void dangerzone_tryFix_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }
        private void help_About_Click(object sender, RibbonControlEventArgs e)
        {
            //todo:Draw a about box...
            //AboutBox aboutBox = new AboutBox();
            //aboutBox.Show();
        }
        //Checkbox & Dropdowns Handlers
        private void mergebooks_AIO_Click(object sender, RibbonControlEventArgs e)
        {
            if (mergebooks_AIO.Checked)
            {
                mergebooks_RequireNewBook.Checked = true;
                mergebooks_RequireNewBook.Enabled = false;
                mergebooks_BeginMerge.Visible = false;
            }
            else
            {
                mergebooks_RequireNewBook.Enabled = true;
                mergebooks_BeginMerge.Visible = true;
            }
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
        private void disableConvertExchangeButton()
        {
            if (convert_sourceFormat.SelectedItemIndex == 5 || convert_targetFormat.SelectedItemIndex == 5) convert_Exchange.Enabled = false;
            else convert_Exchange.Enabled = true;
        }
        private void convert_sourceFormat_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            disableConvertExchangeButton();
        }
        private void convert_targetFormat_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            disableConvertExchangeButton();
        }
        //Workers
        private void MergeBooks(bool RequireNewBook, bool MergeAllSheets)
        {
            Excel.Workbook destWorkbook, sourceWorkbook;
            int currentSheetIndex = 1;
            int MergeNum;
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: FileFitterForMerge, MultiSelect: true, Title: "请选择需要合并的工作簿");
            if (FileOpen.GetType() == typeof(bool)) return;
            MergeNum = ((System.Collections.IList)FileOpen).Count;
            if (RequireNewBook) destWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
            else try
                {
                    destWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                }
                catch
                {
                    destWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
                }

            Globals.ThisAddIn.Application.ScreenUpdating = false;
            for (int counter = 1; counter <= MergeNum; counter++)
            {
                sourceWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(Filename: (string)((System.Collections.IList)FileOpen)[counter]);
                foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Worksheets)
                {
                    if (!MergeAllSheets && sourceWorksheet.Index > 1) break;
                    sourceWorksheet.Copy(destWorkbook.Worksheets[currentSheetIndex]);
                    currentSheetIndex++;
                }
                sourceWorkbook.Close();
            }
            if (RequireNewBook)
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                destWorkbook.Worksheets[currentSheetIndex].Delete();
                Globals.ThisAddIn.Application.DisplayAlerts = true;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void MergeSheets(int HeadRowNum, int ContentRowNum, bool FormulaEmbeded)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[1].Activate();
            Excel.Worksheet destWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            Excel.Workbook sourceWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            try
            {
                destWorksheet.Name = "Merge";
            }
            catch
            {
                destWorksheet.Delete();
                MessageBox.Show("确保工作簿中没有以'Merge'为名的工作表，再试一次");
                return;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            if (ContentRowNum != 0)
            {
                int CopyRowBegin, CopyRowEnd, CurrentRowIndex;
                CopyRowBegin = HeadRowNum + 1;
                CopyRowEnd = HeadRowNum + ContentRowNum;
                CurrentRowIndex = 1 + HeadRowNum;
                if (HeadRowNum != 0) RowCP(sourceWorkbook.Sheets[2].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                for (int CurrentSheetIndex = 2; CurrentSheetIndex <= Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count; CurrentSheetIndex++)
                {
                    RowCP(sourceWorkbook.Sheets[CurrentSheetIndex].Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
                    CurrentRowIndex = CurrentRowIndex + ContentRowNum;
                }
                destWorksheet.Cells[1].Select();
            }
            if (ContentRowNum == 0)
            {
                int CopyRowBegin, CopyRowEnd, CurrentRowIndex;
                CopyRowBegin = HeadRowNum + 1;
                CurrentRowIndex = 1 + HeadRowNum;
                if (HeadRowNum != 0) RowCP(sourceWorkbook.Sheets[2].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                for (int CurrentSheetIndex = 2; CurrentSheetIndex <= Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count; CurrentSheetIndex++)
                {
                    CopyRowEnd = FirstEmptyRowOf(sourceWorkbook.Worksheets[CurrentSheetIndex], 10) - 1;
                    if (CopyRowEnd < CopyRowBegin) continue;
                    RowCP(sourceWorkbook.Sheets[CurrentSheetIndex].Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
                    CurrentRowIndex = CurrentRowIndex + 1 + CopyRowEnd - CopyRowBegin;
                }
                destWorksheet.Cells[1].Select();
            }

            Clipboard.Clear();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void MergeSheetsInBooks(bool MergeAllSheets, int HeadRowNum, int ContentRowNum, bool FormulaEmbeded)
        {
            Excel.Workbook destWorkbook, sourceWorkbook;
            Excel.Worksheet destWorksheet;
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: FileFitterForMerge, MultiSelect: true, Title: "请选择需要合并的工作簿");
            if (FileOpen.GetType() == typeof(bool)) return;
            int MergeNum = ((System.Collections.IList)FileOpen).Count;
            destWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
            destWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            destWorksheet.Name = "Merge";
            int CopyRowBegin, CopyRowEnd, CurrentRowIndex;
            CurrentRowIndex = CopyRowBegin = HeadRowNum + 1;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            if (ContentRowNum != 0)
            {
                CopyRowEnd = HeadRowNum + ContentRowNum;
                for (int counter = 1; counter <= MergeNum; counter++)
                {
                    sourceWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(Filename: (string)((System.Collections.IList)FileOpen)[counter]);
                    if (counter == 1 && HeadRowNum != 0) RowCP(sourceWorkbook.Sheets[1].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                    foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Worksheets)
                    {
                        if (!MergeAllSheets && sourceWorksheet.Index > 1) break;
                        RowCP(sourceWorksheet.Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
                        CurrentRowIndex = CurrentRowIndex + 1 + CopyRowEnd - CopyRowBegin;
                    }
                    sourceWorkbook.Close();
                }
            }
            if (ContentRowNum == 0)
            {
                for (int counter = 1; counter <= MergeNum; counter++)
                {
                    sourceWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(Filename: (string)((System.Collections.IList)FileOpen)[counter]);
                    if (counter == 1 && HeadRowNum != 0) RowCP(sourceWorkbook.Sheets[1].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                    foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Worksheets)
                    {
                        if (!MergeAllSheets && sourceWorksheet.Index > 1) break;
                        CopyRowEnd = FirstEmptyRowOf(sourceWorksheet, 10) - 1;
                        if (CopyRowEnd < CopyRowBegin) continue;
                        RowCP(sourceWorksheet.Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
                        CurrentRowIndex = CurrentRowIndex + 1 + CopyRowEnd - CopyRowBegin;
                    }
                    sourceWorkbook.Close();
                }
            }
            destWorksheet.Cells[1].Select();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void RowCP(Excel.Range source, Excel.Range dest, bool IsformulaEmbeded)
        {
            if (IsformulaEmbeded)
            {
                source.Copy();
                dest.PasteSpecial(XlPasteType.xlPasteAllUsingSourceTheme, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                dest.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            }
            else source.Copy(dest);
        }
        private void ConvertWorkbookFormat(Excel.Workbook Workbook, int TargetFormatCode, string TargetFormat)
        {
            //MessageBox.Show(((int)workbook.FileFormat).ToString());
            //return;
            if (TargetFormatCode != 0) Workbook.SaveAs(Workbook.Name + TargetFormat, (XlFileFormat)TargetFormatCode, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);
            else Workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, Workbook.Name + ".pdf");
        }
        private int FirstEmptyRowOf(Excel.Worksheet testSheet, int testCellsNumEachRow)
        {
            int counterR, counterC;
            bool isEmpty;
            for (counterR = 1; counterR < 10000; counterR++)
            {
                isEmpty = true;
                for (counterC = 1; counterC < testCellsNumEachRow; counterC++)
                {
                    if (testSheet.Cells[counterR, counterC].Text.Trim() != "")
                    {
                        isEmpty = false;
                        continue;
                    }
                }
                if (isEmpty) return counterR;
            }
            return 0;
        }
    }
}