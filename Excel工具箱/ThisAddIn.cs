using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Excel工具箱
{
    public partial class ThisAddIn
    {
        private DigitToChnText digitToChnText = new DigitToChnText();
        private const string FileFitter = "Microsoft Excel文件(*.xlsx),*.xlsx,Excel 97-2003 工作簿(*.xls),*xls,CSV(逗号分隔)(*.csv),*.csv";
        public rename_Worksheets_Form SheetRenamer = new rename_Worksheets_Form();
        public Random random = new Random();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public int[] CellSelector(string Prompt)
        {
            Excel.Range range;
            int[] RangePosition = new int[2];
            try
            {
                range = (Application.InputBox(Prompt: Prompt, Type: 8));
                range = range.Range["A1"];
                RangePosition[0] = range.Row;
                RangePosition[1] = range.Column;
                return RangePosition;
            }
            catch
            {
                return new int[] { 0, 0 };
            }
        }
        public string DigiInChinese(int i)
        {
            return digitToChnText.Convert(i.ToString());
        }
        public string RandomString()
        {
            return random.Next(1000, 9999).ToString() + random.Next(1000, 9999).ToString();
        }
        public string[] SortStrings(string[] ArrayToSort)
        {
            return SortUtil.SortArray(ArrayToSort, 1);
        }
        public void MergeBooks(bool RequireNewBook, bool MergeAllSheets)
        {
            Excel.Workbook destWorkbook, sourceWorkbook;
            int currentSheetIndex = 1;
            int MergeNum;
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: FileFitter, MultiSelect: true, Title: "请选择需要合并的工作簿");
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
        public void MergeSheets(int HeadRowNum, int ContentRowNum, bool FormulaEmbeded)
        {
            if (!ActiveWorkbookExists()) return;
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
                if (HeadRowNum != 0) RangeCopy(sourceWorkbook.Sheets[2].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                for (int CurrentSheetIndex = 2; CurrentSheetIndex <= Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count; CurrentSheetIndex++)
                {
                    RangeCopy(sourceWorkbook.Sheets[CurrentSheetIndex].Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
                    CurrentRowIndex = CurrentRowIndex + ContentRowNum;
                }
                destWorksheet.Cells[1].Select();
            }
            if (ContentRowNum == 0)
            {
                int CopyRowBegin, CopyRowEnd, CurrentRowIndex;
                CopyRowBegin = HeadRowNum + 1;
                CurrentRowIndex = 1 + HeadRowNum;
                if (HeadRowNum != 0) RangeCopy(sourceWorkbook.Sheets[2].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                for (int CurrentSheetIndex = 2; CurrentSheetIndex <= Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count; CurrentSheetIndex++)
                {
                    CopyRowEnd = FirstEmptyRowOf(sourceWorkbook.Worksheets[CurrentSheetIndex], 10) - 1;
                    if (CopyRowEnd < CopyRowBegin) continue;
                    RangeCopy(sourceWorkbook.Sheets[CurrentSheetIndex].Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
                    CurrentRowIndex = CurrentRowIndex + 1 + CopyRowEnd - CopyRowBegin;
                }
                destWorksheet.Cells[1].Select();
            }

            Clipboard.Clear();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        public void MergeSheetsInBooks(bool MergeAllSheets, int HeadRowNum, int ContentRowNum, bool FormulaEmbeded)
        {
            Excel.Workbook destWorkbook, sourceWorkbook;
            Excel.Worksheet destWorksheet;
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: FileFitter, MultiSelect: true, Title: "请选择需要合并的工作簿");
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
                    if (counter == 1 && HeadRowNum != 0) RangeCopy(sourceWorkbook.Sheets[1].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                    foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Worksheets)
                    {
                        if (!MergeAllSheets && sourceWorksheet.Index > 1) break;
                        RangeCopy(sourceWorksheet.Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
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
                    if (counter == 1 && HeadRowNum != 0) RangeCopy(sourceWorkbook.Sheets[1].Rows["1:" + HeadRowNum.ToString()], destWorksheet.Rows[1], FormulaEmbeded);
                    foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Worksheets)
                    {
                        if (!MergeAllSheets && sourceWorksheet.Index > 1) break;
                        CopyRowEnd = FirstEmptyRowOf(sourceWorksheet, 10) - 1;
                        if (CopyRowEnd < CopyRowBegin) continue;
                        RangeCopy(sourceWorksheet.Rows[CopyRowBegin.ToString() + ":" + CopyRowEnd.ToString()], destWorksheet.Rows[CurrentRowIndex], FormulaEmbeded);
                        CurrentRowIndex = CurrentRowIndex + 1 + CopyRowEnd - CopyRowBegin;
                    }
                    sourceWorkbook.Close();
                }
            }
            destWorksheet.Cells[1].Select();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        public void RangeCopy(Excel.Range source, Excel.Range dest, bool IsformulaEmbeded)
        {
            if (IsformulaEmbeded)
            {
                source.Copy();
                dest.PasteSpecial(XlPasteType.xlPasteAllUsingSourceTheme, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                dest.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            }
            else source.Copy(dest);
        }
        public void ConvertWorkbookFormat(Excel.Workbook Workbook, int TargetFormatCode, string TargetFormat)
        {
            //MessageBox.Show(((int)workbook.FileFormat).ToString());
            //return;
            if (TargetFormatCode != 0) Workbook.SaveAs(Workbook.Name + TargetFormat, (XlFileFormat)TargetFormatCode, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);
            else Workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, Workbook.Name + ".pdf");
        }
        public int FirstEmptyRowOf(Excel.Worksheet testSheet, int testCellsNumEachRow)
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
        public bool ActiveWorkbookExists()
        {
            try
            {
                string test = Application.ActiveWorkbook.Name;
                return true;
            }
            catch
            {
                return false;
            }
        }

        #region 从Internet上copy的代码

        /// 本程序用于将小写数字变成大写中文数字
        /// 算法设计：黄晶
        /// 程序制作：黄晶
        /// 时间：2004年8月12日
        class DigitToChnText
        {
            private readonly char[] chnText;
            private readonly char[] chnDigit;

            public DigitToChnText()
            {
                chnText = new char[] { '零', '一', '二', '三', '四', '五', '六', '七', '八', '九' };
                chnDigit = new char[] { '十', '百', '千', '万', '亿' };
            }

            public string Convert(string strDigit)
            {
                // 检查输入数字
                decimal dec;
                try
                {
                    dec = decimal.Parse(strDigit);
                }
                catch (FormatException)
                {
                    throw new Exception("输入数字的格式不正确。");
                }
                catch (Exception e)
                {
                    throw e;
                }

                if (dec <= -10000000000000000m || dec >= 10000000000000000m)
                {
                    throw new Exception("输入数字太大或太小，超出范围。");
                }

                StringBuilder strResult = new StringBuilder();

                // 提取符号部分
                // '+'在最前
                if ("+" == strDigit.Substring(0, 1))
                {
                    strDigit = strDigit.Substring(1);
                }
                // '-'在最前
                else if ("-" == strDigit.Substring(0, 1))
                {
                    strResult.Append('负');
                    strDigit = strDigit.Substring(1);
                }
                // '+'在最后
                else if ("+" == strDigit.Substring(strDigit.Length - 1, 1))
                {
                    strDigit = strDigit.Substring(0, strDigit.Length - 1);
                }
                // '-'在最后
                else if ("-" == strDigit.Substring(strDigit.Length - 1, 1))
                {
                    strResult.Append('负');
                    strDigit = strDigit.Substring(0, strDigit.Length - 1);
                }

                // 提取整数和小数部分
                int indexOfPoint;
                if (-1 == (indexOfPoint = strDigit.IndexOf('.'))) // 如果没有小数部分
                {
                    strResult.Append(ConvertIntegral(strDigit));
                }
                else // 有小数部分
                {
                    // 先转换整数部分
                    if (0 == indexOfPoint) // 如果“.”是第一个字符
                    {
                        strResult.Append('零');
                    }
                    else
                    {
                        strResult.Append(ConvertIntegral(strDigit.Substring(0, indexOfPoint)));
                    }

                    // 再转换小数部分
                    if (strDigit.Length - 1 != indexOfPoint) // 如果“.”不是最后一个字符
                    {
                        strResult.Append('点');
                        strResult.Append(ConvertFractional(strDigit.Substring(indexOfPoint + 1)));
                    }
                }

                return strResult.ToString();
            }

            // 转换整数部分
            protected string ConvertIntegral(string strIntegral)
            {
                // 去掉数字前面所有的'0'
                // 并把数字分割到字符数组中
                char[] integral = ((long.Parse(strIntegral)).ToString()).ToCharArray();

                // 变成中文数字并添加中文数位
                StringBuilder strInt = new StringBuilder();

                int i;
                int digit;
                digit = integral.Length - 1;

                // 处理最高位到十位的所有数字
                for (i = 0; i < integral.Length - 1; i++)
                {
                    strInt.Append(chnText[integral[i] - '0']);

                    if (0 == digit % 4)     // '万' 或 '亿'
                    {
                        if (4 == digit || 12 == digit)
                        {
                            strInt.Append(chnDigit[3]); // '万'
                        }
                        else if (8 == digit)
                        {
                            strInt.Append(chnDigit[4]); // '亿'
                        }
                    }
                    else         // '十'，'百'或'千'
                    {
                        strInt.Append(chnDigit[digit % 4 - 1]);
                    }

                    digit--;
                }

                // 如果个位数不是'0'
                // 或者个位数为‘0’但只有一位数
                // 则添加相应的中文数字
                if ('0' != integral[integral.Length - 1] || 1 == integral.Length)
                {
                    strInt.Append(chnText[integral[i] - '0']);
                }

                // 遍历整个字符串
                i = 0;
                while (i < strInt.Length)
                {
                    int j = i;

                    bool bDoSomething = false;

                    // 查找所有相连的“零X”结构
                    while (j < strInt.Length - 1 && "零" == strInt.ToString().Substring(j, 1))
                    {
                        string strTemp = strInt.ToString().Substring(j + 1, 1);

                        // 如果是“零万”或者“零亿”则停止查找
                        if ("万" == strTemp || "亿" == strTemp)
                        {
                            bDoSomething = true;
                            break;
                        }

                        j += 2;
                    }

                    if (j != i) // 如果找到“零X”结构，则全部删除
                    {
                        strInt = strInt.Remove(i, j - i);

                        // 除了在最尾处，或后面不是"零万"或"零亿"的情况下, 
                        // 其他处均补入一个“零”
                        if (i <= strInt.Length - 1 && !bDoSomething)
                        {
                            strInt = strInt.Insert(i, '零');
                            i++;
                        }
                    }

                    if (bDoSomething) // 如果找到"零万"或"零亿"结构
                    {
                        strInt = strInt.Remove(i, 1); // 去掉'零'
                        i++;
                        continue;
                    }

                    // 指针每次可移动2位
                    i += 2;
                }

                // 遇到“亿万”变成“亿零”或"亿"
                int index = strInt.ToString().IndexOf("亿万");
                if (-1 != index)
                {
                    if (strInt.Length - 2 != index &&  // 如果"亿万"不在最后
                     (index + 2 < strInt.Length && "零" != strInt.ToString().Substring(index + 2, 1))) // 并且其后没有"零"
                        strInt = strInt.Replace("亿万", "亿零", index, 2);
                    else
                        strInt = strInt.Replace("亿万", "亿", index, 2);
                }

                // 开头为“一十”改为“十”
                if (strInt.Length > 1 && "一十" == strInt.ToString().Substring(0, 2))
                {
                    strInt = strInt.Remove(0, 1);
                }

                return strInt.ToString();
            }

            // 转换小数部分
            protected string ConvertFractional(string strFractional)
            {
                char[] fractional = strFractional.ToCharArray();
                StringBuilder strFrac = new StringBuilder();

                // 变成中文数字
                int i;
                for (i = 0; i < fractional.Length; i++)
                {
                    strFrac.Append(chnText[fractional[i] - '0']);
                }

                return strFrac.ToString();
            }
        }
        /// 作者：Rex_IT
        /// 来源：CSDN
        /// 原文：https://blog.csdn.net/liuxiaoshuang002/article/details/53761838 
        public class SortUtil
        {
            /// <summary>
            /// 排序
            /// </summary>
            /// <param name="arr">排序字符串数组</param>
            /// <param name="type">类型：1发音，2笔画</param>
            /// <returns></returns>
            public static string[] SortArray(string[] arr, int? type = 1)
            {
                //发音 LCID：0x00000804
                if (type.Value == 1)
                {
                    CultureInfo PronoCi = new CultureInfo(2052);
                    Array.Sort(arr);
                }
                else
                {
                    //笔画数 LCID：0x00020804
                    CultureInfo StrokCi = new CultureInfo(133124);
                    Thread.CurrentThread.CurrentCulture = StrokCi;
                    Array.Sort(arr);
                }
                return arr;
            }
            /// <summary>
            /// 排序
            /// </summary>
            /// <param name="arrlist">排序字符串数组</param>
            /// <param name="type">类型：1发音，2笔画</param>
            /// <returns></returns>
            public static List<string> SortList(List<string> arrlist, int? type = 1)
            {
                string[] arr = arrlist.ToArray();
                //发音 LCID：0x00000804
                if (type.Value == 1)
                {
                    CultureInfo PronoCi = new CultureInfo(2052);
                    Array.Sort(arr);
                }
                else
                {
                    //笔画数 LCID：0x00020804
                    CultureInfo StrokCi = new CultureInfo(133124);
                    Thread.CurrentThread.CurrentCulture = StrokCi;
                    Array.Sort(arr);
                }
                return arr.ToList<string>();
            }
        }

        #endregion

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
