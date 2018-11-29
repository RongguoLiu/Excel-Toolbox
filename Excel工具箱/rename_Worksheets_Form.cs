using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Excel工具箱
{
    public partial class rename_Worksheets_Form : Form
    {
        int[] RelatedRangePosition = new int[] { 0, 0 };

        public rename_Worksheets_Form()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void IsValueRelated_CheckedChanged(object sender, EventArgs e)
        {
            if (IsValueRelated.Checked)
            {
                RelatedRangePosition = Globals.ThisAddIn.CellSelector("请选择相关单元格");
                if (RelatedRangePosition[0] == 0 || RelatedRangePosition[1] == 0)
                {
                    RelatedRangePosition = new int[] { 0, 0 };
                    IsValueRelated.Checked = false;
                    IsValueRelated.Text = "新文件名与单元格值相关";
                    insert_CellValue.Enabled = false;
                }
                else
                {
                    IsValueRelated.Text = "新文件名与该单元格值相关：R" + RelatedRangePosition[0].ToString() + "C" + RelatedRangePosition[1].ToString();
                    insert_CellValue.Enabled = true;
                }
            }
            else
            {
                RelatedRangePosition = new int[] { 0, 0 };
                IsValueRelated.Text = "新文件名与单元格值相关";
                insert_CellValue.Enabled = false;
                NewName.Text = "";
            }
        }

        private void insert_CellValue_Click(object sender, EventArgs e)
        {
            NewName.Text = NewName.Text + "^Val";
        }

        private void insert_NumA_Click(object sender, EventArgs e)
        {
            NewName.Text = NewName.Text + "^1";
        }

        private void insert_NumC_Click(object sender, EventArgs e)
        {
            NewName.Text = NewName.Text + "^一";
        }

        private void begin_Rename_Click(object sender, EventArgs e)
        {
            if (!Globals.ThisAddIn.ActiveWorkbookExists()) return;
            string NewnameCurrentSheet;
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                NewnameCurrentSheet = NewName.Text;
                if (IsValueRelated.Checked) NewnameCurrentSheet = NewnameCurrentSheet.Replace("^Val", worksheet.Cells[RelatedRangePosition[0], RelatedRangePosition[1]].Value);
                NewnameCurrentSheet = NewnameCurrentSheet.Replace("^1", worksheet.Index.ToString());
                NewnameCurrentSheet = NewnameCurrentSheet.Replace("^一", Globals.ThisAddIn.DigiInChinese(worksheet.Index));
                for(int i = 0; true; i++)
                {
                    try
                    {
                        if (i == 0) worksheet.Name = NewnameCurrentSheet;
                        else worksheet.Name = NewnameCurrentSheet + "(" + i.ToString() + ")";
                        break;
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
        }
    }
}
