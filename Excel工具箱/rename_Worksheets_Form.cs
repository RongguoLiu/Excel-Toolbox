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
        int[] RangePosition = new int[] { 0, 0 };

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
                RangePosition = Globals.ThisAddIn.CellSelector();
                if (RangePosition[0] == 0 || RangePosition[1] == 0)
                {
                    RangePosition = new int[] { 0, 0 };
                    IsValueRelated.Checked = false;
                    IsValueRelated.Text = "新文件名与单元格值相关";
                    insert_CellValue.Enabled = false;
                }
                else
                {
                    IsValueRelated.Text = "新文件名与该单元格值相关：(" + RangePosition[0].ToString() + "," + RangePosition[1].ToString() + ")";
                    insert_CellValue.Enabled = true;
                }
            }
            else
            {
                RangePosition = new int[] { 0, 0 };
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
            string NewnameCurrentSheet;
            foreach(Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                NewnameCurrentSheet = NewName.Text;
                if (IsValueRelated.Checked) NewnameCurrentSheet = NewnameCurrentSheet.Replace("^Val", worksheet.Cells[RangePosition[0], RangePosition[1]].Value);
                NewnameCurrentSheet = NewnameCurrentSheet.Replace("^1", worksheet.Index.ToString());
                NewnameCurrentSheet = NewnameCurrentSheet.Replace("^一", Globals.ThisAddIn.IndexInChinese(worksheet.Index));
                try
                {
                    worksheet.Name = NewnameCurrentSheet;
                }
                catch
                {
                    MessageBox.Show("出现了错误，位于表"+worksheet.Index.ToString()+"。是否重名？");
                    continue;
                }
            }
        }
    }
}
