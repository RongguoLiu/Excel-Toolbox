using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Threading;
using System.Diagnostics;

namespace Excel工具箱
{
    public partial class UniversalConvert_Form : Form
    {
        public UniversalConvert_Form()
        {
            InitializeComponent();
        }

        private void PowerPointToPDF_Click(object sender, EventArgs e)
        {
            bool needKill = true;
            Process[] pros = Process.GetProcesses();
            for (int i = 0; i < pros.Count(); i++)
            {
                if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                {
                    needKill = false;
                }
            }
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: "*.*,*.*", Title: "请选择要转换的PPT", MultiSelect: true);
            if (FileOpen.GetType() == typeof(bool)) return;
            int ConvertNum = ((System.Collections.IList)FileOpen).Count;
            PowerPoint.Application application = new PowerPoint.Application();
            for (int counter = 1; counter <= ConvertNum; counter++)
            {
                try
                {
                    PowerPoint.Presentation ppt;
                    ppt = application.Presentations.Open((string)((System.Collections.IList)FileOpen)[counter]);
                    ppt.ExportAsFixedFormat(Path: (string)((System.Collections.IList)FileOpen)[counter] + ".pdf", FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF);
                    ppt.Close();
                }
                catch
                {
                    MessageBox.Show("出现了错误，文件名：" + (string)((System.Collections.IList)FileOpen)[counter]);
                    continue;
                }
            }
            Activate();
            if (needKill)
            {
                pros = Process.GetProcesses();
                for (int i = 0; i < pros.Count(); i++)
                {
                    if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                    {
                        pros[i].Kill();
                    }
                }
            }
            MessageBox.Show("转换完成");
        }
    }
}
