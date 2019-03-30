using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
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
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: "*.*,*.*", Title: "请选择要转换的PPT", MultiSelect: true);
            if (FileOpen.GetType() == typeof(bool)) return;
            int ConvertNum = ((System.Collections.IList)FileOpen).Count;
            PowerPoint.Application PowerPointApplication = new PowerPoint.Application();
            PowerPoint.Presentation ppt;
            string path;
            for (int counter = 1; counter <= ConvertNum; counter++)
            {
                try
                {
                    ppt = PowerPointApplication.Presentations.Open((string)((System.Collections.IList)FileOpen)[counter]);
                    path = (string)((System.Collections.IList)FileOpen)[counter] + ".pdf";
                    path.Replace(".pptx", "");
                    path.Replace(".ppt", "");
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
            Marshal.ReleaseComObject(PowerPointApplication);
            MessageBox.Show("转换完成");
        }

        private void WordToPDF_Click(object sender, EventArgs e)
        {
            object FileOpen = Globals.ThisAddIn.Application.GetOpenFilename(FileFilter: "*.*,*.*", Title: "请选择要转换的PPT", MultiSelect: true);
            if (FileOpen.GetType() == typeof(bool)) return;
            int ConvertNum = ((System.Collections.IList)FileOpen).Count;
            Word.Application WordApplication = new Word.Application();
            Word.Document document;
            string path;
            for (int counter = 1; counter <= ConvertNum; counter++)
            {
                try
                {
                    document = WordApplication.Documents.Open((string)((System.Collections.IList)FileOpen)[counter]);
                    path = (string)((System.Collections.IList)FileOpen)[counter] + ".pdf";
                    path.Replace(".docx", "");
                    path.Replace(".doc", "");
                    document.ExportAsFixedFormat(OutputFileName: path, ExportFormat: WdExportFormat.wdExportFormatPDF);
                    document.Close();
                }
                catch
                {
                    MessageBox.Show("出现了错误，文件名：" + (string)((System.Collections.IList)FileOpen)[counter]);
                    continue;
                }
            }
            Activate();
            WordApplication.Quit();
            MessageBox.Show("转换完成");

        }

        private void PowerPointToSelection_Click(object sender, EventArgs e)
        {
        }
    }
}
