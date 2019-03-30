using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace Excel工具箱
{
    public partial class AboutBox_Form : Form
    {
        //ThreadStart childref;
        //Thread childThread;

        public AboutBox_Form()
        {
            InitializeComponent();
            //childref = new ThreadStart(OpenSupportWorkbook);
            //childThread = new Thread(childref);
        }

        private void OpenSupportWorkbook()
        {
            Globals.ThisAddIn.Application.Workbooks.Open("https://github.com/RongguoLiu/Excel-Toolbox/raw/master/%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E.xlsx");
        }

        private void AboutBox_Form_Shown(object sender, EventArgs e)
        {
            OpenSupportWorkbook();
            Close();
        }

        private void AboutBox_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            //childThread.Abort();
        }
    }
}
