namespace Excel工具箱
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl15 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl16 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl17 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl18 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl19 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl20 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl21 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl22 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl23 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl24 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl25 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl26 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl27 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl28 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl29 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl30 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl31 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl32 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupMergeBooks = this.Factory.CreateRibbonGroup();
            this.mergebooks_RequireNewBook = this.Factory.CreateRibbonCheckBox();
            this.mergebooks_MergeAllSheets = this.Factory.CreateRibbonCheckBox();
            this.mergebooks_AIO = this.Factory.CreateRibbonCheckBox();
            this.groupMergeSheets = this.Factory.CreateRibbonGroup();
            this.mergesheets_HeadRowNum = this.Factory.CreateRibbonDropDown();
            this.mergesheets_contentRowNum = this.Factory.CreateRibbonDropDown();
            this.mergesheets_isFunctionEmbeded = this.Factory.CreateRibbonCheckBox();
            this.convert = this.Factory.CreateRibbonGroup();
            this.convert_sourceFormat = this.Factory.CreateRibbonDropDown();
            this.convert_targetFormat = this.Factory.CreateRibbonDropDown();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.others = this.Factory.CreateRibbonGroup();
            this.dangerous_zone = this.Factory.CreateRibbonGroup();
            this.updateView = this.Factory.CreateRibbonCheckBox();
            this.showAlert = this.Factory.CreateRibbonCheckBox();
            this.support = this.Factory.CreateRibbonGroup();
            this.mergebooks_BeginMerge = this.Factory.CreateRibbonButton();
            this.mergesheets_BeginMerge = this.Factory.CreateRibbonButton();
            this.convert_Exchange = this.Factory.CreateRibbonButton();
            this.convert_BeginConvert = this.Factory.CreateRibbonButton();
            this.others_DeleteOtherSheets = this.Factory.CreateRibbonButton();
            this.LookForFirstEmptyRow = this.Factory.CreateRibbonButton();
            this.others_ClrClipboard = this.Factory.CreateRibbonButton();
            this.help_About = this.Factory.CreateRibbonButton();
            this.convert_Spreater = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupMergeBooks.SuspendLayout();
            this.groupMergeSheets.SuspendLayout();
            this.convert.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.others.SuspendLayout();
            this.dangerous_zone.SuspendLayout();
            this.support.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupMergeBooks);
            this.tab1.Groups.Add(this.groupMergeSheets);
            this.tab1.Groups.Add(this.convert);
            this.tab1.Groups.Add(this.others);
            this.tab1.Groups.Add(this.dangerous_zone);
            this.tab1.Groups.Add(this.support);
            this.tab1.Label = "工具箱";
            this.tab1.Name = "tab1";
            // 
            // groupMergeBooks
            // 
            this.groupMergeBooks.Items.Add(this.mergebooks_RequireNewBook);
            this.groupMergeBooks.Items.Add(this.mergebooks_MergeAllSheets);
            this.groupMergeBooks.Items.Add(this.mergebooks_AIO);
            this.groupMergeBooks.Items.Add(this.mergebooks_BeginMerge);
            this.groupMergeBooks.Label = "合并工作簿";
            this.groupMergeBooks.Name = "groupMergeBooks";
            // 
            // mergebooks_RequireNewBook
            // 
            this.mergebooks_RequireNewBook.Checked = true;
            this.mergebooks_RequireNewBook.Label = "创建新簿用于合并";
            this.mergebooks_RequireNewBook.Name = "mergebooks_RequireNewBook";
            // 
            // mergebooks_MergeAllSheets
            // 
            this.mergebooks_MergeAllSheets.Label = "合并所有工作表";
            this.mergebooks_MergeAllSheets.Name = "mergebooks_MergeAllSheets";
            // 
            // mergebooks_AIO
            // 
            this.mergebooks_AIO.Label = "一体化操作流程";
            this.mergebooks_AIO.Name = "mergebooks_AIO";
            this.mergebooks_AIO.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mergebooks_AIO_Click);
            // 
            // groupMergeSheets
            // 
            this.groupMergeSheets.Items.Add(this.mergesheets_HeadRowNum);
            this.groupMergeSheets.Items.Add(this.mergesheets_contentRowNum);
            this.groupMergeSheets.Items.Add(this.mergesheets_isFunctionEmbeded);
            this.groupMergeSheets.Items.Add(this.mergesheets_BeginMerge);
            this.groupMergeSheets.Label = "合并工作表";
            this.groupMergeSheets.Name = "groupMergeSheets";
            // 
            // mergesheets_HeadRowNum
            // 
            ribbonDropDownItemImpl1.Label = "无表头";
            ribbonDropDownItemImpl2.Label = "1";
            ribbonDropDownItemImpl3.Label = "2";
            ribbonDropDownItemImpl4.Label = "3";
            ribbonDropDownItemImpl5.Label = "4";
            ribbonDropDownItemImpl6.Label = "5";
            ribbonDropDownItemImpl7.Label = "6";
            ribbonDropDownItemImpl8.Label = "7";
            ribbonDropDownItemImpl9.Label = "8";
            ribbonDropDownItemImpl10.Label = "9";
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl1);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl2);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl3);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl4);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl5);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl6);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl7);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl8);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl9);
            this.mergesheets_HeadRowNum.Items.Add(ribbonDropDownItemImpl10);
            this.mergesheets_HeadRowNum.Label = "表头行数";
            this.mergesheets_HeadRowNum.Name = "mergesheets_HeadRowNum";
            // 
            // mergesheets_contentRowNum
            // 
            ribbonDropDownItemImpl11.Label = "不确定";
            ribbonDropDownItemImpl12.Label = "1";
            ribbonDropDownItemImpl13.Label = "2";
            ribbonDropDownItemImpl14.Label = "3";
            ribbonDropDownItemImpl15.Label = "4";
            ribbonDropDownItemImpl16.Label = "5";
            ribbonDropDownItemImpl17.Label = "6";
            ribbonDropDownItemImpl18.Label = "7";
            ribbonDropDownItemImpl19.Label = "8";
            ribbonDropDownItemImpl20.Label = "9";
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl11);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl12);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl13);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl14);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl15);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl16);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl17);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl18);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl19);
            this.mergesheets_contentRowNum.Items.Add(ribbonDropDownItemImpl20);
            this.mergesheets_contentRowNum.Label = "正表行数";
            this.mergesheets_contentRowNum.Name = "mergesheets_contentRowNum";
            // 
            // mergesheets_isFunctionEmbeded
            // 
            this.mergesheets_isFunctionEmbeded.Label = "含公式或函数";
            this.mergesheets_isFunctionEmbeded.Name = "mergesheets_isFunctionEmbeded";
            // 
            // convert
            // 
            this.convert.Items.Add(this.convert_sourceFormat);
            this.convert.Items.Add(this.convert_targetFormat);
            this.convert.Items.Add(this.buttonGroup1);
            this.convert.Label = "工作簿格式转换";
            this.convert.Name = "convert";
            // 
            // convert_sourceFormat
            // 
            ribbonDropDownItemImpl21.Label = "*.xlsx";
            ribbonDropDownItemImpl22.Label = "*.xlsb";
            ribbonDropDownItemImpl23.Label = "*.xlsm";
            ribbonDropDownItemImpl24.Label = "*.xls";
            ribbonDropDownItemImpl25.Label = "*.csv";
            ribbonDropDownItemImpl26.Label = "*.*";
            this.convert_sourceFormat.Items.Add(ribbonDropDownItemImpl21);
            this.convert_sourceFormat.Items.Add(ribbonDropDownItemImpl22);
            this.convert_sourceFormat.Items.Add(ribbonDropDownItemImpl23);
            this.convert_sourceFormat.Items.Add(ribbonDropDownItemImpl24);
            this.convert_sourceFormat.Items.Add(ribbonDropDownItemImpl25);
            this.convert_sourceFormat.Items.Add(ribbonDropDownItemImpl26);
            this.convert_sourceFormat.Label = "源格式";
            this.convert_sourceFormat.Name = "convert_sourceFormat";
            this.convert_sourceFormat.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.convert_sourceFormat_SelectionChanged);
            // 
            // convert_targetFormat
            // 
            ribbonDropDownItemImpl27.Label = ".xlsx";
            ribbonDropDownItemImpl27.OfficeImageId = "FileSaveAsExcelXlsx";
            ribbonDropDownItemImpl28.Label = ".xlsb";
            ribbonDropDownItemImpl28.OfficeImageId = "FileSaveAsExcelXlsb";
            ribbonDropDownItemImpl29.Label = ".xlsm";
            ribbonDropDownItemImpl29.OfficeImageId = "FileSaveAsExcelXlsxMacro";
            ribbonDropDownItemImpl30.Label = ".xls";
            ribbonDropDownItemImpl30.OfficeImageId = "FileSaveAsExcel97_2003";
            ribbonDropDownItemImpl31.Label = ".csv";
            ribbonDropDownItemImpl31.OfficeImageId = "CommaStyle";
            ribbonDropDownItemImpl32.Label = ".pdf";
            ribbonDropDownItemImpl32.OfficeImageId = "FileSaveAsPdfOrXps";
            this.convert_targetFormat.Items.Add(ribbonDropDownItemImpl27);
            this.convert_targetFormat.Items.Add(ribbonDropDownItemImpl28);
            this.convert_targetFormat.Items.Add(ribbonDropDownItemImpl29);
            this.convert_targetFormat.Items.Add(ribbonDropDownItemImpl30);
            this.convert_targetFormat.Items.Add(ribbonDropDownItemImpl31);
            this.convert_targetFormat.Items.Add(ribbonDropDownItemImpl32);
            this.convert_targetFormat.Label = "目标格式";
            this.convert_targetFormat.Name = "convert_targetFormat";
            this.convert_targetFormat.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.convert_targetFormat_SelectionChanged);
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.convert_Exchange);
            this.buttonGroup1.Items.Add(this.convert_Spreater);
            this.buttonGroup1.Items.Add(this.convert_BeginConvert);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // others
            // 
            this.others.Items.Add(this.others_DeleteOtherSheets);
            this.others.Items.Add(this.LookForFirstEmptyRow);
            this.others.Items.Add(this.others_ClrClipboard);
            this.others.Label = "闲杂工具";
            this.others.Name = "others";
            // 
            // dangerous_zone
            // 
            this.dangerous_zone.Items.Add(this.updateView);
            this.dangerous_zone.Items.Add(this.showAlert);
            this.dangerous_zone.Label = "危险区域";
            this.dangerous_zone.Name = "dangerous_zone";
            // 
            // updateView
            // 
            this.updateView.Checked = true;
            this.updateView.Label = "更新视图";
            this.updateView.Name = "updateView";
            this.updateView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateView_Click);
            // 
            // showAlert
            // 
            this.showAlert.Checked = true;
            this.showAlert.Label = "显示警告";
            this.showAlert.Name = "showAlert";
            this.showAlert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showAlert_Click);
            // 
            // support
            // 
            this.support.Items.Add(this.help_About);
            this.support.Label = "帮助和反馈";
            this.support.Name = "support";
            // 
            // mergebooks_BeginMerge
            // 
            this.mergebooks_BeginMerge.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mergebooks_BeginMerge.Label = "开始合并";
            this.mergebooks_BeginMerge.Name = "mergebooks_BeginMerge";
            this.mergebooks_BeginMerge.OfficeImageId = "PivotExportToExcel";
            this.mergebooks_BeginMerge.ShowImage = true;
            this.mergebooks_BeginMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mergebooks_BeginMerge_Click);
            // 
            // mergesheets_BeginMerge
            // 
            this.mergesheets_BeginMerge.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mergesheets_BeginMerge.Label = "开始合并";
            this.mergesheets_BeginMerge.Name = "mergesheets_BeginMerge";
            this.mergesheets_BeginMerge.OfficeImageId = "TableExcelSpreadsheetInsert";
            this.mergesheets_BeginMerge.ShowImage = true;
            this.mergesheets_BeginMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mergesheets_BeginMerge_Click);
            // 
            // convert_Exchange
            // 
            this.convert_Exchange.Label = "交换";
            this.convert_Exchange.Name = "convert_Exchange";
            this.convert_Exchange.OfficeImageId = "TabOrder";
            this.convert_Exchange.ShowImage = true;
            this.convert_Exchange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.convert_Exchange_Click);
            // 
            // convert_BeginConvert
            // 
            this.convert_BeginConvert.Label = "开始转换";
            this.convert_BeginConvert.Name = "convert_BeginConvert";
            this.convert_BeginConvert.OfficeImageId = "FileSaveAsOtherFormats";
            this.convert_BeginConvert.ShowImage = true;
            this.convert_BeginConvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.convert_BeginConvert_Click);
            // 
            // others_DeleteOtherSheets
            // 
            this.others_DeleteOtherSheets.Label = "删除多余工作表";
            this.others_DeleteOtherSheets.Name = "others_DeleteOtherSheets";
            this.others_DeleteOtherSheets.OfficeImageId = "SheetDelete";
            this.others_DeleteOtherSheets.ShowImage = true;
            this.others_DeleteOtherSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.others_DeleteOtherSheets_Click);
            // 
            // LookForFirstEmptyRow
            // 
            this.LookForFirstEmptyRow.Label = "查找首个空行";
            this.LookForFirstEmptyRow.Name = "LookForFirstEmptyRow";
            this.LookForFirstEmptyRow.OfficeImageId = "TableRowSelect";
            this.LookForFirstEmptyRow.ShowImage = true;
            this.LookForFirstEmptyRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.others_LookForFirstEmptyRow_Click);
            // 
            // others_ClrClipboard
            // 
            this.others_ClrClipboard.Label = "清除剪贴板";
            this.others_ClrClipboard.Name = "others_ClrClipboard";
            this.others_ClrClipboard.OfficeImageId = "Clear";
            this.others_ClrClipboard.ShowImage = true;
            this.others_ClrClipboard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.others_ClrClipboard_Click);
            // 
            // help_About
            // 
            this.help_About.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.help_About.Label = "关于和 bug反馈";
            this.help_About.Name = "help_About";
            this.help_About.OfficeImageId = "Info";
            this.help_About.ShowImage = true;
            this.help_About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.help_About_Click);
            // 
            // convert_Spreater
            // 
            this.convert_Spreater.Enabled = false;
            this.convert_Spreater.Label = "　 ";
            this.convert_Spreater.Name = "convert_Spreater";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupMergeBooks.ResumeLayout(false);
            this.groupMergeBooks.PerformLayout();
            this.groupMergeSheets.ResumeLayout(false);
            this.groupMergeSheets.PerformLayout();
            this.convert.ResumeLayout(false);
            this.convert.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.others.ResumeLayout(false);
            this.others.PerformLayout();
            this.dangerous_zone.ResumeLayout(false);
            this.dangerous_zone.PerformLayout();
            this.support.ResumeLayout(false);
            this.support.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMergeBooks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton mergebooks_BeginMerge;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMergeSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown mergesheets_HeadRowNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown mergesheets_contentRowNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox mergesheets_isFunctionEmbeded;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton mergesheets_BeginMerge;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup others;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox mergebooks_RequireNewBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox mergebooks_MergeAllSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton others_DeleteOtherSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LookForFirstEmptyRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup dangerous_zone;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox updateView;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox showAlert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton others_ClrClipboard;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup support;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton help_About;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox mergebooks_AIO;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup convert;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown convert_sourceFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown convert_targetFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton convert_BeginConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton convert_Exchange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton convert_Spreater;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
