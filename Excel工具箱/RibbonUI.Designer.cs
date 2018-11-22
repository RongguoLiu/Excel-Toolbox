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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupMergeBooks = this.Factory.CreateRibbonGroup();
            this.mergebooks_RequireNewBook = this.Factory.CreateRibbonCheckBox();
            this.mergebooks_MergeAllSheets = this.Factory.CreateRibbonCheckBox();
            this.mergebooks_AIO = this.Factory.CreateRibbonCheckBox();
            this.mergebooks_BeginMerge = this.Factory.CreateRibbonButton();
            this.groupMergeSheets = this.Factory.CreateRibbonGroup();
            this.mergesheets_HeadRowNum = this.Factory.CreateRibbonDropDown();
            this.mergesheets_contentRowNum = this.Factory.CreateRibbonDropDown();
            this.mergesheets_isFunctionEmbeded = this.Factory.CreateRibbonCheckBox();
            this.mergesheets_BeginMerge = this.Factory.CreateRibbonButton();
            this.others = this.Factory.CreateRibbonGroup();
            this.others_DeleteOtherSheets = this.Factory.CreateRibbonButton();
            this.LookForFirstEmptyRow = this.Factory.CreateRibbonButton();
            this.others_ClrClipboard = this.Factory.CreateRibbonButton();
            this.dangerous_zone = this.Factory.CreateRibbonGroup();
            this.updateView = this.Factory.CreateRibbonCheckBox();
            this.showAlert = this.Factory.CreateRibbonCheckBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.help_About = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupMergeBooks.SuspendLayout();
            this.groupMergeSheets.SuspendLayout();
            this.others.SuspendLayout();
            this.dangerous_zone.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupMergeBooks);
            this.tab1.Groups.Add(this.groupMergeSheets);
            this.tab1.Groups.Add(this.others);
            this.tab1.Groups.Add(this.dangerous_zone);
            this.tab1.Groups.Add(this.group1);
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
            // mergebooks_BeginMerge
            // 
            this.mergebooks_BeginMerge.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mergebooks_BeginMerge.Label = "开始合并";
            this.mergebooks_BeginMerge.Name = "mergebooks_BeginMerge";
            this.mergebooks_BeginMerge.OfficeImageId = "PivotExportToExcel";
            this.mergebooks_BeginMerge.ShowImage = true;
            this.mergebooks_BeginMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mergebooks_BeginMerge_Click);
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
            // mergesheets_BeginMerge
            // 
            this.mergesheets_BeginMerge.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mergesheets_BeginMerge.Label = "开始合并";
            this.mergesheets_BeginMerge.Name = "mergesheets_BeginMerge";
            this.mergesheets_BeginMerge.OfficeImageId = "TableExcelSpreadsheetInsert";
            this.mergesheets_BeginMerge.ShowImage = true;
            this.mergesheets_BeginMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mergesheets_BeginMerge_Click);
            // 
            // others
            // 
            this.others.Items.Add(this.others_DeleteOtherSheets);
            this.others.Items.Add(this.LookForFirstEmptyRow);
            this.others.Items.Add(this.others_ClrClipboard);
            this.others.Label = "闲杂工具";
            this.others.Name = "others";
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
            // group1
            // 
            this.group1.Items.Add(this.help_About);
            this.group1.Label = "帮助和反馈";
            this.group1.Name = "group1";
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
            this.others.ResumeLayout(false);
            this.others.PerformLayout();
            this.dangerous_zone.ResumeLayout(false);
            this.dangerous_zone.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton help_About;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox mergebooks_AIO;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
