namespace Excel工具箱
{
    partial class RibbonUI : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonUI()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
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
            this.groupConvert = this.Factory.CreateRibbonGroup();
            this.convert_sourceFormat = this.Factory.CreateRibbonDropDown();
            this.convert_targetFormat = this.Factory.CreateRibbonDropDown();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.cellActions = this.Factory.CreateRibbonGroup();
            this.groupRename = this.Factory.CreateRibbonGroup();
            this.others = this.Factory.CreateRibbonGroup();
            this.dangerous_zone = this.Factory.CreateRibbonGroup();
            this.dangerzone_updateView = this.Factory.CreateRibbonCheckBox();
            this.dangerzone_showAlert = this.Factory.CreateRibbonCheckBox();
            this.support = this.Factory.CreateRibbonGroup();
            this.mergebooks_BeginMerge = this.Factory.CreateRibbonButton();
            this.mergesheets_BeginMerge = this.Factory.CreateRibbonButton();
            this.convert_Exchange = this.Factory.CreateRibbonButton();
            this.convert_BeginConvert = this.Factory.CreateRibbonButton();
            this.cellActions_ConvertToValue = this.Factory.CreateRibbonButton();
            this.cellActions_ConvertToString = this.Factory.CreateRibbonButton();
            this.cellActions_HighlightCurrentRC = this.Factory.CreateRibbonToggleButton();
            this.rename_RenameWorksheets = this.Factory.CreateRibbonButton();
            this.rename_SortSheets = this.Factory.CreateRibbonButton();
            this.rename_RenameWorkbooks = this.Factory.CreateRibbonButton();
            this.others_DeleteOtherSheets = this.Factory.CreateRibbonButton();
            this.LookForFirstEmptyRow = this.Factory.CreateRibbonButton();
            this.others_ClrClipboard = this.Factory.CreateRibbonButton();
            this.others_ConvertUI = this.Factory.CreateRibbonButton();
            this.dangerzone_tryFix = this.Factory.CreateRibbonButton();
            this.help_About = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupMergeBooks.SuspendLayout();
            this.groupMergeSheets.SuspendLayout();
            this.groupConvert.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.cellActions.SuspendLayout();
            this.groupRename.SuspendLayout();
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
            this.tab1.Groups.Add(this.groupConvert);
            this.tab1.Groups.Add(this.cellActions);
            this.tab1.Groups.Add(this.groupRename);
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
            // groupConvert
            // 
            this.groupConvert.DialogLauncher = ribbonDialogLauncherImpl1;
            this.groupConvert.Items.Add(this.convert_sourceFormat);
            this.groupConvert.Items.Add(this.convert_targetFormat);
            this.groupConvert.Items.Add(this.buttonGroup1);
            this.groupConvert.Label = "格式转换";
            this.groupConvert.Name = "groupConvert";
            this.groupConvert.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.groupConvert_DialogLauncherClick);
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
            ribbonDropDownItemImpl27.Tag = "51";
            ribbonDropDownItemImpl28.Label = ".xlsb";
            ribbonDropDownItemImpl28.OfficeImageId = "FileSaveAsExcelXlsb";
            ribbonDropDownItemImpl28.Tag = "50";
            ribbonDropDownItemImpl29.Label = ".xlsm";
            ribbonDropDownItemImpl29.OfficeImageId = "FileSaveAsExcelXlsxMacro";
            ribbonDropDownItemImpl29.Tag = "52";
            ribbonDropDownItemImpl30.Label = ".xls";
            ribbonDropDownItemImpl30.OfficeImageId = "FileSaveAsExcel97_2003";
            ribbonDropDownItemImpl30.Tag = "56";
            ribbonDropDownItemImpl31.Label = ".csv";
            ribbonDropDownItemImpl31.OfficeImageId = "CommaStyle";
            ribbonDropDownItemImpl31.Tag = "6";
            ribbonDropDownItemImpl32.Label = ".pdf";
            ribbonDropDownItemImpl32.OfficeImageId = "FileSaveAsPdfOrXps";
            ribbonDropDownItemImpl32.Tag = "0";
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
            this.buttonGroup1.Items.Add(this.convert_BeginConvert);
            this.buttonGroup1.Items.Add(this.others_ConvertUI);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // cellActions
            // 
            this.cellActions.Items.Add(this.cellActions_ConvertToValue);
            this.cellActions.Items.Add(this.cellActions_ConvertToString);
            this.cellActions.Items.Add(this.cellActions_HighlightCurrentRC);
            this.cellActions.Label = "单元格操作";
            this.cellActions.Name = "cellActions";
            // 
            // groupRename
            // 
            this.groupRename.Items.Add(this.rename_RenameWorksheets);
            this.groupRename.Items.Add(this.rename_SortSheets);
            this.groupRename.Items.Add(this.rename_RenameWorkbooks);
            this.groupRename.Label = "批量重命名";
            this.groupRename.Name = "groupRename";
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
            this.dangerous_zone.Items.Add(this.dangerzone_updateView);
            this.dangerous_zone.Items.Add(this.dangerzone_showAlert);
            this.dangerous_zone.Items.Add(this.dangerzone_tryFix);
            this.dangerous_zone.Label = "危险区域";
            this.dangerous_zone.Name = "dangerous_zone";
            this.dangerous_zone.Visible = false;
            // 
            // dangerzone_updateView
            // 
            this.dangerzone_updateView.Checked = true;
            this.dangerzone_updateView.Label = "更新视图";
            this.dangerzone_updateView.Name = "dangerzone_updateView";
            this.dangerzone_updateView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateView_Click);
            // 
            // dangerzone_showAlert
            // 
            this.dangerzone_showAlert.Checked = true;
            this.dangerzone_showAlert.Label = "显示警告";
            this.dangerzone_showAlert.Name = "dangerzone_showAlert";
            this.dangerzone_showAlert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showAlert_Click);
            // 
            // support
            // 
            this.support.Items.Add(this.help_About);
            this.support.Label = "帮助和反馈";
            this.support.Name = "support";
            this.support.Visible = false;
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
            // cellActions_ConvertToValue
            // 
            this.cellActions_ConvertToValue.Label = "转化为数值";
            this.cellActions_ConvertToValue.Name = "cellActions_ConvertToValue";
            this.cellActions_ConvertToValue.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cellActions_ConvertToValue_Click);
            // 
            // cellActions_ConvertToString
            // 
            this.cellActions_ConvertToString.Label = "转换为文本";
            this.cellActions_ConvertToString.Name = "cellActions_ConvertToString";
            this.cellActions_ConvertToString.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cellActions_ConvertToString_Click);
            // 
            // cellActions_HighlightCurrentRC
            // 
            this.cellActions_HighlightCurrentRC.Label = "高亮当前行列";
            this.cellActions_HighlightCurrentRC.Name = "cellActions_HighlightCurrentRC";
            this.cellActions_HighlightCurrentRC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cellActions_HighlightCurrentRC_Click);
            // 
            // rename_RenameWorksheets
            // 
            this.rename_RenameWorksheets.Label = "重命名工作表";
            this.rename_RenameWorksheets.Name = "rename_RenameWorksheets";
            this.rename_RenameWorksheets.OfficeImageId = "DatasheetColumnRename";
            this.rename_RenameWorksheets.ShowImage = true;
            this.rename_RenameWorksheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rename_RenameWorksheets_Click);
            // 
            // rename_SortSheets
            // 
            this.rename_SortSheets.Label = "排序工作表";
            this.rename_SortSheets.Name = "rename_SortSheets";
            this.rename_SortSheets.OfficeImageId = "SortAscendingExcel";
            this.rename_SortSheets.ShowImage = true;
            this.rename_SortSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rename_SortSheets_Click);
            // 
            // rename_RenameWorkbooks
            // 
            this.rename_RenameWorkbooks.Label = "重命名工作簿";
            this.rename_RenameWorkbooks.Name = "rename_RenameWorkbooks";
            this.rename_RenameWorkbooks.OfficeImageId = "UpgradeWorkbook";
            this.rename_RenameWorkbooks.ShowImage = true;
            this.rename_RenameWorkbooks.Visible = false;
            this.rename_RenameWorkbooks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rename_RenameWorkbooks_Click);
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
            // others_ConvertUI
            // 
            this.others_ConvertUI.Label = "…";
            this.others_ConvertUI.Name = "others_ConvertUI";
            this.others_ConvertUI.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.others_ConvertUI_Click);
            // 
            // dangerzone_tryFix
            // 
            this.dangerzone_tryFix.Label = "修复视图问题";
            this.dangerzone_tryFix.Name = "dangerzone_tryFix";
            this.dangerzone_tryFix.OfficeImageId = "TableOfContentsUpdate";
            this.dangerzone_tryFix.ShowImage = true;
            this.dangerzone_tryFix.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dangerzone_tryFix_Click);
            // 
            // help_About
            // 
            this.help_About.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.help_About.Label = "帮助和 bug反馈";
            this.help_About.Name = "help_About";
            this.help_About.OfficeImageId = "Info";
            this.help_About.ShowImage = true;
            this.help_About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.help_About_Click);
            // 
            // RibbonUI
            // 
            this.Name = "RibbonUI";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonUI_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupMergeBooks.ResumeLayout(false);
            this.groupMergeBooks.PerformLayout();
            this.groupMergeSheets.ResumeLayout(false);
            this.groupMergeSheets.PerformLayout();
            this.groupConvert.ResumeLayout(false);
            this.groupConvert.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.cellActions.ResumeLayout(false);
            this.cellActions.PerformLayout();
            this.groupRename.ResumeLayout(false);
            this.groupRename.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox dangerzone_updateView;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox dangerzone_showAlert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton others_ClrClipboard;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup support;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton help_About;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox mergebooks_AIO;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown convert_sourceFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown convert_targetFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton convert_BeginConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton convert_Exchange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton dangerzone_tryFix;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupRename;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rename_RenameWorksheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rename_RenameWorkbooks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rename_SortSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup cellActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cellActions_ConvertToValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cellActions_ConvertToString;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton cellActions_HighlightCurrentRC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton others_ConvertUI;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonUI Ribbon1
        {
            get { return this.GetRibbon<RibbonUI>(); }
        }
    }
}
