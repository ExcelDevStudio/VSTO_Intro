namespace AddInObjectModelDemo1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.RibbonObjectModel = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.btnSetStatusMsg = this.Factory.CreateRibbonButton();
            this.btnClearStatusMsg = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.btnCalculateSheet = this.Factory.CreateRibbonButton();
            this.btnReCalculateAll = this.Factory.CreateRibbonButton();
            this.menu3 = this.Factory.CreateRibbonMenu();
            this.btnWorkSheetFunction = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.menu4 = this.Factory.CreateRibbonMenu();
            this.btnAddWorkbook = this.Factory.CreateRibbonButton();
            this.btnOpenWorkbook = this.Factory.CreateRibbonButton();
            this.menu5 = this.Factory.CreateRibbonMenu();
            this.btnGetActiveWorkbook = this.Factory.CreateRibbonButton();
            this.btnGetWorkbookName = this.Factory.CreateRibbonButton();
            this.btnGetWorkbookIndex = this.Factory.CreateRibbonButton();
            this.menu6 = this.Factory.CreateRibbonMenu();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.btnSaveAs = this.Factory.CreateRibbonButton();
            this.btnClose = this.Factory.CreateRibbonButton();
            this.menu7 = this.Factory.CreateRibbonMenu();
            this.btnLoopWorkbooks = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.menu8 = this.Factory.CreateRibbonMenu();
            this.btnAddSheet = this.Factory.CreateRibbonButton();
            this.menu9 = this.Factory.CreateRibbonMenu();
            this.btnActiveSheet = this.Factory.CreateRibbonButton();
            this.btnGetSheetByName = this.Factory.CreateRibbonButton();
            this.btnGetSheetByIndex = this.Factory.CreateRibbonButton();
            this.menu10 = this.Factory.CreateRibbonMenu();
            this.btnRenameSheet = this.Factory.CreateRibbonButton();
            this.btnReOrderSheet = this.Factory.CreateRibbonButton();
            this.btnDeleteSheet = this.Factory.CreateRibbonButton();
            this.menu11 = this.Factory.CreateRibbonMenu();
            this.btnLoopSheets = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.menu12 = this.Factory.CreateRibbonMenu();
            this.btnActiveCellApp = this.Factory.CreateRibbonButton();
            this.btnSheetObjectCell = this.Factory.CreateRibbonButton();
            this.btnR1C1Notation = this.Factory.CreateRibbonButton();
            this.menu13 = this.Factory.CreateRibbonMenu();
            this.btnReadValue = this.Factory.CreateRibbonButton();
            this.btnSetValue = this.Factory.CreateRibbonButton();
            this.btnSetValueCont = this.Factory.CreateRibbonButton();
            this.menu14 = this.Factory.CreateRibbonMenu();
            this.btnLoopCells = this.Factory.CreateRibbonButton();
            this.menu15 = this.Factory.CreateRibbonMenu();
            this.btnCellFormat = this.Factory.CreateRibbonButton();
            this.btnFormatWithStyle = this.Factory.CreateRibbonButton();
            this.RibbonObjectModel.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // RibbonObjectModel
            // 
            this.RibbonObjectModel.Groups.Add(this.group1);
            this.RibbonObjectModel.Groups.Add(this.group2);
            this.RibbonObjectModel.Groups.Add(this.group3);
            this.RibbonObjectModel.Groups.Add(this.group4);
            this.RibbonObjectModel.Label = "Object Model";
            this.RibbonObjectModel.Name = "RibbonObjectModel";
            // 
            // group1
            // 
            this.group1.Items.Add(this.menu1);
            this.group1.Items.Add(this.menu2);
            this.group1.Items.Add(this.menu3);
            this.group1.Label = "Application";
            this.group1.Name = "group1";
            // 
            // menu1
            // 
            this.menu1.Items.Add(this.btnSetStatusMsg);
            this.menu1.Items.Add(this.btnClearStatusMsg);
            this.menu1.Label = "Example 1";
            this.menu1.Name = "menu1";
            // 
            // btnSetStatusMsg
            // 
            this.btnSetStatusMsg.Label = "Set Status Message ";
            this.btnSetStatusMsg.Name = "btnSetStatusMsg";
            this.btnSetStatusMsg.ShowImage = true;
            this.btnSetStatusMsg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetStatusMsg_Click);
            // 
            // btnClearStatusMsg
            // 
            this.btnClearStatusMsg.Label = "Clear Status Message ";
            this.btnClearStatusMsg.Name = "btnClearStatusMsg";
            this.btnClearStatusMsg.ShowImage = true;
            this.btnClearStatusMsg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearStatusMsg_Click);
            // 
            // menu2
            // 
            this.menu2.Items.Add(this.btnCalculateSheet);
            this.menu2.Items.Add(this.btnReCalculateAll);
            this.menu2.Label = "Example 2";
            this.menu2.Name = "menu2";
            // 
            // btnCalculateSheet
            // 
            this.btnCalculateSheet.Label = "Recalculate Sheet ";
            this.btnCalculateSheet.Name = "btnCalculateSheet";
            this.btnCalculateSheet.ShowImage = true;
            this.btnCalculateSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalculateSheet_Click);
            // 
            // btnReCalculateAll
            // 
            this.btnReCalculateAll.Label = "Recalculate All";
            this.btnReCalculateAll.Name = "btnReCalculateAll";
            this.btnReCalculateAll.ShowImage = true;
            this.btnReCalculateAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReCalculateAll_Click);
            // 
            // menu3
            // 
            this.menu3.Items.Add(this.btnWorkSheetFunction);
            this.menu3.Label = "Example 3";
            this.menu3.Name = "menu3";
            // 
            // btnWorkSheetFunction
            // 
            this.btnWorkSheetFunction.Label = "WorksheetFunction";
            this.btnWorkSheetFunction.Name = "btnWorkSheetFunction";
            this.btnWorkSheetFunction.ShowImage = true;
            this.btnWorkSheetFunction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWorkSheetFunction_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.menu4);
            this.group2.Items.Add(this.menu5);
            this.group2.Items.Add(this.menu6);
            this.group2.Items.Add(this.menu7);
            this.group2.Label = "Workbook";
            this.group2.Name = "group2";
            // 
            // menu4
            // 
            this.menu4.Items.Add(this.btnAddWorkbook);
            this.menu4.Items.Add(this.btnOpenWorkbook);
            this.menu4.Label = "Example 1";
            this.menu4.Name = "menu4";
            // 
            // btnAddWorkbook
            // 
            this.btnAddWorkbook.Label = "Add Workbook";
            this.btnAddWorkbook.Name = "btnAddWorkbook";
            this.btnAddWorkbook.ShowImage = true;
            this.btnAddWorkbook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddWorkbook_Click);
            // 
            // btnOpenWorkbook
            // 
            this.btnOpenWorkbook.Label = "Open Workbook";
            this.btnOpenWorkbook.Name = "btnOpenWorkbook";
            this.btnOpenWorkbook.ShowImage = true;
            this.btnOpenWorkbook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenWorkbook_Click);
            // 
            // menu5
            // 
            this.menu5.Items.Add(this.btnGetActiveWorkbook);
            this.menu5.Items.Add(this.btnGetWorkbookName);
            this.menu5.Items.Add(this.btnGetWorkbookIndex);
            this.menu5.Label = "Example 2";
            this.menu5.Name = "menu5";
            // 
            // btnGetActiveWorkbook
            // 
            this.btnGetActiveWorkbook.Label = "Active Workbook";
            this.btnGetActiveWorkbook.Name = "btnGetActiveWorkbook";
            this.btnGetActiveWorkbook.ShowImage = true;
            this.btnGetActiveWorkbook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetActiveWorkbook_Click);
            // 
            // btnGetWorkbookName
            // 
            this.btnGetWorkbookName.Label = "By Workbook Name";
            this.btnGetWorkbookName.Name = "btnGetWorkbookName";
            this.btnGetWorkbookName.ShowImage = true;
            this.btnGetWorkbookName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetWorkbookName_Click);
            // 
            // btnGetWorkbookIndex
            // 
            this.btnGetWorkbookIndex.Label = "By Workbook Index";
            this.btnGetWorkbookIndex.Name = "btnGetWorkbookIndex";
            this.btnGetWorkbookIndex.ShowImage = true;
            this.btnGetWorkbookIndex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetWorkbookIndex_Click);
            // 
            // menu6
            // 
            this.menu6.Items.Add(this.btnSave);
            this.menu6.Items.Add(this.btnSaveAs);
            this.menu6.Items.Add(this.btnClose);
            this.menu6.Label = "Example 3";
            this.menu6.Name = "menu6";
            // 
            // btnSave
            // 
            this.btnSave.Label = "Save";
            this.btnSave.Name = "btnSave";
            this.btnSave.ShowImage = true;
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.Label = "Save As";
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.ShowImage = true;
            this.btnSaveAs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveAs_Click);
            // 
            // btnClose
            // 
            this.btnClose.Label = "Close";
            this.btnClose.Name = "btnClose";
            this.btnClose.ShowImage = true;
            this.btnClose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClose_Click);
            // 
            // menu7
            // 
            this.menu7.Items.Add(this.btnLoopWorkbooks);
            this.menu7.Label = "Example 4";
            this.menu7.Name = "menu7";
            // 
            // btnLoopWorkbooks
            // 
            this.btnLoopWorkbooks.Label = "Loop Open Workbooks";
            this.btnLoopWorkbooks.Name = "btnLoopWorkbooks";
            this.btnLoopWorkbooks.ShowImage = true;
            this.btnLoopWorkbooks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoopWorkbooks_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.menu8);
            this.group3.Items.Add(this.menu9);
            this.group3.Items.Add(this.menu10);
            this.group3.Items.Add(this.menu11);
            this.group3.Label = "Worksheet";
            this.group3.Name = "group3";
            // 
            // menu8
            // 
            this.menu8.Items.Add(this.btnAddSheet);
            this.menu8.Label = "Example 1";
            this.menu8.Name = "menu8";
            // 
            // btnAddSheet
            // 
            this.btnAddSheet.Label = "Add Worksheet";
            this.btnAddSheet.Name = "btnAddSheet";
            this.btnAddSheet.ShowImage = true;
            this.btnAddSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddSheet_Click);
            // 
            // menu9
            // 
            this.menu9.Items.Add(this.btnActiveSheet);
            this.menu9.Items.Add(this.btnGetSheetByName);
            this.menu9.Items.Add(this.btnGetSheetByIndex);
            this.menu9.Label = "Example 2";
            this.menu9.Name = "menu9";
            // 
            // btnActiveSheet
            // 
            this.btnActiveSheet.Label = "Active Worksheet";
            this.btnActiveSheet.Name = "btnActiveSheet";
            this.btnActiveSheet.ShowImage = true;
            this.btnActiveSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActiveSheet_Click);
            // 
            // btnGetSheetByName
            // 
            this.btnGetSheetByName.Label = "By Sheet Name";
            this.btnGetSheetByName.Name = "btnGetSheetByName";
            this.btnGetSheetByName.ShowImage = true;
            this.btnGetSheetByName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetSheetByName_Click);
            // 
            // btnGetSheetByIndex
            // 
            this.btnGetSheetByIndex.Label = "By Sheet Index";
            this.btnGetSheetByIndex.Name = "btnGetSheetByIndex";
            this.btnGetSheetByIndex.ShowImage = true;
            this.btnGetSheetByIndex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetSheetByIndex_Click);
            // 
            // menu10
            // 
            this.menu10.Items.Add(this.btnRenameSheet);
            this.menu10.Items.Add(this.btnReOrderSheet);
            this.menu10.Items.Add(this.btnDeleteSheet);
            this.menu10.Label = "Example 3";
            this.menu10.Name = "menu10";
            // 
            // btnRenameSheet
            // 
            this.btnRenameSheet.Label = "Rename Sheet";
            this.btnRenameSheet.Name = "btnRenameSheet";
            this.btnRenameSheet.ShowImage = true;
            this.btnRenameSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRenameSheet_Click);
            // 
            // btnReOrderSheet
            // 
            this.btnReOrderSheet.Label = "Re-Order Sheet";
            this.btnReOrderSheet.Name = "btnReOrderSheet";
            this.btnReOrderSheet.ShowImage = true;
            this.btnReOrderSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReOrderSheet_Click);
            // 
            // btnDeleteSheet
            // 
            this.btnDeleteSheet.Label = "Delete Sheet";
            this.btnDeleteSheet.Name = "btnDeleteSheet";
            this.btnDeleteSheet.ShowImage = true;
            this.btnDeleteSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteSheet_Click);
            // 
            // menu11
            // 
            this.menu11.Items.Add(this.btnLoopSheets);
            this.menu11.Label = "Example 4";
            this.menu11.Name = "menu11";
            // 
            // btnLoopSheets
            // 
            this.btnLoopSheets.Label = "Loop Sheets";
            this.btnLoopSheets.Name = "btnLoopSheets";
            this.btnLoopSheets.ShowImage = true;
            this.btnLoopSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoopSheets_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.menu12);
            this.group4.Items.Add(this.menu13);
            this.group4.Items.Add(this.menu14);
            this.group4.Items.Add(this.menu15);
            this.group4.Label = "Range";
            this.group4.Name = "group4";
            // 
            // menu12
            // 
            this.menu12.Items.Add(this.btnActiveCellApp);
            this.menu12.Items.Add(this.btnSheetObjectCell);
            this.menu12.Items.Add(this.btnR1C1Notation);
            this.menu12.Label = "Example 1";
            this.menu12.Name = "menu12";
            // 
            // btnActiveCellApp
            // 
            this.btnActiveCellApp.Label = "Active Cell";
            this.btnActiveCellApp.Name = "btnActiveCellApp";
            this.btnActiveCellApp.ShowImage = true;
            this.btnActiveCellApp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActiveCellApp_Click);
            // 
            // btnSheetObjectCell
            // 
            this.btnSheetObjectCell.Label = "Sheet Object Cell";
            this.btnSheetObjectCell.Name = "btnSheetObjectCell";
            this.btnSheetObjectCell.ShowImage = true;
            this.btnSheetObjectCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSheetObjectCell_Click);
            // 
            // btnR1C1Notation
            // 
            this.btnR1C1Notation.Label = "R1C1 Notation";
            this.btnR1C1Notation.Name = "btnR1C1Notation";
            this.btnR1C1Notation.ShowImage = true;
            this.btnR1C1Notation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnR1C1Notation_Click);
            // 
            // menu13
            // 
            this.menu13.Items.Add(this.btnReadValue);
            this.menu13.Items.Add(this.btnSetValue);
            this.menu13.Items.Add(this.btnSetValueCont);
            this.menu13.Label = "Example 2";
            this.menu13.Name = "menu13";
            // 
            // btnReadValue
            // 
            this.btnReadValue.Label = "Read Value";
            this.btnReadValue.Name = "btnReadValue";
            this.btnReadValue.ShowImage = true;
            this.btnReadValue.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadValue_Click);
            // 
            // btnSetValue
            // 
            this.btnSetValue.Label = "Set Value";
            this.btnSetValue.Name = "btnSetValue";
            this.btnSetValue.ShowImage = true;
            this.btnSetValue.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetValue_Click);
            // 
            // btnSetValueCont
            // 
            this.btnSetValueCont.Label = "Set Values Contiguous";
            this.btnSetValueCont.Name = "btnSetValueCont";
            this.btnSetValueCont.ShowImage = true;
            this.btnSetValueCont.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetValueCont_Click);
            // 
            // menu14
            // 
            this.menu14.Items.Add(this.btnLoopCells);
            this.menu14.Label = "Example 3";
            this.menu14.Name = "menu14";
            // 
            // btnLoopCells
            // 
            this.btnLoopCells.Label = "Loop Cells";
            this.btnLoopCells.Name = "btnLoopCells";
            this.btnLoopCells.ShowImage = true;
            this.btnLoopCells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoopCells_Click);
            // 
            // menu15
            // 
            this.menu15.Items.Add(this.btnCellFormat);
            this.menu15.Items.Add(this.btnFormatWithStyle);
            this.menu15.Label = "Example 4";
            this.menu15.Name = "menu15";
            // 
            // btnCellFormat
            // 
            this.btnCellFormat.Label = "Cell Format";
            this.btnCellFormat.Name = "btnCellFormat";
            this.btnCellFormat.ShowImage = true;
            this.btnCellFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCellFormat_Click);
            // 
            // btnFormatWithStyle
            // 
            this.btnFormatWithStyle.Label = "Using Style";
            this.btnFormatWithStyle.Name = "btnFormatWithStyle";
            this.btnFormatWithStyle.ShowImage = true;
            this.btnFormatWithStyle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatWithStyle_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.RibbonObjectModel);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.RibbonObjectModel.ResumeLayout(false);
            this.RibbonObjectModel.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab RibbonObjectModel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearStatusMsg;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetStatusMsg;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalculateSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReCalculateAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWorkSheetFunction;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddWorkbook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenWorkbook;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetActiveWorkbook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetWorkbookName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetWorkbookIndex;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveAs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClose;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoopWorkbooks;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActiveSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetSheetByName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetSheetByIndex;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRenameSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReOrderSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoopSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActiveCellApp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSheetObjectCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnR1C1Notation;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetValueCont;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoopCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu15;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCellFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatWithStyle;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
