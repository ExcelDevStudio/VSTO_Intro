using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WorkBookObjectModelDemo1
{
    public partial class Sheet2
    {
        private void Sheet2_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.btnAddWorkbook.Click += new System.EventHandler(this.btnAddWorkbook_Click);
            this.btnGetActiveWorkbook.Click += new System.EventHandler(this.btnGetActiveWorkbook_Click);
            this.btnGetWorkbookName.Click += new System.EventHandler(this.btnGetWorkbookName_Click);
            this.btnGetWorkbookIndex.Click += new System.EventHandler(this.btnGetWorkbookIndex_Click);
            this.btnOpenWorkbook.Click += new System.EventHandler(this.btnOpenWorkbook_Click);
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            this.btnSaveAs.Click += new System.EventHandler(this.btnSaveAs_Click);
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            this.btnLoopWorkbooks.Click += new System.EventHandler(this.btnLoopWorkbooks_Click);
            this.Startup += new System.EventHandler(this.Sheet2_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet2_Shutdown);

        }


        #endregion

        private void btnAddWorkbook_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Add();

            MessageBox.Show(string.Format("New workbook add: {0}",  wb.Name));
        }

        private void btnOpenWorkbook_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Open(@"C:\ExcelDevStudio\WorkbookDemo1.xlsx");

            MessageBox.Show(string.Format("Opened workbook: {0}", wb.Name));
        }

        private void btnGetActiveWorkbook_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.ActiveWorkbook;
            
            MessageBox.Show(string.Format("Got reference to active workbook: {0}", wb.Name));
        }

        private void btnGetWorkbookName_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks["WorkbookDemo1.xlsx"];
            wb.Activate();

            MessageBox.Show(string.Format("Got reference to workbook by name: {0}", wb.Name));
        }

        private void btnGetWorkbookIndex_Click(object sender, EventArgs e)
        {
            int wbCount = Globals.ThisWorkbook.Application.Workbooks.Count;
            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks[wbCount];
            wb.Activate();

            MessageBox.Show(string.Format("Got reference to workbook by index: {0}", wb.Name));
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.ActiveWorkbook;
            wb.Save();

            MessageBox.Show(string.Format("Worbook saved: {0}", wb.Name));
        }

        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks["WorkbookDemo1.xlsx"];
            wb.SaveAs(@"C:\ExcelDevStudio\WorkbookDemo2.xlsx");

            MessageBox.Show(string.Format("Workbook saved as: {0}", wb.Name));
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks["WorkbookDemo2.xlsx"];
            var name = wb.Name;
            wb.Close();

            MessageBox.Show(string.Format("Workbook saved as: {0}", name));
        }

        private void btnLoopWorkbooks_Click(object sender, EventArgs e)
        {
            string msg = "";

            // method 1
            foreach (Excel.Workbook wb in Globals.ThisWorkbook.Application.Workbooks)
            {
                msg += "Workbook Name: " + wb.Name;
                msg += "\n";
            }

            //// method 2
            //int count = Globals.ThisWorkbook.Application.Workbooks.Count;
            //for (int i = 1; i <= count; i++)
            //{
            //    //workbooks collection is 1 index based
            //    msg += "Workbook Name: " + Globals.ThisWorkbook.Application.Workbooks[i].Name;
            //    msg += "\n";
            //}
            MessageBox.Show(msg);
        }
    }
}
