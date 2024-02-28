using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WorkBookObjectModelDemo1
{
    public partial class Sheet3
    {
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.btnAddSheet.Click += new System.EventHandler(this.btnAddSheet_Click);
            this.btnGetActiveSheet.Click += new System.EventHandler(this.btnActiveSheet_Click);
            this.btnGetSheetByName.Click += new System.EventHandler(this.btnGetSheetByName_Click);
            this.btnGetSheetByIndex.Click += new System.EventHandler(this.btnGetSheetByIndex_Click);
            this.btnRenameSheet.Click += new System.EventHandler(this.btnRenameSheet_Click);
            this.btnReOrderSheet.Click += new System.EventHandler(this.btnReOrderSheet_Click);
            this.btnDeleteSheet.Click += new System.EventHandler(this.btnDeleteSheet_Click);
            this.btnLoopSheets.Click += new System.EventHandler(this.btnLoopSheets_Click);
            this.Startup += new System.EventHandler(this.Sheet3_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet3_Shutdown);

        }

        #endregion

        private void btnAddSheet_Click(object sender, EventArgs e)
        {
            //method 1
            Excel.Worksheet sht =  Globals.ThisWorkbook.Worksheets.Add();

            //method 2
            //Excel.Worksheet sht = Globals.ThisWorkbook.Sheets.Add();

            MessageBox.Show(string.Format("Added new sheet: {0}", sht.Name));
        }

        private void btnActiveSheet_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sht = Globals.ThisWorkbook.Application.ActiveSheet;
            sht.Activate();

            MessageBox.Show(string.Format("Got reference to active sheet: {0}", sht.Name));
        }

        private void btnGetSheetByName_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sht = Globals.ThisWorkbook.Application.Worksheets["Application"];
            sht.Activate();

            MessageBox.Show(string.Format("Got reference to sheet by name: {0}", sht.Name));
        }

        private void btnGetSheetByIndex_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sht = Globals.ThisWorkbook.Application.Worksheets[1];
            sht.Activate();

            MessageBox.Show(string.Format("Got reference to sheet by index: {0}", sht.Name));
        }

        private void btnRenameSheet_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sht = Globals.ThisWorkbook.Worksheets.Add(); //as Excel.Worksheet;
            sht.Activate();

            //rename to New Sheet
            sht.Name = "New Sheet";

            MessageBox.Show(string.Format("Added and rename sheet: {0}", sht.Name));
        }

        private void btnReOrderSheet_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sht = Globals.ThisWorkbook.Application.Worksheets["New Sheet"];

            //move before first sheet 
            sht.Move(Globals.ThisWorkbook.Application.Worksheets[1]);
        }

        private void btnDeleteSheet_Click(object sender, EventArgs e)
        {

            Excel.Worksheet sht = Globals.ThisWorkbook.Application.Worksheets["New Sheet"];
            string sheetName = sht.Name;

            sht.Delete();

            MessageBox.Show(string.Format("Deleted name: {0}", sheetName));
        }

        private void btnLoopSheets_Click(object sender, EventArgs e)
        {
            string msg = "";
            //method 1
            foreach ( Excel.Worksheet sht in Globals.ThisWorkbook.Worksheets)
            {
                msg += "Sheet Name: " + sht.Name;
                msg += "\n";
            }

            //method 2
            int count = Globals.ThisWorkbook.Worksheets.Count;
            for (int i = 0; i < count; i++)
            {
                msg += "Sheet Name: " + Globals.ThisWorkbook.Worksheets[i].Name;
                msg += "\n";
            }


            MessageBox.Show(msg);
        }
    }
}
