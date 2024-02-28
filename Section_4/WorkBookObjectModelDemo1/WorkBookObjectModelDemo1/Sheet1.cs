using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WorkBookObjectModelDemo1
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {

            Globals.ThisWorkbook.SheetActivate += Sheet1_Activaste;

            ////turn off and on screen updating
            //Globals.ThisWorkbook.Application.ScreenUpdating = false;
            //Globals.ThisWorkbook.Application.ScreenUpdating = true;

            //// check available memory for excel application
            //var memoryFree = Globals.ThisWorkbook.Application.MemoryFree;

            ////display 
            //Globals.ThisWorkbook.Application.DisplayScrollBars = true;
            //Globals.ThisWorkbook.Application.DisplayFormulaBar = true;
            //Globals.ThisWorkbook.Application.ActiveWindow.DisplayWorkbookTabs = true;

        }

        private void Sheet1_Activaste(object Sh)
        {
            Globals.ThisWorkbook.Application.ActiveWindow.Zoom = 100;
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.btnSetStatusMsg.Click += new System.EventHandler(this.btnSetStatusMsg_Click);
            this.btnClearStatusMsg.Click += new System.EventHandler(this.btnClearStatusMsg_Click);
            this.btnWorkSheetFunction.Click += new System.EventHandler(this.btnWorkSheetFunction_Click);
            this.btnCalculateSheet.Click += new System.EventHandler(this.btnCalculateSheet_Click);
            this.btnReCalculateAll.Click += new System.EventHandler(this.btnReCalculateAll_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void btnSetStatusMsg_Click(object sender, EventArgs e)
        {
            Globals.ThisWorkbook.Application.StatusBar = "Running process";
        }

        private void btnClearStatusMsg_Click(object sender, EventArgs e)
        {
            Globals.ThisWorkbook.Application.StatusBar = "";
        }

        private void btnCalculateSheet_Click(object sender, EventArgs e)
        {
            //first get current to return it to same state at the end
            var calMode = Globals.Sheet1.Application.Calculation;

            //make calcuation mode manuel
            Globals.Sheet1.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            Globals.Sheet1.EnableCalculation = false;
            Globals.Sheet1.EnableCalculation = true;
            Globals.Sheet1.CalculateMethod();

            //put back original state
            Globals.Sheet1.Application.Calculation = calMode;
        }

        private void btnReCalculateAll_Click(object sender, EventArgs e)
        {
            //For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.
            Globals.ThisWorkbook.Application.CalculateFullRebuild();
        }

        private void btnWorkSheetFunction_Click(object sender, EventArgs e)
        {
            var result = Globals.ThisWorkbook.Application.WorksheetFunction.Sum(75, 35);
            MessageBox.Show( result.ToString());
        }
    }
}
