using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WorkBookObjectModelDemo1
{
    public partial class Sheet4
    {
        private void Sheet4_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet4_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.btnR1C1Notation.Click += new System.EventHandler(this.btnR1C1Notation_Click);
            this.btnActiveCellApp.Click += new System.EventHandler(this.btnActiveCell_Click);
            this.btnSheetObjectCell.Click += new System.EventHandler(this.btnSheetObjectCell_Click);
            this.btnReadValue.Click += new System.EventHandler(this.btnReadValue_Click);
            this.btnSetValue.Click += new System.EventHandler(this.btnSetValue_Click);
            this.btnSetValueCont.Click += new System.EventHandler(this.btnSetValueCont_Click);
            this.btnLoopCells.Click += new System.EventHandler(this.btnLoopCells_Click);
            this.btnCellFormat.Click += new System.EventHandler(this.btnCellFormat_Click);
            this.btnFormatWithStyle.Click += new System.EventHandler(this.btnFormatWithStyle_Click);
            this.Startup += new System.EventHandler(this.Sheet4_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet4_Shutdown);

        }

        #endregion

        private void btnActiveCell_Click(object sender, EventArgs e)
        {
            Excel.Range rng =  Globals.ThisWorkbook.Application.ActiveCell;
            MessageBox.Show(string.Format("Current active cell address: {0}", rng.Address));
        }

        private void btnSheetObjectCell_Click(object sender, EventArgs e)
        {
            Excel.Range rng = Globals.Sheet4.Range["A1"] ;
            MessageBox.Show(string.Format("Object sheet reference address: {0}", rng.Address));
        }

        private void btnR1C1Notation_Click(object sender, EventArgs e)
        {
            Excel.Range rng = Globals.Sheet4.Cells[1, 1];
            MessageBox.Show(string.Format("R1C1 Notation object sheet reference address: {0}", rng.Address));
        }

        private void btnReadValue_Click(object sender, EventArgs e)
        {
            Excel.Range rng = Globals.Sheet4.Range["G17"];
            MessageBox.Show(string.Format("Object sheet reference address: {0}", rng.Value2));
        }

        private void btnSetValue_Click(object sender, EventArgs e)
        {
            Excel.Range rng = Globals.Sheet4.Range["H17"];
            rng.Value2 = "New value";
        }

        private void btnSetValueCont_Click(object sender, EventArgs e)
        {
            Excel.Range rng = Globals.Sheet4.Range["I17:J17,G18,H18:J18"];
            rng.Value2 = 54;
        }

        private void btnLoopCells_Click(object sender, EventArgs e)
        {
            Excel.Range rng = Globals.Sheet4.Range["G17:H17,I17:J17,G18,H18:J18"];
            
            string msg = "";
            foreach (Excel.Range item in rng)
            {
                msg += string.Format("Cell {0}: Value: {1} ", item.Address, item.Value2);
                msg += "\n";
            }
            MessageBox.Show(msg);
        }

        private void btnCellFormat_Click(object sender, EventArgs e)
        {
            Excel.Range rng = Globals.Sheet4.Range["R14"];
            
            rng.Font.Bold = true;
            rng.Font.Italic = true;
            rng.Font.Name = "Arial Black";
            rng.Borders.ColorIndex = 4; // color index is from 1 to 56
            rng.Interior.Color = Color.FromArgb(56, 75, 3);

            //format entire border
            rng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range rng2 = Globals.Sheet4.Range["R16"];
            //format bottom only
            rng2.Borders.Item[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDash;
            rng2.NumberFormat = "#,##0.00";
        }

        private void btnFormatWithStyle_Click(object sender, EventArgs e)
        {
            Excel.Style style = Globals.ThisWorkbook.Styles.Add("My style");
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.Font.Name = "Arial Black";
            style.Borders.Color = XlRgbColor.rgbBlack;
            style.Interior.Color = XlRgbColor.rgbBlueViolet;
            style.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range rng = Globals.Sheet4.Range["R18"];
            rng.Style = style;
        }
    }
}
