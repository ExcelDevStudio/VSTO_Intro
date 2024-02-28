using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel_PIA= Microsoft.Office.Interop.Excel;
using Excel_VSTO = Microsoft.Office.Tools.Excel;

namespace AddInObjectModelDemo1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnSetStatusMsg_Click(object sender, RibbonControlEventArgs e)
        {
            //Document solution
            //Globals.ThisWorkbook.Application.StatusBar = "Running process";

            Globals.ThisAddIn.Application.Application.StatusBar = "Running process";
        }

        private void btnClearStatusMsg_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Application.StatusBar = "";
        }

        private void btnCalculateSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);

            //Excel_PIA.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            //first get current to return it to same state at the end
            var calMode = sheet.Application.Calculation;

            //make calcuation mode manual
            sheet.Application.Calculation = Excel_PIA.XlCalculation.xlCalculationManual;

            sheet.EnableCalculation = false;
            sheet.EnableCalculation = true;
            sheet.CalculateMethod();

            //put back original state
            sheet.Application.Calculation = calMode;

            // ***  DOCUMENT LEVEL SOLUTION CODE ***
            ////first get current to return it to same state at the end
            //var calMode = Globals.Sheet1.Application.Calculation;

            ////make calcuation mode manuel
            //Globals.Sheet1.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            //Globals.Sheet1.EnableCalculation = false;
            //Globals.Sheet1.EnableCalculation = true;
            //Globals.Sheet1.CalculateMethod();

            ////put back original state
            //Globals.Sheet1.Application.Calculation = calMode;
        }

        private void btnReCalculateAll_Click(object sender, RibbonControlEventArgs e)
        {
            //For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.
            Globals.ThisAddIn.Application.CalculateFullRebuild();

            // ***  DOCUMENT LEVEL SOLUTION CODE ***
            //Globals.ThisWorkbook.Application.CalculateFullRebuild();
        }

        private void btnWorkSheetFunction_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.WorksheetFunction.Sum(75, 35);
            MessageBox.Show(result.ToString());

            // ***  DOCUMENT LEVEL SOLUTION CODE ***
            //var result = Globals.ThisWorkbook.Application.WorksheetFunction.Sum(75, 35);
        }

        private void btnAddWorkbook_Click(object sender, RibbonControlEventArgs e)
        {
            
            //PIA object
            Excel_PIA.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add();

            //VSTO Object
            //Excel_VSTO.Workbook wb = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.Workbooks.Add());

            MessageBox.Show(string.Format("New workbook add: {0}", wb.Name));
        }

        private void btnOpenWorkbook_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(@"C:\ExcelDevStudio\WorkbookDemo1.xlsx");

            MessageBox.Show(string.Format("Opened workbook: {0}", wb.Name));
        }

        private void btnGetActiveWorkbook_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Workbook wb = Globals.Factory.GetVstoObject( Globals.ThisAddIn.Application.ActiveWorkbook);

            MessageBox.Show(string.Format("Got reference to active workbook: {0}", wb.Name));
        }

        private void btnGetWorkbookName_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Workbook wb = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.Workbooks["WorkbookDemo1.xlsx"]);

            MessageBox.Show(string.Format("Got reference to workbook by name: {0}", wb.Name));
        }

        private void btnGetWorkbookIndex_Click(object sender, RibbonControlEventArgs e)
        {
            int wbCount = Globals.ThisAddIn.Application.Workbooks.Count;
            Excel_VSTO.Workbook wb = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.Workbooks[wbCount]);
            wb.Activate();

            MessageBox.Show(string.Format("Got reference to workbook by index: {0}", wb.Name));
        }

        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            wb.Save();

            MessageBox.Show(string.Format("Worbook saved: {0}", wb.Name));
        }

        private void btnSaveAs_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Workbook wb = Globals.ThisAddIn.Application.Workbooks["WorkbookDemo1.xlsx"];
            wb.SaveAs(@"C:\ExcelDevStudio\WorkbookDemo2.xlsx");

            MessageBox.Show(string.Format("Workbook saved as: {0}", wb.Name));
        }

        private void btnClose_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Workbook wb = Globals.ThisAddIn.Application.Workbooks["WorkbookDemo2.xlsx"];
            wb.Close();

            MessageBox.Show(string.Format("Closed workbook: {0}", wb.Name));
        }

        private void btnLoopWorkbooks_Click(object sender, RibbonControlEventArgs e)
        {
            string msg = "";

            // method 1
            foreach (Excel_PIA.Workbook wb in Globals.ThisAddIn.Application.Workbooks)
            {
                //convert to vsto workbook that has implenmation
                Excel_VSTO. Workbook wb_vsto = Globals.Factory.GetVstoObject(wb);
                msg += "Workbook Name: " + wb_vsto.Name;
                msg += "\n";
            }

            //// method 2
            //int count = Globals.ThisAddIn.Application.Workbooks.Count;
            //for (int i = 1; i <= count; i++)
            //{
            //    //convert to vsto workbook that has implenmation
            //    Excel_VSTO.Workbook wb = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.Workbooks[i]);

            //    //workbooks collection is 1 index based
            //    msg += "Workbook Name: " + wb.Name;
            //    msg += "\n";
            //}
            MessageBox.Show(msg);
        }

        private void btnAddSheet_Click(object sender, RibbonControlEventArgs e)
        {
            //method 1
            // will add to active workbook
            Globals.ThisAddIn.Application.Worksheets.Add();

            //method 2
            // will add to active workbook
            //Globals.ThisAddIn.Application.Sheets.Add();
        }

        private void btnActiveSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(  Globals.ThisAddIn.Application.ActiveSheet);
            sht.Activate();

            MessageBox.Show(string.Format("Got reference to active sheet: {0}", sht.Name));
        }

        private void btnGetSheetByName_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Worksheet sht = Globals.Factory.GetVstoObject( Globals.ThisAddIn.Application.Worksheets["Sheet1"]);
            sht.Activate();

            MessageBox.Show(string.Format("Got reference to sheet by name: {0}", sht.Name));
        }

        private void btnGetSheetByIndex_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.Worksheets[1]);
            sht.Activate();

            MessageBox.Show(string.Format("Got reference to sheet by index: {0}", sht.Name));
        }

        private void btnRenameSheet_Click(object sender, RibbonControlEventArgs e)
        {

            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject( Globals.ThisAddIn.Application.Worksheets.Add()); 
            sht.Activate();

            //rename to New Sheet
            sht.Name = "New Sheet";

            MessageBox.Show(string.Format("Added and rename sheet: {0}", sht.Name));
        }

        private void btnReOrderSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Worksheet sht = Globals.ThisAddIn.Application.Worksheets["New Sheet"];

            //move before first sheet 
            sht.Move(Globals.ThisAddIn.Application.Worksheets[1]);
        }

        private void btnDeleteSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Worksheet sht = Globals.ThisAddIn.Application.Worksheets["New Sheet"];
            string sheetName = sht.Name;
            sht.Delete();
            MessageBox.Show(string.Format("Deleted name: {0}", sheetName));
        }

        private void btnLoopSheets_Click(object sender, RibbonControlEventArgs e)
        {
            string msg = "";

            //method 1
            Excel_VSTO.Workbook wb = Globals.Factory.GetVstoObject( Globals.ThisAddIn.Application.ActiveWorkbook);
            foreach (Excel_PIA.Worksheet sht in wb.Worksheets)
            {
                msg += "Sheet Name: " + sht.Name;
                msg += "\n";
            }

            //method 2
            //int count = wb.Worksheets.Count;
            //for (int i = 0; i < count; i++)
            //{
            //    msg += "Sheet Name: " + wb.Worksheets[i].Name;
            //    msg += "\n";
            //}

            MessageBox.Show(msg);
        }

        private void btnActiveCellApp_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Range rng = Globals.ThisAddIn.Application.ActiveCell;
            MessageBox.Show(string.Format("Current active cell address: {0}", rng.Address));
        }

        private void btnSheetObjectCell_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);

            Excel_PIA.Range rng = sht.Range["C12"];
            MessageBox.Show(string.Format("Object sheet reference address: {0}", rng.Address));
        }

        private void btnR1C1Notation_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);

            //Get cell "C15"
            Excel_PIA.Range rng = sht.Cells[15, 3];
            MessageBox.Show(string.Format("R1C1 Notation object sheet reference address: {0}", rng.Address));
        }

        private void btnReadValue_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);
            Excel_PIA.Range rng = sht.Range["I9"];
            MessageBox.Show(string.Format("Object sheet reference address: {0}", rng.Value2));
        }

        private void btnSetValue_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);
            Excel_PIA.Range rng = sht.Range["I11"];
            rng.Value2 = "New value";
        }

        private void btnSetValueCont_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);
            Excel_PIA.Range rng = sht.Range["H14:I14,H15:I16,H18"];
            rng.Value2 = 54;
        }

        private void btnLoopCells_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);

            Excel_PIA.Range rng = sht.Range["M9:M12"];

            string msg = "";
            foreach (Excel_PIA.Range item in rng)
            {
                msg += string.Format("Cell {0}: Value: {1} ", item.Address, item.Value2);
                msg += "\n";
            }
            MessageBox.Show(msg);
        }

        private void btnCellFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);

            Excel_PIA.Range rng = sht.Range["R14"];

            rng.Font.Bold = true;
            rng.Font.Italic = true;
            rng.Font.Name = "Arial Black";
            rng.Borders.ColorIndex = 4; // color index is from 1 to 56
            rng.Interior.Color = Color.FromArgb(56, 75, 3);

            //format entire border
            rng.Borders.LineStyle = Excel_PIA.XlLineStyle.xlContinuous;

            Excel_PIA.Range rng2 = sht.Range["R16"];
            //format bottom only
            rng2.Borders.Item[Excel_PIA.XlBordersIndex.xlEdgeBottom].LineStyle = Excel_PIA.XlLineStyle.xlDash;
            rng2.NumberFormat = "#,##0.00";
        }

        private void btnFormatWithStyle_Click(object sender, RibbonControlEventArgs e)
        {
            Excel_PIA.Style style = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("My style");
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.Font.Name = "Arial Black";
            style.Borders.Color = Excel_PIA.XlRgbColor.rgbBlack;
            style.Interior.Color = Excel_PIA.XlRgbColor.rgbBlueViolet;
            style.Borders.LineStyle = Excel_PIA.XlLineStyle.xlContinuous;

            Excel_VSTO.Worksheet sht = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);

            Excel_PIA.Range rng = sht.Range["R18"];
            rng.Style = style;
        }

    }
}
