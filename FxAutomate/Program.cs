using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
// using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace FxAutomate
{
    class Program
    {

        public static void CopyDays(string asOfDate = "2023-03-01", 
                                    string fxCol = "£/$")
        {

            // SHOULD BE FORMAT LIKE THIS ->
            // var targetSearchDate = "1-Mar-23";

            var sourceSearchDate = DateTime.ParseExact(asOfDate, "yyyy-MM-dd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
            var targetSearchDate = DateTime.ParseExact(asOfDate, "yyyy-MM-dd", CultureInfo.InvariantCulture).ToString("d-MMM-yy");

            var excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            excelApp.EnableSound = false;
            
            try
            {
                string dir = @"C:\Users\evanl\source\repos\j4j\Sheets";
                string sourceFile = @"fx.xlsx";
                string targetFile = @"EX_Rates_FM.xlsm";
                string targetSheet = "Data";

                var sourceWorkbook = (Workbook)excelApp.Workbooks.Open(dir + "\\" + sourceFile);
                var sourceWorksheet = (Worksheet)sourceWorkbook.Sheets["fx"];

                var targetWorkbook = (Workbook)excelApp.Workbooks.Open(dir + "\\" + targetFile);
                var targetWorksheet = (Worksheet)targetWorkbook.Sheets[targetSheet];

                Excel.Range sourceRange = sourceWorksheet.Columns[1];
                Excel.Range sourceFind = sourceRange.Find(sourceSearchDate);
                var referenceRow = sourceFind.Row;

                Excel.Range targetRange = targetWorksheet.Columns[1];
                Excel.Range targetFind = targetRange.Find(What: targetSearchDate, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlWhole);
                var targetRow = targetFind.Row;

                var templateReferenceRange = (string)targetWorksheet.Range["B" + targetRow.ToString()].Formula;
                var splitFormula = templateReferenceRange.Split(new string[] { "!B" }, StringSplitOptions.None);
                var searchText = splitFormula[splitFormula.Length - 1];

                var updateRow = targetWorksheet.Range["B" + targetRow.ToString() + ":ZZ" + targetRow.ToString()];
                updateRow.Replace(What: searchText,
                                    Replacement: referenceRow,
                                    LookAt: Excel.XlLookAt.xlPart,
                                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                                    MatchCase: false,
                                    SearchFormat: false,
                                    ReplaceFormat: false
                                );

                int sourceMaxRowCount = sourceWorksheet.UsedRange.Rows.Count;
                var sourceMaxDate = (string)sourceWorksheet.Range["A" + sourceMaxRowCount.ToString()].Value;
                int offset = sourceMaxRowCount - referenceRow;

                // var fillDownRange = targetWorksheet.Range[targetRow.ToString() + ":" + (targetRow + offset).ToString()];
                var fillDownRange = targetWorksheet.Range[targetRow.ToString() + ":" + (targetRow + offset).ToString()];
                fillDownRange.FillDown();

                sourceWorkbook.Close();

                // ============================
                // Fill Downs ->
                // ============================

                var sourceMaxDateDt = Convert.ToDateTime(sourceMaxDate);
                // Daily New - >
                var weeklyDataSheet = (Worksheet)targetWorkbook.Sheets["WeeklyData"];
                var dailyNewSheet = (Worksheet)targetWorkbook.Sheets["DailyNew"];
                var monthlyDataSheet = (Worksheet)targetWorkbook.Sheets["MonthlyData"];

                Utils.FillDownUpToDate(weeklyDataSheet, sourceMaxDateDt);
                Utils.FillDownUpToDate(dailyNewSheet, sourceMaxDateDt);
                Utils.FillDownUpToDate(monthlyDataSheet, sourceMaxDateDt);


                // Save it and close.
                targetWorkbook.SaveAs(dir + "\\" + targetFile.Replace(".xlsm", "_modified.xlsm"));
                targetWorkbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }


        static void Main(string[] args)
        {
            // CopyDays();
            // CopyDays(asOfDate: "2023-02-25");
            CopyDays();
        }
    }
}
