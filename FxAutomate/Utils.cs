using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace FxAutomate
{
    class Utils
    {
        //public void LoadTemplate(Microsoft.Office.Interop.Excel.Workbook currentWorkbook,
        //                         string templateFilePath,
        //                         string sheetName)
        //{
        //    this.templateWorkbook = (Workbook)Globals.ThisAddIn.Application.Workbooks.Open(templateFilePath, false, true);
        //    var templateSheet = templateWorkbook.Worksheets[1];
        //    templateSheet.Name = sheetName;
        //}
        public static string GetRelativeExcelFilePath(string folderName, string fileName)
        {
            // string buildFolderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            // string templateFilePath = Path.Combine(Application.StartupPath, @"..\..\Solution Items", templateFileName);
            // string buildFolderPath = @"C:\Users\evanl\source\repos\Surfboard\Surfboard\bin\x64\Debug\" + folderName;

            string buildFolderPath = @"C:\Users\evanl\source\repos\FxAutomate\FxAutomate\bin\x64\Debug" + folderName;
            string filePath = Path.Combine(buildFolderPath, fileName);
            return filePath;
        }

        public static void FillDownUpToDate(Worksheet worksheet, DateTime targetAsOfDate, string dateCol = "A")
        {
            var maxRow = worksheet.UsedRange.Rows.Count;
            var startRange = worksheet.Range[dateCol + maxRow.ToString()];

            try
            {
                var maxRangeAsOfDate = Convert.ToDateTime(startRange.Value);
                var increment = maxRangeAsOfDate - targetAsOfDate;

                if (maxRangeAsOfDate < targetAsOfDate)
                {
                    var fillDownRange = worksheet.Range[(maxRow).ToString() + ":" + (maxRow + increment).ToString()];
                    fillDownRange.FillDown();
                }
            }
            catch
            {
                Console.WriteLine("Failed to convert / fill down.");
                return;
            }
        }


        internal static object LoadWorkbook()
        {
            throw new NotImplementedException();
        }
    }
}
