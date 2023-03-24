using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sitek
{
    public class ExcelWorksheet
    {
        public void CreateFile()
        {
            var filepath = "Контролируещие органы.xlsx";
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Add();
            var sheet = workbook.ActiveSheet as ExcelWorksheet;

            workbook.SaveAs(filepath);
            workbook.Close();
            excelApp.Quit();
        }
    }
}
