using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static Sitek.Program;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sitek
{
    public class ExcelWorksheet
    {
        public void CreateFile(Root json)
        {
            var curDir = Environment.CurrentDirectory;
            var filename = string.Format(@"\Контролирующие органы.xlsx");
            var filepath = (curDir + filename);

            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Add();
            Excel._Worksheet worksheet = excelApp.ActiveSheet;

            worksheet.Cells[1, "A"] = "Номер";
            worksheet.Cells[1, "B"] = "Тип";
            worksheet.Cells[1, "C"] = "Код";
            worksheet.Cells[1, "D"] = "Название";
            worksheet.Cells[1, "E"] = "Регион";

            var dataTable = new System.Data.DataTable();

            int row = 2;

            foreach (var item in json.cus.OrderBy(x => x.type).ThenBy(x => x.code))
            {
                if (item.region == "18")
                {
                    worksheet.Cells[row, 1] = $"{row}";
                    worksheet.Cells[row, 2] = (string)item.type;
                    worksheet.Cells[row, 3] = (string)item.code;
                    worksheet.Cells[row, 4] = (string)item.name;
                    worksheet.Cells[row, 5] = (string)item.region;
                    row++;
                }
            }

            Console.WriteLine("\n" + "Excel был создан по пути: " + filepath);

            excelApp.DisplayAlerts = false;
            worksheet.SaveAs(filepath);
            workbook.Close();
            excelApp.Quit();
        }
    }
}
