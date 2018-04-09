using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acts
{
    public class ExcelImporter
    {
        private string Path { get; }
        public ExcelImporter()
        {
            Path = "C:\\temp\\Values.xlsx";
        }

        public ExcelImporter(string path)
        {
            Path = path;
        }

        public string[,] GetData()
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();

            //Книга
            var workBookExcel = excelApp.Workbooks.Open(Path);

            //Таблица
            var workSheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)workBookExcel.Sheets[1];

            var lastCell = workSheetExcel.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            var result = new string[lastCell.Column, lastCell.Row];

            for (int i = 0; i < (int)lastCell.Column; i++)
            {
                for (int j = 0; j < (int)lastCell.Row; j++)
                {
                    result[i, j] = workSheetExcel.Cells[j + 1, i + 1].Text.ToString();
                }
            }

            workBookExcel.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            excelApp.Quit(); // вышел из Excel
            GC.Collect(); // убрал за собой

            return result;
        }
    }
}
