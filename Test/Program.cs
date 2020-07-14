using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31.xlsx";

            try
            {
                using (var excelWorkbook = new XLWorkbook(fileName))
                {
                    var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                    foreach (var dataRow in nonEmptyDataRows)
                    {
                        var cell = dataRow.Cell(1);
                        Console.WriteLine(cell.Value);
                        Console.WriteLine(cell.Style.Alignment.Indent);
                    }

                    Console.ReadKey();
                }
            }
            catch (System.IO.IOException e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Console.ReadKey();
            }
            
        }
    }
}
