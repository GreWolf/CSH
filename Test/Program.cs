using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {

            List<string> paths = new List<string> {
                @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31.xlsx",
                @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 209.34.xlsx",
            };

            var EHandler = new ExcelHandlerModel();

            EHandler.paths = paths;

            EHandler.ParseExcelFiles();
            EHandler.SaveResult(@"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - result.xlsx");
            EHandler.SaveSummary(@"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - summarize.xlsx");

        }
    }
}
