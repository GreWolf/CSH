using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Threading.Tasks;
using Test;

namespace Test
{
    class Program
    {
        static public void Main(string[] args)
        {
            var model = new ExcelHandlerModel2();

            DataRow row = model.SummaryTableTest.NewRow();
            row["Файл"] = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31.xlsx";
            model.SummaryTableTest.Rows.Add(row);

            model.ParseExcelFiles();
            //model.SaveResult();
            model.ShowResult();
            //model.CloseApp();

        }


    }


}
