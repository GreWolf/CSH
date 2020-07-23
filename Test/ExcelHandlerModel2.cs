using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Prism.Mvvm;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test
{
    class ExcelHandlerModel2 : BindableBase
    {

        Excel.Application app;
        Excel._Workbook workbook;
        Excel._Worksheet worksheet;
        Excel.ListObject table;

        public DataTable SummaryTableTest = new DataTable();

        public ExcelHandlerModel2()
        {

            SummaryTableTest.Clear();
            SummaryTableTest.Columns.Add("Файл");
            SummaryTableTest.Columns.Add("Статус");
            SummaryTableTest.Columns.Add("Количество контрактов", System.Type.GetType("System.Int32"));

            app = new Excel.Application();
            workbook = app.Workbooks.Add(Type.Missing);

            worksheet = workbook.ActiveSheet;

            table = worksheet.ListObjects.Add();
            table.Name = "Результат";
            Excel.Range header = table.HeaderRowRange;

            header[1, 1].Value = "ОСВ";
            header[1, 2].Value = "КФО";
            header[1, 3].Value = "Контрагент";
            header[1, 4].Value = "Договор";
            header[1, 5].Value = "Дебет";
            header[1, 6].Value = "Кредит";

            table.ListColumns.Item[5].Range.NumberFormatLocal = @"# ##0,00";
            table.ListColumns.Item[6].Range.NumberFormatLocal = @"# ##0,00";

            //table.ListColumns.Item[5].Range.NumberFormat = @"@";

        }


        public void ParseExcelFiles()
        {
            foreach (DataRow row in SummaryTableTest.Rows)
            {
                var wb_path = (string)row["Файл"];
                try
                {
                    int ContractCountPerKFO = ParseExcelFile(wb_path);
                    row["Статус"] = "Обработано";
                    row["Количество контрактов"] = ContractCountPerKFO;
                }

                catch (System.IO.IOException)
                {
                    row["Статус"] = "Ошибка";
                    row["Количество контрактов"] = DBNull.Value;
                    continue;
                }
            }

            table.ShowTotals = true;

            table.ListColumns.Item[5].TotalsCalculation = Excel.XlTotalsCalculation.xlTotalsCalculationSum;
            table.ListColumns.Item[6].TotalsCalculation = Excel.XlTotalsCalculation.xlTotalsCalculationSum;

            table.Range.Columns.AutoFit();

            table.ListRows.Item[2].Delete();

            worksheet.Columns[3].ColumnWidth = 50;
            worksheet.Columns[4].ColumnWidth = 50;

        }


        public int ParseExcelFile(string wb_path)
        {
            Excel.Workbook wb = app.Workbooks.Open(wb_path, ReadOnly: true);

            Excel._Worksheet ws = wb.ActiveSheet;

            string OSV, KFO, partner, contract;
            
            double? debet, kredit;

            int ContractCountPerKFO = 0;

            OSV = "";
            KFO = "";
            partner = "";

            bool ReadMode = false;

            foreach (Excel.Range row in ws.UsedRange.Rows)
            {
                string firstCellValue = row.Cells[1, 1].Text;
                //string firstCellValue = (string)(row.Cells[1, 1] as Excel.Range).Value;

                if (firstCellValue == "Итого")
                {
                    break;
                }

                if (firstCellValue == "Договоры")
                {
                    ReadMode = true;
                    continue;
                }

                if (ReadMode)
                {
                    switch (row.Cells[1, 1].IndentLevel)
                    {
                        case 0:
                            OSV = firstCellValue;
                            break;

                        case 2:
                            KFO = firstCellValue;
                            break;

                        case 4:
                            partner = firstCellValue;
                            break;

                        case 6:
                            contract = firstCellValue;
                            debet = (double?)(row.Cells[1, 19] as Excel.Range).Value;
                            kredit = (double?)(row.Cells[1, 21] as Excel.Range).Value;

                            if (debet != null || kredit != null)
                            {

                                Excel.ListRow destRow = table.ListRows.Add();
                                //destRow.Range.Cells[1, 1] = Convert.ToString(OSV);
                                destRow.Range.Cells[1, 1].NumberFormatLocal = @"@";
                                destRow.Range.Cells[1, 1] = OSV;
                                destRow.Range.Cells[1, 2] = KFO;
                                destRow.Range.Cells[1, 3] = partner;
                                destRow.Range.Cells[1, 4] = contract;
                                destRow.Range.Cells[1, 5] = (debet == null) ? 0 : debet;
                                destRow.Range.Cells[1, 6] = (kredit == null) ? 0 : kredit;

                                ContractCountPerKFO++;

                            }

                            break;
                    }

                }
            }

            wb.Close();
            return ContractCountPerKFO;

        }

        public void SaveResult(string path = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - result2.xlsx")
        {

            workbook.SaveAs(path);
            workbook.Close();
            CloseApp();

        }

        public void ShowResult()
        {
            app.Visible = true;
            //CloseApp();
        }

        public void CloseApp()
        {
            app.Quit();

            GC.Collect();
            Marshal.FinalReleaseComObject(table);
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(app);
        }
    }
}
