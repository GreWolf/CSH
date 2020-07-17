using System;
using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;

namespace Test
{
    class ExcelHandler
    {
        public List<string> paths { get; set; } = new List<string> {
                @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31.xlsx",
                @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 209.34.xlsx",
            };

        private DataTable resultTable = new DataTable();
        private DataTable SummaryTable = new DataTable();



        public ExcelHandler()
        {
            resultTable.Clear();
            resultTable.Columns.Add("ОСВ");
            resultTable.Columns.Add("КФО", System.Type.GetType("System.Int32"));
            resultTable.Columns.Add("Контрагент");
            resultTable.Columns.Add("Договор");
            resultTable.Columns.Add("Дебет", System.Type.GetType("System.Double"));
            resultTable.Columns.Add("Кредит", System.Type.GetType("System.Double"));

            SummaryTable.Clear();
            SummaryTable.Columns.Add("ОСВ");
            SummaryTable.Columns.Add("Количество контрактов", System.Type.GetType("System.Int32"));

        }
       


        public void ParseExcelFiles()
        {
            foreach (var wb_path in paths)
            {
                try
                {
                    ParseExcelFile(wb_path);
                }


                catch (System.IO.IOException e)
                {
                    Console.WriteLine("{0} - ошибка {1}!", wb_path, e.GetType().Name);
                    continue;
                }

                Console.WriteLine("{0} - успешно!", wb_path);
            }
        }




        private void ParseExcelFile(string wb_path)
        {
            using (var excelWorkbook = new XLWorkbook(wb_path))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                int n = 0;

                string OSV, KFO, partner, contract, debet, kredit;

                int ContractCountPerKFO = 0;

                OSV = "";
                KFO = "";
                partner = "";

                bool ReadMode = false;

                foreach (var dataRow in nonEmptyDataRows)
                {

                    var cell = dataRow.Cell(1);
                    var cellValue = cell.Value.ToString();

                    if (cellValue == "Итого")
                    {
                        break;
                    }

                    if (cellValue == "Договоры")
                    {
                        ReadMode = true;
                        continue;
                    }

                    if (ReadMode == true)
                    {

                        switch (cell.Style.Alignment.Indent)
                        {
                            case 0:
                                OSV = cellValue;
                                break;

                            case 2:
                                KFO = cellValue;
                                break;

                            case 4:
                                partner = cellValue;
                                break;

                            case 6:
                                contract = cellValue;
                                debet = dataRow.Cell(19).Value.ToString();
                                kredit = dataRow.Cell(21).Value.ToString();

                                if (debet != "" || kredit != "")
                                {
                                    DataRow _row = resultTable.NewRow();

                                    _row["ОСВ"] = OSV;
                                    _row["КФО"] = KFO;
                                    _row["Контрагент"] = partner;
                                    _row["Договор"] = contract;
                                    _row["Дебет"] = debet == "" ? 0 : Convert.ToDouble(debet.Replace(".", ","));
                                    _row["Кредит"] = kredit == "" ? 0 : Convert.ToDouble(kredit.Replace(".", ","));

                                    resultTable.Rows.Add(_row);
                                    ContractCountPerKFO++;

                                }

                                n += 1;

                                break;
                        }
                    }
                }

                DataRow _summarizeRow = SummaryTable.NewRow();
                _summarizeRow["ОСВ"] = OSV;
                _summarizeRow["Количество контрактов"] = ContractCountPerKFO;
                SummaryTable.Rows.Add(_summarizeRow);

            }
        }



        public void SaveResult(string path = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - result.xlsx")
        {
            XLWorkbook destWB = new XLWorkbook();
            destWB.Worksheets.Add(resultTable, "Результат");
            destWB.SaveAs(path);
            //Console.ReadKey();
        }


        public void SaveSummary(string path = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - summarize.xlsx")
        {
            XLWorkbook destWB = new XLWorkbook();
            destWB.Worksheets.Add(SummaryTable, "Итоги");
            destWB.SaveAs(path);
            //Console.ReadKey();
        }


    }
}
