using ClosedXML.Excel;
using System;
using System.Data;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //const string fileName = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31.xlsx";

            string[] paths = {
                @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31.xlsx",
                @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 209.34.xlsx",
            };

            DataTable dt = new DataTable();

            dt.Clear();
            dt.Columns.Add("ОСВ");
            dt.Columns.Add("КФО", System.Type.GetType("System.Int32"));
            dt.Columns.Add("Контрагент");
            dt.Columns.Add("Договор");
            dt.Columns.Add("Дебет", System.Type.GetType("System.Double"));
            dt.Columns.Add("Кредит", System.Type.GetType("System.Double"));


            foreach (var wb_path in paths)
            {
                try
                {
                    ParseExcelFile(wb_path, ref dt);
                }


                catch (System.IO.IOException e)
                {
                    Console.WriteLine("{0} - ошибка {1}!", wb_path, e.GetType().Name);
                    continue;
                }

                Console.WriteLine("{0} - успешно!", wb_path);
            }

            XLWorkbook destWB = new XLWorkbook();
            destWB.Worksheets.Add(dt, "Результат");
            destWB.SaveAs(@"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - result.xlsx");
            Console.ReadKey();

        }


        static public void ParseExcelFile(string wb_path, ref DataTable dt)
        {
            using (var excelWorkbook = new XLWorkbook(wb_path))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                int n = 0;

                string OSV, KFO, partner, contract, debet, kredit;

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
                                    DataRow _row = dt.NewRow();

                                    _row["ОСВ"] = OSV;
                                    _row["КФО"] = KFO;
                                    _row["Контрагент"] = partner;
                                    _row["Договор"] = contract;
                                    _row["Дебет"] = debet == "" ? 0 : Convert.ToDouble(debet.Replace(".", ","));
                                    _row["Кредит"] = kredit == "" ? 0 : Convert.ToDouble(kredit.Replace(".", ","));

                                    dt.Rows.Add(_row);

                                }

                                n += 1;

                                break;
                        }
                    }
                }
            }
        }




    }
}
