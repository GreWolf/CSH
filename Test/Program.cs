using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            const string fileName = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31.xlsx";

            DataTable dt = new DataTable();

            dt.Clear();
            dt.Columns.Add("ОСВ");
            dt.Columns.Add("КФО");
            dt.Columns.Add("Контрагент");
            dt.Columns.Add("Договор");

            DataColumn colString = new DataColumn("Дебет");
            colString.DataType = System.Type.GetType("System.Double");
            dt.Columns.Add(colString);

            DataColumn colString_2 = new DataColumn("Кредит");
            colString_2.DataType = System.Type.GetType("System.Double");
            dt.Columns.Add(colString_2);

            bool ReadMode = false;

            try
            {
                using (var excelWorkbook = new XLWorkbook(fileName))
                {
                    var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                    int n = 0;


                    string OSV, KFO, partner, contract, debet, kredit;
                    OSV = "";
                    KFO = "";
                    partner = "";

                    foreach (var dataRow in nonEmptyDataRows)
                    {

                        

                        var cell = dataRow.Cell(1);

                        if (cell.Value.ToString() == "Итого")
                        {
                            break;
                        }

                        if (cell.Value.ToString() == "Договоры")
                        {
                            ReadMode = true;
                            continue;
                        }

                        if (ReadMode == true) {

                            switch (cell.Style.Alignment.Indent)
                            {
                                case 0:
                                    OSV = cell.Value.ToString();
                                    //Console.WriteLine("ОСВ");
                                    break;
                                case 2:
                                    KFO = cell.Value.ToString();
                                    //Console.WriteLine("КФО");
                                    break;
                                case 4:
                                    partner = cell.Value.ToString();
                                    //Console.WriteLine("Поставщик");
                                    break;
                                case 6:
                                    contract = cell.Value.ToString();
                                    debet = dataRow.Cell(19).Value.ToString();
                                    kredit = dataRow.Cell(21).Value.ToString();

                                    DataRow _row = dt.NewRow();

                                    _row["ОСВ"] = OSV;
                                    _row["КФО"] = KFO;
                                    _row["Контрагент"] = partner;
                                    _row["Договор"] = contract;
                                    var dbt = debet == "" ? 0 : Convert.ToDouble(debet.Replace(".", ","));
                                    _row["Дебет"] = debet == "" ? 0 : Convert.ToDouble(debet.Replace(".", ","));
                                    _row["Кредит"] = kredit == "" ? 0 : Convert.ToDouble(kredit.Replace(".", ","));

                                    dt.Rows.Add(_row);

                                    Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", OSV, KFO, partner, contract, debet, kredit);

                                    n += 1;

                                    break;
                            }       
                            

                            //Console.WriteLine(cell.Value);
                            //Console.WriteLine(cell.Style.Alignment.Indent);
                        }

                        //if (n >= 20)
                        //{
                        //    break;
                        //}

                        
                    }

                    //Console.ReadKey();
                }
            }


            catch (System.IO.IOException e)
            {
                Console.WriteLine("{0} Exception caught!", e);
                Console.ReadKey();
            }

            XLWorkbook destWB = new XLWorkbook();
            destWB.Worksheets.Add(dt, "Результат");
            destWB.SaveAs(@"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - result.xlsx");

        }
    }
}
