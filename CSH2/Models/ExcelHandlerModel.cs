using System;
using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using Prism.Mvvm;
using Excel = Microsoft.Office.Interop.Excel;


namespace CSH2.Models
{
    class ExcelHandlerModel : BindableBase
    {
    
        public DataTable ResultTable = new DataTable();

        public DataTable SummaryTableTest = new DataTable();

        public ExcelHandlerModel()
        {
            ResultTable.Clear();
            ResultTable.Columns.Add("ОСВ");
            ResultTable.Columns.Add("КФО", System.Type.GetType("System.Int32"));
            ResultTable.Columns.Add("Контрагент");
            ResultTable.Columns.Add("Договор");
            ResultTable.Columns.Add("Дебет", System.Type.GetType("System.Double"));
            ResultTable.Columns.Add("Кредит", System.Type.GetType("System.Double"));

            SummaryTableTest.Clear();
            SummaryTableTest.Columns.Add("Файл");
            SummaryTableTest.Columns.Add("Статус");
            SummaryTableTest.Columns.Add("Количество контрактов", System.Type.GetType("System.Int32"));

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
            
        }


        private int ParseExcelFile(string wb_path)
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
                                    DataRow _ResultRow = ResultTable.NewRow();

                                    _ResultRow["ОСВ"] = OSV;
                                    _ResultRow["КФО"] = KFO;
                                    _ResultRow["Контрагент"] = partner;
                                    _ResultRow["Договор"] = contract;
                                    _ResultRow["Дебет"] = debet == "" ? 0 : Convert.ToDouble(debet.Replace(".", ","));
                                    _ResultRow["Кредит"] = kredit == "" ? 0 : Convert.ToDouble(kredit.Replace(".", ","));

                                    ResultTable.Rows.Add(_ResultRow);
                                    RaisePropertyChanged(nameof(ResultTable));

                                    ContractCountPerKFO++;

                                }

                                n += 1;

                                break;
                        }
                    }
                }


                return ContractCountPerKFO;

            }
        }


        public void SaveResult(string path = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - result.xlsx")
        {
            XLWorkbook destWB = new XLWorkbook();
            destWB.Worksheets.Add(ResultTable, "Результат");
            destWB.SaveAs(path);
            //Console.ReadKey();
        }


        public void SaveSummary(string path = @"D:\GoogleDrive\Roslesinforg\Дела\2020.07.14 - Ц\ОСВ 205.31 - summarize.xlsx")
        {
            XLWorkbook destWB = new XLWorkbook();
            destWB.Worksheets.Add(SummaryTableTest, "Итоги");
            destWB.SaveAs(path);
            //Console.ReadKey();
        }


        public void ExportToExcel(DataTable tbl, string excelFilePath = null)
        {
            try
            {
                if (tbl == null || tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var excelApp = new Excel.Application();
                excelApp.Workbooks.Add();

                // single worksheet
                Excel._Worksheet workSheet = excelApp.ActiveSheet;

                // column headings
                for (var i = 0; i < tbl.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = tbl.Columns[i].ColumnName;
                }

                // rows
                for (var i = 0; i < tbl.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < tbl.Columns.Count; j++)
                    {
                        workSheet.Cells[i + 2, j + 1] = tbl.Rows[i][j];
                    }
                }

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        workSheet.SaveAs(excelFilePath);
                        excelApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                            + ex.Message);
                    }
                }
                else
                { // no file path is given
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }


    }
}
