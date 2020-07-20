using CSH2.Models;
using DocumentFormat.OpenXml.Bibliography;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using Microsoft.Win32;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSH2.ViewModels
{
    class MainWindowModel : BindableBase
    {
        readonly ExcelHandlerModel _model = new ExcelHandlerModel();

        public MainWindowModel()
        {
            _model.PropertyChanged += (s, e) => { RaisePropertyChanged(e.PropertyName); };


            Start = new DelegateCommand(() => {
                _model.ResultTable.Clear();
                //_model.SummaryTableTest.Clear();
                _model.ParseExcelFiles();
                _model.ExportToExcel(_model.ResultTable);


                //string temppath = Path.GetTempPath() + "temp.xlsx";
                //_model.SaveResult(temppath);
                //Console.WriteLine(temppath);

                //Excel.Application xlApp = new Excel.Application();  // create new Excel application
                //xlApp.Visible = true;                               // application becomes visible
                //xlApp.Workbooks.Open(temppath, ReadOnly: true);          // open the workbook from file path

                //xlApp.Quit();
                //_model.SaveResult();
                //_model.SaveSummary();   
            });
        }


        public DelegateCommand Start { get;  }

        public DataTable ResultTable => _model.ResultTable;
        
        public DataTable SummaryTableTest => _model.SummaryTableTest;




    }

}
