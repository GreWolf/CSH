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


            //Start = new DelegateCommand(() =>
            //{
            //    _model.ResultTable.Clear();
            //    _model.ParseExcelFiles();
            //    _model.ExportToExcel(_model.ResultTable);
            //});

            Start = new DelegateCommand(async () => await _start());
        }


        public DelegateCommand Start { get;  }

        public DataTable ResultTable => _model.ResultTable;
        
        public DataTable SummaryTableTest => _model.SummaryTableTest;


        private async Task _start()
        {
            await Task.Run(() => _model.ResultTable.Clear());
            await Task.Run(() => _model.ParseExcelFiles());
            await Task.Run(() => _model.ExportToExcel(_model.ResultTable));
            //_model.ExportToExcel(_model.ResultTable);
        }




    }

}
