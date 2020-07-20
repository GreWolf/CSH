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

namespace CSH2.ViewModels
{
    class MainWindowModel : BindableBase
    {
        readonly ExcelHandlerModel _model = new ExcelHandlerModel();

        public MainWindowModel()
        {
            _model.PropertyChanged += (s, e) => { RaisePropertyChanged(e.PropertyName); };

            //SetPaths = new DelegateCommand<List<string>>(paths =>
            //{
            //    _model.paths = paths;

            //});

            Start = new DelegateCommand(() => {
                _model.ResultTable.Clear();
                _model.SummaryTable.Clear();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                _model.ParseExcelFiles();
                //_model.SaveResult();
                //_model.SaveSummary();   
            });
        }

        //public DelegateCommand<List<string>> SetPaths { get;  }

        public DelegateCommand Start { get;  }

        public ObservableCollection<string> Paths => _model.paths;

        public DataTable ResultTable => _model.ResultTable;
        public DataTable SummaryTable => _model.SummaryTable;

    }

}
