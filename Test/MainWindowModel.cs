using DocumentFormat.OpenXml.Bibliography;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class MainWindowModel : BindableBase
    {
        readonly ExcelHandlerModel _model = new ExcelHandlerModel();

        public MainWindowModel()
        {
            _model.PropertyChanged += (s, e) => { RaisePropertyChanged(e.PropertyName); };

            SetPaths = new DelegateCommand<List<string>>(paths =>
            {
                _model.paths = paths;
            }
                );

        }

        public DelegateCommand<List<string>> SetPaths { get;  }

        public DataTable ResultTable => _model.ResultTable;
        public DataTable SummaryTable => _model.SummaryTable;
    }
}
