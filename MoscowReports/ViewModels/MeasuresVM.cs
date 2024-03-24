using Files.Abstracts;
using Files;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Files.Tables;

namespace MoscowReports
{
    public class MeasuresVM : INotifyPropertyChanged
    {
        private IProgress<object[]> _progress;

        public string? _sourcePath;
        public string? _targetPath;
        private OpenFileDialog _openFileDialog { get; set; }

        private RelayCommand? _sourcePathCommand;
        private RelayCommand? _targetPathCommand;
        private RelayCommand? _runCommand;
        private string _filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

        public event PropertyChangedEventHandler? PropertyChanged;
        public string? SourcePath 
        {
            get => _sourcePath;
            set
            {
                if (_sourcePath != value)
                {
                    _sourcePath = value;
                    RunCommand.NotifyCanExecuteChanged();
                    OnPropertyChanged();
                }
            }
        }
        public string? TargetPath
        {
            get => _targetPath;
            set
            {
                if (_targetPath != value)
                {
                    _targetPath = value;
                    RunCommand.NotifyCanExecuteChanged();
                    OnPropertyChanged();
                }
            }
        }
        public MeasuresVM(IProgress<object[]> progress)
        {
            _openFileDialog = new OpenFileDialog();
            _openFileDialog.Filter = _filter;
            _progress = progress;
        }

        public RelayCommand SourcePathCommand
        {
            get
            {
                return _sourcePathCommand ?? (_sourcePathCommand = new RelayCommand((ob) =>
                  {
                      try
                      {
                          SourcePath = GetPath();
                      }
                      catch (Exception ex)
                      {
                          MessageBox.Show(ex.Message);
                      }
                  }));
            }
        }
        public RelayCommand TargetPathCommand
        {
            get
            {
                return _targetPathCommand ?? (_targetPathCommand = new RelayCommand((ob) =>
                {
                    try
                    {
                        TargetPath = GetPath();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }));
            }
        }
        public RelayCommand RunCommand
        {
            get
            {
                return _runCommand ?? (_runCommand = new RelayCommand((ob) =>
                {

                    try
                    {
                        Thread myThread = new Thread(() =>
                        {
                            DataAnalyzer dataAnalyzer = new DataAnalyzer();
                            object sourceTable = null;

                            switch (dataAnalyzer.CheckTypeOfTable(SourcePath))
                            {
                                case Table.AcumulativeTable:
                                    sourceTable = new AcumulativeTable(dataAnalyzer, SourcePath); ;
                                    break;
                                case Table.MeasuresTable:
                                    sourceTable = new MeasuresTable(dataAnalyzer, SourcePath);
                                    break;
                                default:
                                    break;
                            }
                            MeasuresTable receiverTable = new MeasuresTable(dataAnalyzer, TargetPath);
                            FileHandler handler = new MeasuresFileHandler(sourceTable, receiverTable, _progress);

                            handler.HandleFiles();

                            DialogResult result = MessageBox.Show("Мероприятия заполнены!",
                                "Сообщение",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                        });
                        myThread.IsBackground = true;
                        myThread.Start();
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                },(ob) => SourcePath != TargetPath));
            }
        }
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
        private string GetPath()
        {
            _openFileDialog.ShowDialog();
            return _openFileDialog.FileName;
        }
    }
}
