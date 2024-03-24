using System.ComponentModel;
using System.Runtime.CompilerServices;
using Logging;
using Files;
using Files.Abstracts;
using Files.Tables;
namespace MoscowReports.ViewModels
{
    public class MoscowReportVM : INotifyPropertyChanged
    {
        private IProgress<object[]> _progress;
      
        private string? _moscowTablePath;
        private string? _piktsTablePath;
        private string? _annexesPath;
        private string _filesFolder;
        private DateTime _limitDate;
        private bool _errorFiles;
        private OpenFileDialog _openFileDialog;
        private FolderBrowserDialog _folderBrowserDialog;

        private RelayCommand? _moscowTablePathCommand;
        private RelayCommand? _piktsTablePathCommand;
        private RelayCommand? _annexesPathCommand;
        private RelayCommand? _runCommand;
        private string _filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

        public event PropertyChangedEventHandler? PropertyChanged;
        public MoscowReportVM(IProgress<object[]> progress)
        {
            
            _limitDate = DateTime.Now.AddMonths(-6);
            _folderBrowserDialog = new FolderBrowserDialog();
            _openFileDialog = new OpenFileDialog();
            _openFileDialog.Filter = _filter;
            _progress = progress;

            
        }
        public string? MoscowTablePath
        {
            get => _moscowTablePath;
            set
            {
                if (_moscowTablePath != value)
                {
                    _moscowTablePath = value;
                    RunCommand.NotifyCanExecuteChanged();
                    OnPropertyChanged();
                }
            }
        }
        public string? PiktsTablePath
        {
            get => _piktsTablePath;
            set
            {
                if (_piktsTablePath != value)
                {
                    _piktsTablePath = value;
                    RunCommand.NotifyCanExecuteChanged();
                    OnPropertyChanged();
                }
            }
        }
        public string? AnnexesPath
        {
            get => _annexesPath;
            set
            {
                if (_annexesPath != value)
                {
                    _annexesPath = value;
                    OnPropertyChanged();
                }
            }
        }
        public DateTime LimitDate
        {
            get => _limitDate;
            set
            {
                if (_limitDate != value)
                {
                    _limitDate = value;
                    OnPropertyChanged();
                }
            }
        }
        public bool ErrorFiles
        {
            get => _errorFiles;
            set
            {
                if (_errorFiles != value)
                {
                    _errorFiles = value;
                    OnPropertyChanged();
                }
            }
        }
        public RelayCommand MoscowTablePathCommand
        {
            get
            {
                return _moscowTablePathCommand ?? (_moscowTablePathCommand = new RelayCommand((ob) =>
                {
                    try
                    {
                        MoscowTablePath = GetPath();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }));
            }
        }
        public RelayCommand PiktsTablePathCommand
        {
            get
            {
                return _piktsTablePathCommand ?? (_piktsTablePathCommand = new RelayCommand((ob) =>
                {
                    try
                    {
                        PiktsTablePath = GetPath();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }));
            }
        }
        public RelayCommand AnnexesPathCommand
        {
            get
            {
                return _annexesPathCommand ?? (_annexesPathCommand = new RelayCommand((ob) =>
                {
                    try
                    {
                        if (_folderBrowserDialog.ShowDialog() == DialogResult.Cancel)
                            return;

                        AnnexesPath = _folderBrowserDialog.SelectedPath;
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
                            MoscowTable moscowTable = new MoscowTable(dataAnalyzer, MoscowTablePath);
                            PiktsTable piktsTable = new PiktsTable(dataAnalyzer, PiktsTablePath);
                            _filesFolder = CreateSaveFolder();

                            FileManager fileManager = new SurveyReportsFilesLoader(LimitDate, AnnexesPath,
                                _filesFolder, _progress);
                            FileHandler fileHandler = new SurveyReportsFileHandler(dataAnalyzer, moscowTable,
                                _filesFolder, piktsTable, _progress);

                            if (ErrorFiles)
                                (fileHandler as SurveyReportsFileHandler).ErrorFiles = true;

                            Loger.ClearLogFile();
                            
                            fileManager.GetInputFiles();
                            fileHandler.HandleFiles();

                            DialogResult result = MessageBox.Show("Московская таблица заполнена!\n\tЖми ОК",
                                "Сообщение",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);

                            if (result == DialogResult.OK)
                                Loger.OpenLogFile();

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
                },(ob) => MoscowTablePath != PiktsTablePath));
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
        private string CreateSaveFolder()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string savingPath = Path.Combine(desktopPath, "Приложения А для московской таблицы");

            if (!Directory.Exists(savingPath))
                Directory.CreateDirectory(savingPath);

            return savingPath;
        }
        
    }
}
