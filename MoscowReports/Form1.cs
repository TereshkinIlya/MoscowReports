using Files;
using Files.Abstracts;
using MoscowReports.ViewModels;
using System.Diagnostics;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MoscowReports
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(this.Form1_FormClosing);

            dateTimePicker1.Value = DateTime.Now.AddMonths(-6);
            Progress<object[]> _moscProgress = new Progress<object[]>();
            Progress<object[]> _measProgress = new Progress<object[]>();

            _moscProgress.ProgressChanged += (s, args) =>
            {
                progressBar1.Maximum = (int)args[0];
                progressBar1.Value = (int)args[1];
                moscPrgBarText.Text = (string)args[2];
                progressBar1.Style = (ProgressBarStyle)args[3];
                progressBar1.Update();
            };
            _measProgress.ProgressChanged += (s, args) =>
            {
                progressBar2.Maximum = (int)args[0];
                progressBar2.Value = (int)args[1];
                measPrgBarText.Text = (string)args[2];
                progressBar2.Style = (ProgressBarStyle)args[3];
                progressBar2.Update();
            };

            Tabcontrol.TabPages["Measures"].DataContext = new MeasuresVM(_measProgress);
            Tabcontrol.TabPages["MoscowReport"].DataContext = new MoscowReportVM(_moscProgress);

            textBox1.DataBindings.Add(new Binding("Text", Tabcontrol.TabPages["Measures"].DataContext, "SourcePath"));
            textBox2.DataBindings.Add(new Binding("Text", Tabcontrol.TabPages["Measures"].DataContext, "TargetPath"));

            textBox3.DataBindings.Add(new Binding("Text", Tabcontrol.TabPages["MoscowReport"].DataContext, "MoscowTablePath"));
            textBox4.DataBindings.Add(new Binding("Text", Tabcontrol.TabPages["MoscowReport"].DataContext, "PiktsTablePath"));
            textBox5.DataBindings.Add(new Binding("Text", Tabcontrol.TabPages["MoscowReport"].DataContext, "AnnexesPath"));
            dateTimePicker1.DataBindings.Add(new Binding("Value", Tabcontrol.TabPages["MoscowReport"].DataContext, "LimitDate"));
            checkBox1.DataBindings.Add(new Binding("Checked", Tabcontrol.TabPages["MoscowReport"].DataContext, "ErrorFiles"));
            
            sourceButton.DataBindings.Add(new Binding("Command", Tabcontrol.TabPages["Measures"].DataContext, "SourcePathCommand", true));
            receiverButton.DataBindings.Add(new Binding("Command", Tabcontrol.TabPages["Measures"].DataContext, "TargetPathCommand", true));

            moscowTableButton.DataBindings.Add(new Binding("Command", Tabcontrol.TabPages["MoscowReport"].DataContext, "MoscowTablePathCommand", true));
            piktsButton.DataBindings.Add(new Binding("Command", Tabcontrol.TabPages["MoscowReport"].DataContext, "PiktsTablePathCommand", true));
            annexesButton.DataBindings.Add(new Binding("Command", Tabcontrol.TabPages["MoscowReport"].DataContext, "AnnexesPathCommand", true));
            launchButton2.DataBindings.Add(new Binding("Command", Tabcontrol.TabPages["MoscowReport"].DataContext, "RunCommand", true));
            
            launchButton1.DataBindings.Add(new Binding("Command", Tabcontrol.TabPages["Measures"].DataContext, "RunCommand", true));

        }

        private void launchButton1_Click(object sender, EventArgs e)
        {
            launchButton1.Enabled = false;
            measPrgBarText.Visible = true;
            progressBar2.Style = ProgressBarStyle.Marquee;
        }

        private void launchButton2_Click(object sender, EventArgs e)
        {
            launchButton2.Enabled = false;
            moscPrgBarText.Visible = true;
            checkBox1.Enabled = false;
            progressBar1.Style = ProgressBarStyle.Marquee;
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            int id;

            try
            {
                GetWindowThreadProcessId(ExcelApp.Run.Hwnd, out id);
                
                Process process = Process.GetProcessById(id);
                
                if (!process.HasExited)
                    process.Kill();
            }
            catch(Exception) { return; }
            
            Environment.Exit(Environment.ExitCode);
        }
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
    }
}   
