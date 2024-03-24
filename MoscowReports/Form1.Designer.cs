namespace MoscowReports
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Tabcontrol = new TabControl();
            Measures = new TabPage();
            label10 = new Label();
            label9 = new Label();
            measPrgBarText = new Label();
            label8 = new Label();
            progressBar2 = new ProgressBar();
            launchButton1 = new Button();
            receiverButton = new Button();
            textBox2 = new TextBox();
            sourceButton = new Button();
            textBox1 = new TextBox();
            label2 = new Label();
            label1 = new Label();
            MoscowReport = new TabPage();
            checkBox1 = new CheckBox();
            moscPrgBarText = new Label();
            label7 = new Label();
            dateTimePicker1 = new DateTimePicker();
            progressBar1 = new ProgressBar();
            launchButton2 = new Button();
            label6 = new Label();
            annexesButton = new Button();
            textBox5 = new TextBox();
            label5 = new Label();
            piktsButton = new Button();
            textBox4 = new TextBox();
            label4 = new Label();
            moscowTableButton = new Button();
            textBox3 = new TextBox();
            label3 = new Label();
            Tabcontrol.SuspendLayout();
            Measures.SuspendLayout();
            MoscowReport.SuspendLayout();
            SuspendLayout();
            // 
            // Tabcontrol
            // 
            Tabcontrol.Controls.Add(Measures);
            Tabcontrol.Controls.Add(MoscowReport);
            Tabcontrol.Font = new Font("Segoe UI", 13F);
            Tabcontrol.Location = new Point(0, 0);
            Tabcontrol.Name = "Tabcontrol";
            Tabcontrol.SelectedIndex = 0;
            Tabcontrol.Size = new Size(808, 356);
            Tabcontrol.TabIndex = 0;
            // 
            // Measures
            // 
            Measures.Controls.Add(label10);
            Measures.Controls.Add(label9);
            Measures.Controls.Add(measPrgBarText);
            Measures.Controls.Add(label8);
            Measures.Controls.Add(progressBar2);
            Measures.Controls.Add(launchButton1);
            Measures.Controls.Add(receiverButton);
            Measures.Controls.Add(textBox2);
            Measures.Controls.Add(sourceButton);
            Measures.Controls.Add(textBox1);
            Measures.Controls.Add(label2);
            Measures.Controls.Add(label1);
            Measures.Location = new Point(4, 32);
            Measures.Name = "Measures";
            Measures.Padding = new Padding(3);
            Measures.Size = new Size(800, 320);
            Measures.TabIndex = 0;
            Measures.Text = "Мероприятия";
            Measures.UseVisualStyleBackColor = true;
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Font = new Font("Segoe UI", 10F);
            label10.ForeColor = Color.Red;
            label10.Location = new Point(161, 168);
            label10.Name = "label10";
            label10.Size = new Size(75, 19);
            label10.TabIndex = 19;
            label10.Text = "Приёмник";
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Font = new Font("Segoe UI", 10F);
            label9.ForeColor = Color.Red;
            label9.Location = new Point(161, 66);
            label9.Name = "label9";
            label9.Size = new Size(70, 19);
            label9.TabIndex = 18;
            label9.Text = "Источник";
            // 
            // measPrgBarText
            // 
            measPrgBarText.AutoSize = true;
            measPrgBarText.Font = new Font("Franklin Gothic Book", 9.75F, FontStyle.Italic, GraphicsUnit.Point, 204);
            measPrgBarText.Location = new Point(286, 224);
            measPrgBarText.Name = "measPrgBarText";
            measPrgBarText.Size = new Size(134, 17);
            measPrgBarText.TabIndex = 17;
            measPrgBarText.Text = "Подготовка к работе...";
            measPrgBarText.Visible = false;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Font = new Font("Segoe UI", 10F);
            label8.ForeColor = Color.Red;
            label8.Location = new Point(198, 16);
            label8.Name = "label8";
            label8.Size = new Size(384, 19);
            label8.TabIndex = 15;
            label8.Text = "Не работай с оригиналом! Копируй файл на жесткий диск!";
            // 
            // progressBar2
            // 
            progressBar2.Location = new Point(8, 221);
            progressBar2.Name = "progressBar2";
            progressBar2.Size = new Size(769, 23);
            progressBar2.Step = 1;
            progressBar2.Style = ProgressBarStyle.Continuous;
            progressBar2.TabIndex = 14;
            // 
            // launchButton1
            // 
            launchButton1.Location = new Point(683, 261);
            launchButton1.Name = "launchButton1";
            launchButton1.Size = new Size(94, 38);
            launchButton1.TabIndex = 6;
            launchButton1.Text = "Запуск";
            launchButton1.UseVisualStyleBackColor = true;
            launchButton1.Click += launchButton1_Click;
            // 
            // receiverButton
            // 
            receiverButton.Location = new Point(676, 140);
            receiverButton.Name = "receiverButton";
            receiverButton.Size = new Size(101, 31);
            receiverButton.TabIndex = 5;
            receiverButton.Text = "Выбрать";
            receiverButton.UseVisualStyleBackColor = true;
            // 
            // textBox2
            // 
            textBox2.Font = new Font("Segoe UI", 10F);
            textBox2.Location = new Point(161, 140);
            textBox2.Name = "textBox2";
            textBox2.ReadOnly = true;
            textBox2.Size = new Size(475, 25);
            textBox2.TabIndex = 4;
            // 
            // sourceButton
            // 
            sourceButton.Location = new Point(676, 33);
            sourceButton.Name = "sourceButton";
            sourceButton.Size = new Size(101, 31);
            sourceButton.TabIndex = 3;
            sourceButton.Text = "Выбрать";
            sourceButton.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            textBox1.Font = new Font("Segoe UI", 10F);
            textBox1.Location = new Point(161, 38);
            textBox1.Name = "textBox1";
            textBox1.ReadOnly = true;
            textBox1.Size = new Size(475, 25);
            textBox1.TabIndex = 2;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 13F);
            label2.Location = new Point(8, 139);
            label2.Name = "label2";
            label2.Size = new Size(125, 25);
            label2.TabIndex = 1;
            label2.Text = "Мероприятия";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 13F);
            label1.Location = new Point(8, 12);
            label1.Name = "label1";
            label1.Size = new Size(136, 75);
            label1.TabIndex = 0;
            label1.Text = "Мероприятия\r\n         или\r\nНакопительная";
            // 
            // MoscowReport
            // 
            MoscowReport.Controls.Add(checkBox1);
            MoscowReport.Controls.Add(moscPrgBarText);
            MoscowReport.Controls.Add(label7);
            MoscowReport.Controls.Add(dateTimePicker1);
            MoscowReport.Controls.Add(progressBar1);
            MoscowReport.Controls.Add(launchButton2);
            MoscowReport.Controls.Add(label6);
            MoscowReport.Controls.Add(annexesButton);
            MoscowReport.Controls.Add(textBox5);
            MoscowReport.Controls.Add(label5);
            MoscowReport.Controls.Add(piktsButton);
            MoscowReport.Controls.Add(textBox4);
            MoscowReport.Controls.Add(label4);
            MoscowReport.Controls.Add(moscowTableButton);
            MoscowReport.Controls.Add(textBox3);
            MoscowReport.Controls.Add(label3);
            MoscowReport.Location = new Point(4, 32);
            MoscowReport.Name = "MoscowReport";
            MoscowReport.Padding = new Padding(3);
            MoscowReport.RightToLeft = RightToLeft.No;
            MoscowReport.Size = new Size(800, 320);
            MoscowReport.TabIndex = 1;
            MoscowReport.Text = "Московская таблица";
            MoscowReport.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(391, 270);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(189, 29);
            checkBox1.TabIndex = 17;
            checkBox1.Text = "Файлы с ошибками";
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // moscPrgBarText
            // 
            moscPrgBarText.AutoSize = true;
            moscPrgBarText.Font = new Font("Franklin Gothic Book", 9.75F, FontStyle.Italic, GraphicsUnit.Point, 204);
            moscPrgBarText.Location = new Point(324, 225);
            moscPrgBarText.Name = "moscPrgBarText";
            moscPrgBarText.Size = new Size(134, 17);
            moscPrgBarText.TabIndex = 16;
            moscPrgBarText.Text = "Подготовка к работе...";
            moscPrgBarText.Visible = false;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(8, 268);
            label7.Name = "label7";
            label7.Size = new Size(156, 25);
            label7.TabIndex = 15;
            label7.Text = "Приложения А с :";
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            dateTimePicker1.Location = new Point(196, 268);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(128, 31);
            dateTimePicker1.TabIndex = 14;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(8, 222);
            progressBar1.Minimum = 1;
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(769, 23);
            progressBar1.Step = 1;
            progressBar1.Style = ProgressBarStyle.Continuous;
            progressBar1.TabIndex = 13;
            progressBar1.Value = 1;
            // 
            // launchButton2
            // 
            launchButton2.Location = new Point(676, 261);
            launchButton2.Name = "launchButton2";
            launchButton2.Size = new Size(101, 38);
            launchButton2.TabIndex = 12;
            launchButton2.Text = "Запуск";
            launchButton2.UseVisualStyleBackColor = true;
            launchButton2.Click += launchButton2_Click;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 10F);
            label6.ForeColor = Color.Red;
            label6.Location = new Point(196, 60);
            label6.Name = "label6";
            label6.Size = new Size(384, 19);
            label6.TabIndex = 11;
            label6.Text = "Не работай с оригиналом! Копируй файл на жесткий диск!";
            // 
            // annexesButton
            // 
            annexesButton.Location = new Point(676, 162);
            annexesButton.Name = "annexesButton";
            annexesButton.Size = new Size(101, 31);
            annexesButton.TabIndex = 10;
            annexesButton.Text = "Выбрать";
            annexesButton.UseVisualStyleBackColor = true;
            // 
            // textBox5
            // 
            textBox5.Font = new Font("Segoe UI", 10F);
            textBox5.Location = new Point(196, 168);
            textBox5.Name = "textBox5";
            textBox5.ReadOnly = true;
            textBox5.Size = new Size(458, 25);
            textBox5.TabIndex = 9;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(8, 168);
            label5.Name = "label5";
            label5.Size = new Size(134, 25);
            label5.TabIndex = 8;
            label5.Text = "Приложения А";
            // 
            // piktsButton
            // 
            piktsButton.Location = new Point(676, 98);
            piktsButton.Name = "piktsButton";
            piktsButton.Size = new Size(101, 31);
            piktsButton.TabIndex = 7;
            piktsButton.Text = "Выбрать";
            piktsButton.UseVisualStyleBackColor = true;
            // 
            // textBox4
            // 
            textBox4.Font = new Font("Segoe UI", 10F);
            textBox4.Location = new Point(196, 104);
            textBox4.Name = "textBox4";
            textBox4.ReadOnly = true;
            textBox4.Size = new Size(458, 25);
            textBox4.TabIndex = 6;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(8, 104);
            label4.Name = "label4";
            label4.Size = new Size(163, 25);
            label4.TabIndex = 5;
            label4.Text = "Таблица из ПИКТС";
            // 
            // moscowTableButton
            // 
            moscowTableButton.Location = new Point(676, 25);
            moscowTableButton.Name = "moscowTableButton";
            moscowTableButton.Size = new Size(101, 31);
            moscowTableButton.TabIndex = 4;
            moscowTableButton.Text = "Выбрать";
            moscowTableButton.UseVisualStyleBackColor = true;
            // 
            // textBox3
            // 
            textBox3.Font = new Font("Segoe UI", 10F);
            textBox3.Location = new Point(196, 32);
            textBox3.Name = "textBox3";
            textBox3.ReadOnly = true;
            textBox3.Size = new Size(458, 25);
            textBox3.TabIndex = 3;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(8, 31);
            label3.Name = "label3";
            label3.Size = new Size(182, 25);
            label3.TabIndex = 1;
            label3.Text = "Московская таблица";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(806, 358);
            Controls.Add(Tabcontrol);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Московские отчеты";
            Tabcontrol.ResumeLayout(false);
            Measures.ResumeLayout(false);
            Measures.PerformLayout();
            MoscowReport.ResumeLayout(false);
            MoscowReport.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TabControl Tabcontrol;
        private TabPage Measures;
        private TabPage MoscowReport;
        private Label label2;
        private Label label1;
        private TextBox textBox1;
        private Button launchButton1;
        private Button receiverButton;
        private TextBox textBox2;
        private Button sourceButton;
        private Label label6;
        private Button annexesButton;
        private TextBox textBox5;
        private Label label5;
        private Button piktsButton;
        private TextBox textBox4;
        private Label label4;
        private Button moscowTableButton;
        private TextBox textBox3;
        private Label label3;
        private Button launchButton2;
        private ProgressBar progressBar2;
        private ProgressBar progressBar1;
        private Label label7;
        private DateTimePicker dateTimePicker1;
        private Label moscPrgBarText;
        private Label label8;
        private CheckBox checkBox1;
        private Label measPrgBarText;
        private Label label10;
        private Label label9;
    }
}
