using System.Windows.Forms;

namespace ExcelR
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.ViborFaila = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.YkomplektovannostCTY = new System.Windows.Forms.Button();
            this.StatFact = new System.Windows.Forms.Button();
            this.SvodKad1 = new System.Windows.Forms.Button();
            this.Kad16Vigruzka = new System.Windows.Forms.Button();
            this.StatusDB = new System.Windows.Forms.Label();
            this.SelectDatee = new System.Windows.Forms.TextBox();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.label2 = new System.Windows.Forms.Label();
            this.KommandirovanieMOT = new System.Windows.Forms.Button();
            this.ShtatFaktNew = new System.Windows.Forms.Button();
            this.Form_6_2_CAD = new System.Windows.Forms.Button();
            this.loadingNew2 = new ExcelR.LoadingNew();
            this.loadingNew1 = new ExcelR.LoadingNew();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ViborFaila
            // 
            this.ViborFaila.Enabled = false;
            this.ViborFaila.Location = new System.Drawing.Point(12, 41);
            this.ViborFaila.Name = "ViborFaila";
            this.ViborFaila.Size = new System.Drawing.Size(211, 23);
            this.ViborFaila.TabIndex = 0;
            this.ViborFaila.Text = "Справка об укомплектованности v1";
            this.ViborFaila.UseVisualStyleBackColor = true;
            this.ViborFaila.Click += new System.EventHandler(this.ViborFaila_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(3, 512);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(639, 82);
            this.listBox1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 498);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Log:";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 70);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(211, 23);
            this.button2.TabIndex = 4;
            this.button2.Text = "Справка об укомплектованности v2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 99);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(211, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Должности";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            // 
            // YkomplektovannostCTY
            // 
            this.YkomplektovannostCTY.Location = new System.Drawing.Point(12, 12);
            this.YkomplektovannostCTY.Name = "YkomplektovannostCTY";
            this.YkomplektovannostCTY.Size = new System.Drawing.Size(211, 23);
            this.YkomplektovannostCTY.TabIndex = 7;
            this.YkomplektovannostCTY.Text = "Укомплектованность";
            this.YkomplektovannostCTY.UseVisualStyleBackColor = true;
            this.YkomplektovannostCTY.Click += new System.EventHandler(this.YkomplektovannostCTY_Click);
            // 
            // StatFact
            // 
            this.StatFact.Enabled = false;
            this.StatFact.Location = new System.Drawing.Point(12, 128);
            this.StatFact.Name = "StatFact";
            this.StatFact.Size = new System.Drawing.Size(211, 23);
            this.StatFact.TabIndex = 8;
            this.StatFact.Text = "Штат-Факт";
            this.StatFact.UseVisualStyleBackColor = true;
            this.StatFact.Click += new System.EventHandler(this.StatFact_Click);
            // 
            // SvodKad1
            // 
            this.SvodKad1.Location = new System.Drawing.Point(253, 12);
            this.SvodKad1.Name = "SvodKad1";
            this.SvodKad1.Size = new System.Drawing.Size(211, 23);
            this.SvodKad1.TabIndex = 9;
            this.SvodKad1.Text = "Общий КАД-1";
            this.SvodKad1.UseVisualStyleBackColor = true;
            this.SvodKad1.Click += new System.EventHandler(this.SvodKad1_Click);
            // 
            // Kad16Vigruzka
            // 
            this.Kad16Vigruzka.Location = new System.Drawing.Point(253, 41);
            this.Kad16Vigruzka.Name = "Kad16Vigruzka";
            this.Kad16Vigruzka.Size = new System.Drawing.Size(211, 23);
            this.Kad16Vigruzka.TabIndex = 11;
            this.Kad16Vigruzka.Text = "Выгрузка КАД-16";
            this.Kad16Vigruzka.UseVisualStyleBackColor = true;
            this.Kad16Vigruzka.Click += new System.EventHandler(this.Kad16Vigruzka_Click);
            // 
            // StatusDB
            // 
            this.StatusDB.AutoSize = true;
            this.StatusDB.Location = new System.Drawing.Point(2, 605);
            this.StatusDB.Name = "StatusDB";
            this.StatusDB.Size = new System.Drawing.Size(0, 13);
            this.StatusDB.TabIndex = 12;
            // 
            // SelectDatee
            // 
            this.SelectDatee.Location = new System.Drawing.Point(474, 28);
            this.SelectDatee.Name = "SelectDatee";
            this.SelectDatee.Size = new System.Drawing.Size(159, 20);
            this.SelectDatee.TabIndex = 14;
            this.SelectDatee.MouseClick += new System.Windows.Forms.MouseEventHandler(this.SelectDatee_MouseClick);
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.Location = new System.Drawing.Point(474, 49);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 13;
            this.monthCalendar1.Visible = false;
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(470, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(172, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Укажите дату для сбора КАД-16";
            // 
            // KommandirovanieMOT
            // 
            this.KommandirovanieMOT.Location = new System.Drawing.Point(12, 157);
            this.KommandirovanieMOT.Name = "KommandirovanieMOT";
            this.KommandirovanieMOT.Size = new System.Drawing.Size(211, 23);
            this.KommandirovanieMOT.TabIndex = 16;
            this.KommandirovanieMOT.Text = "Командирование в МОТ";
            this.KommandirovanieMOT.UseVisualStyleBackColor = true;
            this.KommandirovanieMOT.Click += new System.EventHandler(this.KommandirovanieMOT_Click);
            // 
            // ShtatFaktNew
            // 
            this.ShtatFaktNew.Location = new System.Drawing.Point(253, 128);
            this.ShtatFaktNew.Name = "ShtatFaktNew";
            this.ShtatFaktNew.Size = new System.Drawing.Size(211, 23);
            this.ShtatFaktNew.TabIndex = 17;
            this.ShtatFaktNew.Text = "ШтатФакт";
            this.ShtatFaktNew.UseVisualStyleBackColor = true;
            this.ShtatFaktNew.Click += new System.EventHandler(this.ShtatFaktNew_Click);
            // 
            // Form_6_2_CAD
            // 
            this.Form_6_2_CAD.Location = new System.Drawing.Point(253, 157);
            this.Form_6_2_CAD.Name = "Form_6_2_CAD";
            this.Form_6_2_CAD.Size = new System.Drawing.Size(211, 23);
            this.Form_6_2_CAD.TabIndex = 18;
            this.Form_6_2_CAD.Text = "2/6 KAD";
            this.Form_6_2_CAD.UseVisualStyleBackColor = true;
            this.Form_6_2_CAD.Click += new System.EventHandler(this.Form_6_2_CAD_Click);
            // 
            // loadingNew2
            // 
            this.loadingNew2.BackColor = System.Drawing.Color.Transparent;
            this.loadingNew2.Location = new System.Drawing.Point(407, 595);
            this.loadingNew2.Name = "loadingNew2";
            this.loadingNew2.Size = new System.Drawing.Size(235, 23);
            this.loadingNew2.TabIndex = 0;
            this.loadingNew2.Visible = false;
            // 
            // loadingNew1
            // 
            this.loadingNew1.BackColor = System.Drawing.Color.Transparent;
            this.loadingNew1.Location = new System.Drawing.Point(672, 3);
            this.loadingNew1.Name = "loadingNew1";
            this.loadingNew1.Size = new System.Drawing.Size(231, 16);
            this.loadingNew1.TabIndex = 10;
            this.loadingNew1.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(650, 618);
            this.Controls.Add(this.Form_6_2_CAD);
            this.Controls.Add(this.ShtatFaktNew);
            this.Controls.Add(this.KommandirovanieMOT);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SelectDatee);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.StatusDB);
            this.Controls.Add(this.loadingNew2);
            this.Controls.Add(this.Kad16Vigruzka);
            this.Controls.Add(this.SvodKad1);
            this.Controls.Add(this.StatFact);
            this.Controls.Add(this.YkomplektovannostCTY);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.ViborFaila);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExcelR";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private OpenFileDialog openFileDialog1;
        private Button ViborFaila;
        private Label label1;
        private Button button2;
        private Button button1;
        public ListBox listBox1;
        public System.ComponentModel.BackgroundWorker backgroundWorker1;
        private Button YkomplektovannostCTY;
        private Button StatFact;
        private Button SvodKad1;
        public LoadingNew loadingNew1;
        public Label StatusDB;
        private MonthCalendar monthCalendar1;
        public TextBox SelectDatee;
        public LoadingNew loadingNew2;
        public Button Kad16Vigruzka;
        private Label label2;
        private Button KommandirovanieMOT;
        private Button ShtatFaktNew;
        public Button Form_6_2_CAD;
    }
}

