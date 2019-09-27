namespace ExcelR
{
    partial class CAD_2_6
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.ojidan = new System.Windows.Forms.Label();
            this.CAD_6 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.load = new ExcelR.LoadingNew();
            this.PeriodS = new System.Windows.Forms.MonthCalendar();
            this.PeriodPo = new System.Windows.Forms.MonthCalendar();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(209, 18);
            this.label1.TabIndex = 5;
            this.label1.Text = "Выбрать таможенный орган:";
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Аппарат управления ЦТУ",
            "Белгородская таможня",
            "Брянская таможня",
            "Владимирская таможня",
            "Воронежская таможня",
            "Калужская таможня",
            "Курская таможня",
            "Липецкая таможня",
            "Московская таможня",
            "Приокский ТТП",
            "Смоленская таможня",
            "Тверская таможня",
            "Тульская таможня",
            "ЦОТ",
            "Ярославская таможня",
            "ЦТУ"});
            this.comboBox1.Location = new System.Drawing.Point(15, 30);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(350, 24);
            this.comboBox1.TabIndex = 4;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // ojidan
            // 
            this.ojidan.AutoSize = true;
            this.ojidan.Location = new System.Drawing.Point(2, 333);
            this.ojidan.Name = "ojidan";
            this.ojidan.Size = new System.Drawing.Size(77, 13);
            this.ojidan.TabIndex = 7;
            this.ojidan.Text = "Ожидание . . .";
            this.ojidan.Visible = false;
            // 
            // CAD_6
            // 
            this.CAD_6.Enabled = false;
            this.CAD_6.Location = new System.Drawing.Point(15, 277);
            this.CAD_6.Name = "CAD_6";
            this.CAD_6.Size = new System.Drawing.Size(93, 23);
            this.CAD_6.TabIndex = 8;
            this.CAD_6.Text = "6_KAD";
            this.CAD_6.UseVisualStyleBackColor = true;
            this.CAD_6.Click += new System.EventHandler(this.CAD_6_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(531, 243);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "label2";
            this.label2.Visible = false;
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.BackColor = System.Drawing.SystemColors.Control;
            this.checkedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Items.AddRange(new object[] {
            "Текущая неделя",
            "Предыдущая неделя",
            "Текущий месяц",
            "Предыдущий месяц",
            "Текущий квартал",
            "Прошлый квартал",
            "Текущий год",
            "Прошлый год"});
            this.checkedListBox1.Location = new System.Drawing.Point(560, 131);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(137, 122);
            this.checkedListBox1.TabIndex = 11;
            this.checkedListBox1.Visible = false;
            this.checkedListBox1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBox1_ItemCheck);
            // 
            // load
            // 
            this.load.BackColor = System.Drawing.Color.Transparent;
            this.load.Location = new System.Drawing.Point(136, 327);
            this.load.Name = "load";
            this.load.Size = new System.Drawing.Size(235, 29);
            this.load.TabIndex = 6;
            this.load.Visible = false;
            // 
            // PeriodS
            // 
            this.PeriodS.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.PeriodS.Location = new System.Drawing.Point(7, 17);
            this.PeriodS.Name = "PeriodS";
            this.PeriodS.TabIndex = 12;
            this.PeriodS.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.PeriodS_DateSelected);
            // 
            // PeriodPo
            // 
            this.PeriodPo.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.PeriodPo.Location = new System.Drawing.Point(6, 17);
            this.PeriodPo.Name = "PeriodPo";
            this.PeriodPo.TabIndex = 13;
            this.PeriodPo.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.PeriodPo_DateSelected);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox1.Location = new System.Drawing.Point(15, 58);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(356, 216);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Период";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.PeriodS);
            this.groupBox2.Location = new System.Drawing.Point(6, 19);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(169, 188);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "С:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.PeriodPo);
            this.groupBox3.Location = new System.Drawing.Point(181, 19);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(169, 188);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "ПО:";
            // 
            // CAD_2_6
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(374, 355);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.CAD_6);
            this.Controls.Add(this.ojidan);
            this.Controls.Add(this.load);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "CAD_2_6";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CAD_2_6";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label ojidan;
        public LoadingNew load;
        private System.Windows.Forms.Button CAD_6;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        public System.Windows.Forms.Label label2;
        private System.Windows.Forms.MonthCalendar PeriodS;
        private System.Windows.Forms.MonthCalendar PeriodPo;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}