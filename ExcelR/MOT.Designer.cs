namespace ExcelR
{
    partial class MOT
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
            this.SaveMOT = new System.Windows.Forms.Button();
            this.FIOmot = new System.Windows.Forms.TextBox();
            this.DOLJNOSTmot = new System.Windows.Forms.TextBox();
            this.labFIO = new System.Windows.Forms.Label();
            this.labDOLJ = new System.Windows.Forms.Label();
            this.SpisokTamozh = new System.Windows.Forms.ComboBox();
            this.labTAMOJ = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Date_PO = new System.Windows.Forms.TextBox();
            this.Date_S = new System.Windows.Forms.TextBox();
            this.RezultKOLVO = new System.Windows.Forms.Label();
            this.KolVo = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.Calendar_S = new System.Windows.Forms.MonthCalendar();
            this.Calendar_PO = new System.Windows.Forms.MonthCalendar();
            this.test_combo = new System.Windows.Forms.Label();
            this.Save_Exit = new System.Windows.Forms.Button();
            this.Open_MOT = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataSet1 = new System.Data.DataSet();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.SuspendLayout();
            // 
            // SaveMOT
            // 
            this.SaveMOT.Enabled = false;
            this.SaveMOT.Location = new System.Drawing.Point(7, 468);
            this.SaveMOT.Name = "SaveMOT";
            this.SaveMOT.Size = new System.Drawing.Size(242, 23);
            this.SaveMOT.TabIndex = 0;
            this.SaveMOT.Text = "Записать данные";
            this.SaveMOT.UseVisualStyleBackColor = true;
            this.SaveMOT.Click += new System.EventHandler(this.SaveMOT_Click);
            // 
            // FIOmot
            // 
            this.FIOmot.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FIOmot.Location = new System.Drawing.Point(10, 39);
            this.FIOmot.Name = "FIOmot";
            this.FIOmot.Size = new System.Drawing.Size(210, 22);
            this.FIOmot.TabIndex = 1;
            // 
            // DOLJNOSTmot
            // 
            this.DOLJNOSTmot.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.DOLJNOSTmot.Location = new System.Drawing.Point(10, 90);
            this.DOLJNOSTmot.Name = "DOLJNOSTmot";
            this.DOLJNOSTmot.Size = new System.Drawing.Size(210, 22);
            this.DOLJNOSTmot.TabIndex = 2;
            // 
            // labFIO
            // 
            this.labFIO.AutoSize = true;
            this.labFIO.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labFIO.Location = new System.Drawing.Point(6, 16);
            this.labFIO.Name = "labFIO";
            this.labFIO.Size = new System.Drawing.Size(47, 20);
            this.labFIO.TabIndex = 3;
            this.labFIO.Text = "ФИО";
            // 
            // labDOLJ
            // 
            this.labDOLJ.AutoSize = true;
            this.labDOLJ.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labDOLJ.Location = new System.Drawing.Point(7, 66);
            this.labDOLJ.Name = "labDOLJ";
            this.labDOLJ.Size = new System.Drawing.Size(95, 20);
            this.labDOLJ.TabIndex = 4;
            this.labDOLJ.Text = "Должность";
            // 
            // SpisokTamozh
            // 
            this.SpisokTamozh.FormattingEnabled = true;
            this.SpisokTamozh.Items.AddRange(new object[] {
            "Брянская",
            "Владимирская",
            "Воронежская",
            "Курская",
            "Липецкая",
            "Московская",
            "Смоленская",
            "Тверская",
            "Тульская",
            "Ярославская",
            "Белгородская",
            "Калужская"});
            this.SpisokTamozh.Location = new System.Drawing.Point(7, 32);
            this.SpisokTamozh.Name = "SpisokTamozh";
            this.SpisokTamozh.Size = new System.Drawing.Size(242, 21);
            this.SpisokTamozh.TabIndex = 5;
            this.SpisokTamozh.SelectedIndexChanged += new System.EventHandler(this.SpisokTamozh_SelectedIndexChanged);
            // 
            // labTAMOJ
            // 
            this.labTAMOJ.AutoSize = true;
            this.labTAMOJ.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labTAMOJ.Location = new System.Drawing.Point(4, 9);
            this.labTAMOJ.Name = "labTAMOJ";
            this.labTAMOJ.Size = new System.Drawing.Size(152, 20);
            this.labTAMOJ.TabIndex = 6;
            this.labTAMOJ.Text = "Таможенный орган";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(1, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(20, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "С";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(113, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 20);
            this.label2.TabIndex = 10;
            this.label2.Text = "ПО";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Date_PO);
            this.groupBox1.Controls.Add(this.Date_S);
            this.groupBox1.Controls.Add(this.RezultKOLVO);
            this.groupBox1.Controls.Add(this.KolVo);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.Calendar_S);
            this.groupBox1.Controls.Add(this.Calendar_PO);
            this.groupBox1.Enabled = false;
            this.groupBox1.Location = new System.Drawing.Point(7, 191);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(242, 271);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Период командирования";
            // 
            // Date_PO
            // 
            this.Date_PO.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Date_PO.Location = new System.Drawing.Point(116, 44);
            this.Date_PO.Name = "Date_PO";
            this.Date_PO.Size = new System.Drawing.Size(100, 22);
            this.Date_PO.TabIndex = 14;
            this.Date_PO.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Date_PO_MouseClick);
            // 
            // Date_S
            // 
            this.Date_S.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Date_S.Location = new System.Drawing.Point(5, 44);
            this.Date_S.Name = "Date_S";
            this.Date_S.Size = new System.Drawing.Size(101, 22);
            this.Date_S.TabIndex = 13;
            this.Date_S.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Date_S_MouseClick);
            // 
            // RezultKOLVO
            // 
            this.RezultKOLVO.AutoSize = true;
            this.RezultKOLVO.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.RezultKOLVO.Location = new System.Drawing.Point(187, 76);
            this.RezultKOLVO.Name = "RezultKOLVO";
            this.RezultKOLVO.Size = new System.Drawing.Size(18, 20);
            this.RezultKOLVO.TabIndex = 12;
            this.RezultKOLVO.Text = "0";
            // 
            // KolVo
            // 
            this.KolVo.AutoSize = true;
            this.KolVo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.KolVo.Location = new System.Drawing.Point(5, 77);
            this.KolVo.Name = "KolVo";
            this.KolVo.Size = new System.Drawing.Size(124, 16);
            this.KolVo.TabIndex = 11;
            this.KolVo.Text = "Количество дней:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.labFIO);
            this.groupBox2.Controls.Add(this.FIOmot);
            this.groupBox2.Controls.Add(this.DOLJNOSTmot);
            this.groupBox2.Controls.Add(this.labDOLJ);
            this.groupBox2.Enabled = false;
            this.groupBox2.Location = new System.Drawing.Point(3, 55);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(246, 130);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Общие сведения";
            // 
            // Calendar_S
            // 
            this.Calendar_S.Location = new System.Drawing.Point(29, 104);
            this.Calendar_S.Name = "Calendar_S";
            this.Calendar_S.TabIndex = 13;
            this.Calendar_S.Visible = false;
            this.Calendar_S.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.Calendar_S_DateSelected);
            // 
            // Calendar_PO
            // 
            this.Calendar_PO.Location = new System.Drawing.Point(29, 104);
            this.Calendar_PO.Name = "Calendar_PO";
            this.Calendar_PO.TabIndex = 14;
            this.Calendar_PO.Visible = false;
            this.Calendar_PO.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.Calendar_PO_DateChanged);
            this.Calendar_PO.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.Calendar_PO_DateSelected);
            // 
            // test_combo
            // 
            this.test_combo.AutoSize = true;
            this.test_combo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.test_combo.Location = new System.Drawing.Point(566, 19);
            this.test_combo.Name = "test_combo";
            this.test_combo.Size = new System.Drawing.Size(0, 20);
            this.test_combo.TabIndex = 15;
            // 
            // Save_Exit
            // 
            this.Save_Exit.Enabled = false;
            this.Save_Exit.Location = new System.Drawing.Point(7, 526);
            this.Save_Exit.Name = "Save_Exit";
            this.Save_Exit.Size = new System.Drawing.Size(242, 23);
            this.Save_Exit.TabIndex = 16;
            this.Save_Exit.Text = "Сохранить и Выйти";
            this.Save_Exit.UseVisualStyleBackColor = true;
            this.Save_Exit.Click += new System.EventHandler(this.Save_Exit_Click);
            // 
            // Open_MOT
            // 
            this.Open_MOT.Location = new System.Drawing.Point(7, 497);
            this.Open_MOT.Name = "Open_MOT";
            this.Open_MOT.Size = new System.Drawing.Size(242, 23);
            this.Open_MOT.TabIndex = 17;
            this.Open_MOT.Text = "Открыть файл";
            this.Open_MOT.UseVisualStyleBackColor = true;
            this.Open_MOT.Click += new System.EventHandler(this.Open_MOT_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5});
            this.dataGridView1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.dataGridView1.Location = new System.Drawing.Point(272, 9);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(768, 540);
            this.dataGridView1.TabIndex = 18;
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column1.HeaderText = "ФИО";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Width = 59;
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column2.HeaderText = "Должность";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            this.Column2.Width = 90;
            // 
            // Column3
            // 
            this.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column3.HeaderText = "начало";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            this.Column3.Width = 67;
            // 
            // Column4
            // 
            this.Column4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column4.HeaderText = "конец";
            this.Column4.Name = "Column4";
            this.Column4.ReadOnly = true;
            this.Column4.Width = 62;
            // 
            // Column5
            // 
            this.Column5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column5.HeaderText = "дни";
            this.Column5.Name = "Column5";
            this.Column5.ReadOnly = true;
            this.Column5.Width = 50;
            // 
            // MOT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1048, 552);
            this.ControlBox = false;
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.Open_MOT);
            this.Controls.Add(this.labTAMOJ);
            this.Controls.Add(this.SaveMOT);
            this.Controls.Add(this.Save_Exit);
            this.Controls.Add(this.test_combo);
            this.Controls.Add(this.SpisokTamozh);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "MOT";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Командирование в МОТ";
            this.Load += new System.EventHandler(this.MOT_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SaveMOT;
        private System.Windows.Forms.TextBox FIOmot;
        private System.Windows.Forms.TextBox DOLJNOSTmot;
        private System.Windows.Forms.Label labFIO;
        private System.Windows.Forms.Label labDOLJ;
        private System.Windows.Forms.ComboBox SpisokTamozh;
        private System.Windows.Forms.Label labTAMOJ;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label RezultKOLVO;
        private System.Windows.Forms.Label KolVo;
        private System.Windows.Forms.TextBox Date_PO;
        private System.Windows.Forms.TextBox Date_S;
        private System.Windows.Forms.MonthCalendar Calendar_S;
        private System.Windows.Forms.MonthCalendar Calendar_PO;
        private System.Windows.Forms.Label test_combo;
        private System.Windows.Forms.Button Save_Exit;
        private System.Windows.Forms.Button Open_MOT;
        private System.Data.DataSet dataSet1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
    }
}