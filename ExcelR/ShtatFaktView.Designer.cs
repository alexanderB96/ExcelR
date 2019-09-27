namespace ExcelR
{
    partial class ShtatFaktView
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.load = new ExcelR.LoadingNew();
            this.ojidan = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1, 433);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(99, 15);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Location = new System.Drawing.Point(12, 246);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(227, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Сформировать";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // comboBox1
            // 
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
            "Ярославская таможня"});
            this.comboBox1.Location = new System.Drawing.Point(12, 30);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(227, 21);
            this.comboBox1.TabIndex = 2;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(209, 18);
            this.label1.TabIndex = 3;
            this.label1.Text = "Выбрать таможенный орган:";
            // 
            // load
            // 
            this.load.BackColor = System.Drawing.Color.Transparent;
            this.load.Location = new System.Drawing.Point(12, 125);
            this.load.Name = "load";
            this.load.Size = new System.Drawing.Size(235, 29);
            this.load.TabIndex = 4;
            this.load.Visible = false;
            // 
            // ojidan
            // 
            this.ojidan.AutoSize = true;
            this.ojidan.Location = new System.Drawing.Point(88, 109);
            this.ojidan.Name = "ojidan";
            this.ojidan.Size = new System.Drawing.Size(77, 13);
            this.ojidan.TabIndex = 5;
            this.ojidan.Text = "Ожидание . . .";
            this.ojidan.Visible = false;
            // 
            // ShtatFaktView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(247, 278);
            this.Controls.Add(this.ojidan);
            this.Controls.Add(this.load);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ShtatFaktView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Организационно-штатная структура";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.Button button2;
        public LoadingNew load;
        private System.Windows.Forms.Label ojidan;
    }
}