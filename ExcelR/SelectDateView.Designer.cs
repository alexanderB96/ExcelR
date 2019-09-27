namespace ExcelR
{
    partial class SelectDateView
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

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.SelectDate = new System.Windows.Forms.ComboBox();
            this.Calend = new System.Windows.Forms.MonthCalendar();
            this.SuspendLayout();
            // 
            // SelectDate
            // 
            this.SelectDate.FormattingEnabled = true;
            this.SelectDate.Location = new System.Drawing.Point(3, 3);
            this.SelectDate.Name = "SelectDate";
            this.SelectDate.Size = new System.Drawing.Size(160, 21);
            this.SelectDate.TabIndex = 0;
            this.SelectDate.MouseClick += new System.Windows.Forms.MouseEventHandler(this.SelectDate_MouseClick);
            // 
            // Calend
            // 
            this.Calend.Location = new System.Drawing.Point(5, 25);
            this.Calend.Name = "Calend";
            this.Calend.TabIndex = 1;
            this.Calend.Visible = false;
            this.Calend.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.Calend_DateChanged);
            // 
            // SelectDateView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.Calend);
            this.Controls.Add(this.SelectDate);
            this.Name = "SelectDateView";
            this.Size = new System.Drawing.Size(166, 191);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ComboBox SelectDate;
        private System.Windows.Forms.MonthCalendar Calend;
    }
}
