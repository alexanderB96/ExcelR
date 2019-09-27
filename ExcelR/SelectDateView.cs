using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelR
{
    public partial class SelectDateView : UserControl
    {
        public SelectDateView()
        {
            InitializeComponent();
        }

        private void SelectDate_MouseClick(object sender, MouseEventArgs e)
        {
            Calend.Visible = true;
        }

        private void Calend_DateChanged(object sender, DateRangeEventArgs e)
        {
            SelectDate.Text = e.End.ToString("dd/MM/yyyy");
            Calend.Visible = false;
        }
    }
}
