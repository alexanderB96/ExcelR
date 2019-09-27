using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelR
{
    public partial class ConnectDB : Form
    {
        public ConnectDB()
        {
            InitializeComponent();
        }

        private void Connect_Click(object sender, EventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                Form1 fr = new Form1();
            try
            {
                DBOracleUtils db = new DBOracleUtils();
                var login = Convert.ToString(LoginDB);
                var pass = Convert.ToString(PassBD);
                OracleConnection conn = db.GetDBConnection();
                conn.Open();
            }
            catch (Exception s)
            {
                    MessageBox.Show(s.Message,
                         "Error",
                         MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

            finally
            {
                    //MessageBox.Show("Успешно!");
                fr.StatusDB.Text = "Успешное подключение!";
                this.Close();
            }
            }));
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
