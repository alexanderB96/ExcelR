using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ExcelR
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            AUtoriz utoriz = new AUtoriz();

            if (!((userName == "REGIONS\\BobylevAS") | (userName == "REGIONS\\NesterkinaOM")))
            {
                utoriz.ShowDialog();
            }
            

            Application.Run(new Form1());

            
        }
    }
}
