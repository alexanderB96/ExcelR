using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace ExcelR
{
    public partial class MOT : Form
     
    {

        TimeSpan data_rezult;
        DateTime data_s, data_po;
        public string TamojnOrgan;
        public int LastRow;
        public string addr;
        private int ls;

        public MOT()
        {
            InitializeComponent();
        }

        private void MOT_Load(object sender, EventArgs e)
        {

        }

        private void Date_S_MouseClick(object sender, MouseEventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                Calendar_S.Visible = true;
            }));
        }

        private void Date_PO_MouseClick(object sender, MouseEventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                Calendar_PO.Visible = true;
            }));
        }

        private void Calendar_S_DateSelected(object sender, DateRangeEventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                Date_S.Text = e.End.ToString("dd/MM/yyyy");
                Calendar_S.Visible = false;
                data_s = DateTime.Parse(Date_S.Text);

                if (Date_PO.TextLength != 0)
                {

                        data_rezult = data_po - data_s;
                        var data_rez = Convert.ToInt32(data_rezult.Days.ToString());
                        data_rez++;
                        RezultKOLVO.Text = Convert.ToString(data_rez);

                }

                
            }));
           
        }

        private void SpisokTamozh_SelectedIndexChanged(object sender, EventArgs e)
        {
            test_combo.Text = SpisokTamozh.Text;
            TamojnOrgan = SpisokTamozh.Text;
            test_combo.Visible = false;
            groupBox1.Enabled = true;
            groupBox2.Enabled = true;
            SaveMOT.Enabled = true;
            Save_Exit.Enabled = true;

            if ((TamojnOrgan == "Брянская") == true) //Брянск
            {
                LoadDataGrid(1);
                ls = 1;
            }

            if ((TamojnOrgan == "Владимирская") == true) //
            {
                LoadDataGrid(2);
                ls = 2;
            }

            if ((TamojnOrgan == "Воронежская") == true) //
            {
                LoadDataGrid(3);
                ls = 3;
            }

                if ((TamojnOrgan == "Курская") == true) //
                {
                LoadDataGrid(4);
                ls = 4;
            }

                if ((TamojnOrgan == "Липецкая") == true) //
                {
                LoadDataGrid(5);
                ls = 5;
            }

                if ((TamojnOrgan == "Московская") == true) //
                {
                LoadDataGrid(6);
                ls = 6;
            }

                if ((TamojnOrgan == "Смоленская") == true) //
                {
                LoadDataGrid(7);
                ls = 7;
            }

                if ((TamojnOrgan == "Тверская") == true) //
                {
                LoadDataGrid(8);
                ls = 8;
            }

                if ((TamojnOrgan == "Тульская") == true) //
                {
                LoadDataGrid(9);
                ls = 9;
            }

                if ((TamojnOrgan == "Ярославская") == true) //
                {
                LoadDataGrid(10);
                ls = 10;
            }

                if ((TamojnOrgan == "Белгородская") == true) //
                {
                LoadDataGrid(11);
                ls = 11;
            }

                if ((TamojnOrgan == "Калужская") == true) //
                {
                LoadDataGrid(12);
                ls = 12;
            }

            
        }

        private void Save_Exit_Click(object sender, EventArgs e)
        {
            foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
            {
                currentProcess.Kill();

            }
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SaveMOT_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //сам эксель
                string pathSaveSpisokLicz = String.Format(@"D:\Справки\Список_лиц_командированных_в_МОТ");
                string FilNam = (AppDomain.CurrentDomain.BaseDirectory + @"Templates\Список_лиц_командированных_в_МОТ.xlsx");
                Excel.Worksheet WS;// в этот лист
                Workbook WB = ObjWorkExcel.Workbooks.Open(Convert.ToString(FilNam));
                int RowCount = 0;
                WS = (Excel.Worksheet)WB.Sheets[1];
                WS.Activate();
                ObjWorkExcel.DisplayAlerts = false;
                //Считывание с определенных ячеек
                string NameTamozn = Convert.ToString(WS.Range["A1"].Value);


                if ((TamojnOrgan == "Брянская") == true) //Брянск
                {
                    WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                    LastRow = WB.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Владимирская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 2 лист
                    LastRow = WB.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Воронежская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 3 лист
                    LastRow = WB.Sheets[3].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Курская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 4  лист
                    LastRow = WB.Sheets[4].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Липецкая") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 4  лист
                    LastRow = WB.Sheets[5].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Московская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 4  лист
                    LastRow = WB.Sheets[6].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Смоленская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[7];//выбираем 4  лист
                    LastRow = WB.Sheets[7].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Тверская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[8];//выбираем 4  лист
                    LastRow = WB.Sheets[8].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Тульская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[9];//выбираем 4  лист
                    LastRow = WB.Sheets[9].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Ярославская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[10];//выбираем 4  лист
                    LastRow = WB.Sheets[10].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Белгородская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[11];//выбираем 4  лист
                    LastRow = WB.Sheets[11].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                if ((TamojnOrgan == "Калужская") == true) //
                {
                    WS = (Excel.Worksheet)WB.Sheets[12];//выбираем 4  лист
                    LastRow = WB.Sheets[12].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert();
                    }


                    WS.Cells[LastRow, 1] = FIOmot.Text.ToString();
                    WS.Cells[LastRow, 2] = Convert.ToString(DOLJNOSTmot.Text.ToString());
                    WS.Cells[LastRow, 3] = Convert.ToString(Date_S.Text.ToString());
                    WS.Cells[LastRow, 4] = Convert.ToString(Date_PO.Text.ToString());
                    WS.Cells[LastRow, 5] = Convert.ToString(RezultKOLVO.Text);
                }

                WB.SaveAs(pathSaveSpisokLicz, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                WB.Close(true);
                File.Copy(@"D:\Справки\Список_лиц_командированных_в_МОТ.xlsx", AppDomain.CurrentDomain.BaseDirectory + @"Templates\Список_лиц_командированных_в_МОТ.xlsx", true);
                ObjWorkExcel.Quit();

                int list = ls;
                
                FIOmot.Text = "";
                DOLJNOSTmot.Text = "";
                Date_S.Text = "";
                Date_PO.Text = "";
                RezultKOLVO.Text = "";

                LoadDataGrid(list);



            }

            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           

        }

        private void Open_MOT_Click(object sender, EventArgs e)
        {
            try
            {
                string str;
                int rCnt, cCnt;
                int RowCount = 0;
                int list = 1;
                Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //сам эксель

                string FilNam = (AppDomain.CurrentDomain.BaseDirectory + @"Templates\Список_лиц_командированных_в_МОТ.xlsx");
                Excel.Worksheet WS;
                Workbook WB = ObjWorkExcel.Workbooks.Open(Convert.ToString(FilNam));


                // Excel.Range ExcelRange;

                WS = (Excel.Worksheet)WB.Sheets[list];

                // int LastRow = WB.Sheets[list].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                // string Name = "A4:E" + LastRow.ToString(); //берем диапазон ячеек от а1 до еP+последняя строка
                // string Adress = "A4" + ":E" + (RowCount + LastRow - 1).ToString();//формируем адрес куда в существующей книге мы это скопируем

                WS.Activate();
                ObjWorkExcel.DisplayAlerts = false;
                ObjWorkExcel.Visible = true;

                List<string> name = new List<string> { "EXCEL" };//процесс, который нужно убить
                System.Diagnostics.Process[] etc = System.Diagnostics.Process.GetProcesses();//получим процессы
                foreach (System.Diagnostics.Process anti in etc)//обойдем каждый процесс
                {
                    foreach (string s in name)
                    {
                        if (anti.ProcessName.ToLower().Contains(s.ToLower())) //найдем нужный и убьем
                        {
                            anti.Kill();
                            name.Remove(s);
                        }
                    }
                }
                /*

                System.Data.DataTable tb = new System.Data.DataTable();
                ExcelRange = WS.get_Range(Name);



                for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
                {
                    dataGridView1.Rows.Add(1);
                    for(cCnt = 1; cCnt<=5;cCnt++)
                    {
                        str = Convert.ToString((ExcelRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                }
                */
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
          
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void Calendar_PO_DateSelected(object sender, DateRangeEventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                Date_PO.Text = e.End.ToString("dd/MM/yyyy");
                Calendar_PO.Visible = false;
                data_po = DateTime.Parse(Date_PO.Text);

                if (Date_S.TextLength != 0)
                {
                    
                        data_rezult = data_po - data_s;
                        var data_rez = Convert.ToInt32(data_rezult.Days.ToString());
                        data_rez++;
                        RezultKOLVO.Text =Convert.ToString(data_rez);

                    
                }
            }));

            
        }

        private void Calendar_PO_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {/*
            var rez = Convert.ToInt32(RezultKOLVO.Text);
            if ((checkBox1.Checked)==true)
            {
                if (rez > 0)
                {
                    rez++;
                    var rr = rez;
                    RezultKOLVO.Text = Convert.ToString(rr);
                }
               
            }

            if ((checkBox1.Checked) == false)
            {
                if (rez <= 0)
                {
                    RezultKOLVO.Text = Convert.ToString(rez);
                }

            }*/
        }

        private void LoadDataGrid(int list)
        {
            try
            {
                dataGridView1.Rows.Clear();
                string str;
                int rCnt, cCnt;
                int RowCount = 0;
                //int list = 1;
                Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //сам эксель

                string FilNam = (AppDomain.CurrentDomain.BaseDirectory + @"Templates\Список_лиц_командированных_в_МОТ.xlsx");
                Excel.Worksheet WS;
                Workbook WB = ObjWorkExcel.Workbooks.Open(Convert.ToString(FilNam));
                Excel.Range ExcelRange;


                WS = (Excel.Worksheet)WB.Sheets[list];
                int LastRow = WB.Sheets[list].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                string Name = "A4:E" + LastRow.ToString(); //берем диапазон ячеек от а1 до еP+последняя строка
                                                           // string Adress = "A4" + ":E" + (RowCount + LastRow - 1).ToString();//формируем адрес куда в существующей книге мы это скопируем

                WS.Activate();
                ObjWorkExcel.DisplayAlerts = false;
                ObjWorkExcel.Visible = false;




                System.Data.DataTable tb = new System.Data.DataTable();
                ExcelRange = WS.get_Range(Name);



                for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
                {
                    dataGridView1.Rows.Add(1);
                    for (cCnt = 1; cCnt <= 5; cCnt++)
                    {
                        str = Convert.ToString((ExcelRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                }

                ObjWorkExcel.Quit();

                
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
                       
            
           
        }
    }
}
