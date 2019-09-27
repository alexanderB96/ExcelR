using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelR
{
    public partial class Form1 : Form
    {
        private Stream myStream;
        public string FName;
        public string dat;
        
        DateTime Date1 = DateTime.Now;
        public delegate void InvokeDelegate();

        public double[,] arr = new double[15, 3];
        public double[,] arrNew = new double[15, 20];
        public string[] Podrazd = new string[15] {  "=*Аппа**", "=*Бел*" ,"=*Брян*" ,"=*Влад*" ,"=*Вор*" ,"=*Кал*" ,"=*Кур*" ,"=*Лип*" ,"=*Мос*"  ,"=*Смо*","=*Тве*" ,"=*Тул*" ,"=*Яро*" ,"**Центральная**" ,"**Приокск*" };

        public double[] arrShtat = new double[15];
        public double[] arrFakt = new double[15];
        public double[] arrVakan = new double[15];

        public double[] arrShtatSotr = new double[15];
        public double[] arrShtatGos = new double[15];
        public double[] arrShtatRab = new double[15];

        public double[] arrFaktSotr = new double[15];
        public double[] arrFaktGos = new double[15];
        public double[] arrFaktRab = new double[15];

        public double[] arrVakansSotr = new double[15];
        public double[] arrVakansGos = new double[15];
        public double[] arrVakansRab = new double[15];
        public Form1()
        {
            InitializeComponent();
        }

        public void ViborFaila_Click(object sender, EventArgs e)
        {
            try
            {

                OpenFileDialog dlg = new OpenFileDialog
                {
                    Multiselect = true,
                    Title = "Выберите файлы",
                    InitialDirectory = @"D:\",
                    Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                };
                dlg.ShowDialog();
                if (dlg.FileName == String.Empty)
                    return;
                var exePath = AppDomain.CurrentDomain.BaseDirectory;//path to exe file
                var path = Path.Combine(exePath, "Templates\\Укомплектованность.xls");

                Microsoft.Office.Interop.Excel.Application xlsApp1 = new Microsoft.Office.Interop.Excel.Application();//Точно не знаю надо ли заново создавать 
                Workbook xlsBookOpen = xlsApp1.Workbooks.Open(path);//, ReadOnly: true); //рабочая книга

                foreach (string file in dlg.FileNames)
                {
                    /* listBox1.Items.Add(file);*/
                    var ch = "KAD1";
                    var indexOfChar = file.IndexOf(ch); // равно 4
                    if (indexOfChar < 1)
                    {
                        
                        var FileName = file;
                        listBox1.Items.Add(FileName);
                        YkomlektDop(FileName);

                    }
                                        
                    //string[] lines = Convert.ToString[](file);
                }

                foreach (string file in dlg.FileNames)
                {
                    


                    var ch = "KAD1";
                    var indexOfChar = file.IndexOf(ch); // равно 4*/
                    if (indexOfChar > 1)
                    {
                        FName = file;
                        listBox1.Items.Add(file);
                        Ykomlekt(FName, xlsBookOpen);

                    }
                    
                    //string[] lines = Convert.ToString[](file);
                }

                Directory.CreateDirectory(@"D:\Справки");
                string pathSave = String.Format(@"D:\Справки\Справка_Об_Укомплектованности_{0}_{1}_{2}_{3}_{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                xlsBookOpen.SaveAs(pathSave, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlsApp1.Quit();
               
                MessageBox.Show("Расположен по адресу: \nD:\\Справки\\", "Файл успешно сформирован");
            }

            
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
                        
            foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
            {
                currentProcess.Kill();
            }
        }

        private void Ykomlekt(string FName, Workbook xlsBookOpen)
        {
            Form1 frm = new Form1();
            OpenFileDialog dlg = new OpenFileDialog();
            //считывание с Файла 1
            // Открытие книги
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlsBook = xlsApp.Workbooks.Open(Convert.ToString(FName));//, ReadOnly: false); //рабочая книга
            Worksheet xlsSheet;
            xlsSheet = (Worksheet)xlsBook.Sheets[1]; // раздел 1 (лист 1)
            xlsSheet.Activate();
            //Считывание с определенных ячеек
            var StructPodrazdelenie = xlsSheet.Range["B19"].Value;
            //var DataF = xlsSheet.Range["C22"].Value; //берём дату из файлов
            listBox1.Items.Add(StructPodrazdelenie);
            listBox1.TopIndex = listBox1.Items.Count - 1; // log следует за добавлениями файлов
            xlsSheet = (Worksheet)xlsBook.Sheets[4]; // раздел 4 (лист 4)
            xlsSheet.Activate();
            var NachYpr = xlsSheet.Range["G12"].Value; //начальник управления (таможни)
            var PervyiZNY = xlsSheet.Range["J12"].Value;//Первый ЗНУ
            var ZamNY = xlsSheet.Range["M12"].Value; //Заместитель НУ
            var ZNY = xlsSheet.Range["Z12"].Value; // ЗНУ
            var NachSlyg = xlsSheet.Range["AJ12"].Value; // Начальник службы
            var ZNSlyg = xlsSheet.Range["AR12"].Value; // ЗН Службы
            var NachPosta = xlsSheet.Range["AV12"].Value; // Начальник поста
            var ZNPosta = xlsSheet.Range["AY12"].Value; // ЗН Поста
            //
            var NO1 = xlsSheet.Range["BA12"].Value; // НО (отделения) для суммы
            var NO2 = xlsSheet.Range["BG12"].Value;// НО (отделения)  для суммы
            var NO = NO1 + NO2; // НО (отделения) сумма 
            //
            var ZNO = xlsSheet.Range["BE12"].Value; // ЗНО
            //
            var IniyDolgnosti1 = xlsSheet.Range["BK12"].Value;
            var IniyDolgnosti2 = xlsSheet.Range["BM12"].Value;
            var IniyDolgnosti3 = xlsSheet.Range["BO12"].Value;
            var IniyDolgnosti4 = xlsSheet.Range["BQ12"].Value;
            var IniyDolgnosti5 = xlsSheet.Range["BS12"].Value;
            var IniyDolgnosti6 = xlsSheet.Range["P12"].Value;

            
            //
            var IniyDolgnosti = IniyDolgnosti1 + IniyDolgnosti2 + IniyDolgnosti3 + IniyDolgnosti4 + IniyDolgnosti5 + IniyDolgnosti6;// Иные должности
            var OpTech = 0;  // Оперативно-технические (все должности)
            //
            var NOPk1 = xlsSheet.Range["BB12"].Value;
            var NOPk2 = xlsSheet.Range["BH12"].Value;
            var NOPk = NOPk1 + NOPk2; // НО, отделения (противодействия коррупции)
            //
            var IniyDolgPK = 0; // Иные должности (противодействия коррупции)
            xlsSheet = (Worksheet)xlsBook.Sheets[5]; //раздел 5 (лист 5)
            xlsSheet.Activate();

            var NachPostaFGGS = xlsSheet.Range["P13"].Value; ; //Начальник поста ФГГС 
            var ZNPostaFGGS = xlsSheet.Range["S13"].Value; ; //ЗН Поста ФГГС

            var NOFGGS1 = xlsSheet.Range["Y13"].Value; ; // НО отделения ФГГС
            var NOFGGS2 = xlsSheet.Range["AC13"].Value; ; // НО отделения ФГГС
            var NOFGGS = NOFGGS1+ NOFGGS2; // НО отделения ФГГС

            var ZNOFGGS = xlsSheet.Range["AA13"].Value; ; // ЗНО
            
            var IniyDoljnostiFGGS1 = xlsSheet.Range["AE13"].Value; // Иные должности ФГГС
            var IniyDoljnostiFGGS2 = xlsSheet.Range["AG13"].Value; // Иные должности ФГГС
            var IniyDoljnostiFGGS3 = xlsSheet.Range["AI13"].Value; // Иные должности ФГГС
            var IniyDoljnostiFGGS4 = xlsSheet.Range["AK13"].Value; // Иные должности ФГГС
            var IniyDoljnostiFGGS5 = xlsSheet.Range["AM13"].Value; // Иные должности ФГГС
            var IniyDoljnostiFGGS6 = xlsSheet.Range["AO13"].Value; // Иные должности ФГГС
            var IniyDoljnostiFGGS7 = xlsSheet.Range["AQ13"].Value; // Иные должности ФГГС
            var IniyDoljnostiFGGS = IniyDoljnostiFGGS1 + IniyDoljnostiFGGS2+ IniyDoljnostiFGGS3+ IniyDoljnostiFGGS4+ IniyDoljnostiFGGS5 + IniyDoljnostiFGGS6 + IniyDoljnostiFGGS7; // Иные должности ФГГС

            xlsSheet = (Worksheet)xlsBook.Sheets[6];// раздел 6 (лист 6)
            xlsSheet.Activate();
            var Rabotniki = Convert.ToDouble(xlsSheet.Range["E13"].Value); //Работники

            xlsSheet = (Worksheet)xlsBook.Sheets[3];// раздел 6 (лист 6)
            xlsSheet.Activate();
            var PoShtaty1 = Convert.ToDouble(xlsSheet.Range["C12"].Value); //По штату
            var PoShtaty2 = Convert.ToDouble(xlsSheet.Range["J12"].Value); //По штату
            var PoShtaty3 = Convert.ToDouble(xlsSheet.Range["N12"].Value); //По штату
            var PoShtaty = PoShtaty1 + PoShtaty2 + PoShtaty3; //По штату

            var PoFakty1 = Convert.ToDouble(xlsSheet.Range["D12"].Value); //По факту
            var PoFakty2 = Convert.ToDouble(xlsSheet.Range["K12"].Value); //По факту
            var PoFakty3 = Convert.ToDouble(xlsSheet.Range["O12"].Value); //По факту
            var PoFakty = PoFakty1 + PoFakty2 + PoFakty3; //По факту
            //listBox1.Items.Add(NachYpr + "," + PervyiZNY + "," + ZamNY + "," + ZNY + "," + NachSlyg + "," + ZNSlyg + "," + NachPosta + "," + ZNPosta + "," + NO + "," + ZNO + "," + IniyDolgnosti + "," + OpTech + "," + NOPk + "," + IniyDolgPK);
           
            //Закрытие файла
             xlsBook.Close(0);
             xlsApp.Quit();
             GC.Collect();

            

            xlsSheet = (Worksheet)xlsBookOpen.Sheets[1];
            xlsSheet.Activate();
            int structPodr = 100;
            int i = 0;
            if (StructPodrazdelenie == "Аппарат Центрального таможенного управления") //Синий
            {
                structPodr = 5;
                i = 0;
               
            }
            if (StructPodrazdelenie == "Белгородская таможня") // жёлтый
            {
                structPodr = 6;
                i = 1;
            }
            if (StructPodrazdelenie == "Брянская таможня") //Синий
            {
                structPodr = 7;
                i = 2;
            }
            if (StructPodrazdelenie == "Владимирская таможня") // жёлтый
            {
                structPodr = 8;
                i = 3;
            }
            if (StructPodrazdelenie == "Воронежская таможня") //Синий
            {
                structPodr = 9;
                i = 4;
            }
            if (StructPodrazdelenie == "Калужская таможня") // жёлтый
            {
                structPodr = 10;
                i = 5;
            }
            if (StructPodrazdelenie == "Курская таможня") //Синий
            {
                structPodr = 11;
                i = 6;
            }
            if (StructPodrazdelenie == "Липецкая таможня") // жёлтый
            {
                structPodr = 12;
                i = 7;
            }
            if (StructPodrazdelenie == "Московская таможня") //Синий
            {
                structPodr = 13;
                i = 8;
            }
            if (StructPodrazdelenie == "Смоленская  таможня") // жёлтый
            {
                structPodr = 14;
                i = 9;
            }
            if (StructPodrazdelenie == "Тверская  таможня") //Синий
            {
                structPodr = 15;
                i = 10;
            }
            if (StructPodrazdelenie == "Тульская  таможня") // жёлтый
            {
                structPodr = 16;
                i = 11;
            }
            if (StructPodrazdelenie == "Ярославская таможня") // жёлтый
            {
                structPodr = 17;
                i = 12;
            }
            if (StructPodrazdelenie == "Центральная оперативная таможня") //Синий
            {
                structPodr = 18;
                i = 13;
            }
            if (StructPodrazdelenie == "Приокский тыловой таможенный пост") // жёлтый
            {
                structPodr = 19;
                i = 14;
            }
        

            xlsSheet.Cells[structPodr, 3] = NachYpr; // НАЧАЛЬНИК УПРАВЛЕНИЯ
            xlsSheet.Cells[structPodr, 4].value = PervyiZNY; // Первый заместитель
            xlsSheet.Cells[structPodr, 5].value = ZamNY; // Заместитель НУ
            xlsSheet.Cells[structPodr, 6].value = arr[i,0]; // ЗНУ
            xlsSheet.Cells[structPodr, 7].value = NachSlyg; //Начальник службы
            xlsSheet.Cells[structPodr, 8].value = ZNSlyg; // Заместитель начальника службы
            xlsSheet.Cells[structPodr, 9].value = NachPosta; // Начальник поста
            xlsSheet.Cells[structPodr, 10].value = ZNPosta; //Заместитель начальника поста
            xlsSheet.Cells[structPodr, 11].value = NO; //НО (отделения)
            xlsSheet.Cells[structPodr, 12].value = ZNO; // ЗНО
            xlsSheet.Cells[structPodr, 13].value = IniyDolgnosti; // Иные должности
            xlsSheet.Cells[structPodr, 14].value = arr[i,1]; //Оперативно-технические (все должности)
            xlsSheet.Cells[structPodr, 15].value = NOPk; // НО противКорруп
            xlsSheet.Cells[structPodr, 16].value= arr[i,2];// Иные должности ПротивКорруп
            xlsSheet.Cells[structPodr, 17].value= NachPostaFGGS ;// Начальник поста ФГГС
            xlsSheet.Cells[structPodr, 18].value =  ZNPostaFGGS ;// Заместитель начальника поста ФГГС
            xlsSheet.Cells[structPodr, 19].value = NOFGGS; // НО ФГГС
            xlsSheet.Cells[structPodr, 20].value = ZNOFGGS ; // ЗНО ФГГС
            xlsSheet.Cells[structPodr, 21].value = IniyDoljnostiFGGS; // Иные должности ФГГС
            xlsSheet.Cells[structPodr, 22].value = Rabotniki; // Работники ФГГС
            xlsSheet.Cells[structPodr, 24].value = PoShtaty; // По штату
            xlsSheet.Cells[structPodr, 25].value = PoFakty; // По факту
            
           // xlsSheet.Cells[1, 2].value = "Справка об укомплектованности Центрального таможенного управления " + DataF; // По факту
        }


        private void YkomlektDop(string FileName)
        {
            try
            {
              
                Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlsBook = xlsApp.Workbooks.Open(Convert.ToString(FileName));//, ReadOnly: false); //рабочая книга
                Worksheet xlsSheet;
                xlsSheet = (Worksheet)xlsBook.Sheets[1]; // раздел 1 (лист 1)
                xlsSheet.Activate();

                var MyRange = "E2:E1400";

                //разделяем  столбец Кол-во/Дата Образования
                xlsSheet.Range["E2:E1400"].TextToColumns(xlsSheet.get_Range(MyRange, Type.Missing), XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierDoubleQuote, true, Type.Missing, Type.Missing, true, false, Type.Missing, ",", Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Directory.CreateDirectory(@"D:\test");
                int i = 0;
                int ii = 0;
                do
                {
                    string pathSave = String.Format(@"D:\test\{0}_test_{1}_{2}_{3}_{4}_{5}_{6}_{7}_{8}", i, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond, ii);
                    if (i == 0) //Аппарат
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Це* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=*Центр*", XlAutoFilterOperator.xlAnd);
                        var ZNY = xlsSheet.Range["E1411"].Value;
                        arr[i, 0] = ZNY;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        ii++;

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTech = xlsSheet.Range["E1411"].Value;
                        arr[i, 1] = OperatTech;
                        //xlsBook.SaveAs(pathSave, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        ii++;

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Центр**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnosti = xlsSheet.Range["E1411"].Value;
                        arr[i, 2] = IniyDoljnosti;
                        //xlsBook.SaveAs(pathSave, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        ii++;
                        //xlsApp.Quit();
                    }

                    if (i == 1) //Белгород
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Бел* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Белг**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        arr[i, 0] = ZNYb;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        arr[i, 1] = OperatTechb;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Белг**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        arr[i, 2] = IniyDoljnostib;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        //xlsApp.Quit();
                    }
                    if (i == 2) //Брянск
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Бря* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Бря**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;

                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=*отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;

                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Бря**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;

                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        ii++;
                        // xlsApp.Quit();
                    }
                    if (i == 3) // Владимир
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Влад** - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Влад**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Влад**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        //xlsApp.Quit();
                    }
                    if (i == 4) //Воронеж
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Ворон* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ворон**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ворон**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();
                    }
                    if (i == 5) //Калуга
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Калу* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Калу**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Калу**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();
                    }
                    if (i == 6) //Курск
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Кур* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Кур**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Кур**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();

                    }
                    if (i == 7) //Липецк
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Лип* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Лип**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=*отделения*", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>служба");
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Лип**", XlAutoFilterOperator.xlAnd);
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Лип**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();
                    }
                    if (i == 8) //Москва
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Моск* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Моск**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Моск**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();
                    }
                    if (i == 9) //Смоленск
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Смол* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Смол**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Смол**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();
                    }
                    if (i == 10) //Тверь
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Тв* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Твер**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Твер**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        //xlsApp.Quit();
                    }
                    if (i == 11) //Тула
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Тул* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Тул**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Тул**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();

                    }
                    if (i == 12) //Ярослав
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Ярос* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ярос**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ярос**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();
                    }
                    if (i == 13) //Центральна
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Центральной* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Центральная**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Центральная**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        // xlsApp.Quit();
                    }
                    if (i == 14) //Москва
                    {
                        xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Приокс* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Приокс**", XlAutoFilterOperator.xlAnd);
                        var ZNYb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);


                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTechb = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Приокс**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ii++;
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        arr[i, 0] = ZNYb;
                        arr[i, 1] = OperatTechb;
                        arr[i, 2] = IniyDoljnostib;
                        //Directory.CreateDirectory(@"D:\");

                        //xlsApp.Quit();

                    }
                    i++;

                }
                while (i <= 14);
                listBox1.Items.Add("Comlete");
                //Directory.CreateDirectory(@"D:\");
                //string pathSave = String.Format(@"D:\test_{0}_{1}_{2}_{3}_{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                /*xlsBook.SaveAs(pathSave, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlsApp.Quit();
                */
                //xlsSheet.Cells[411, 5].Fo
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void YkomlektDopNew(string FileNameNew)
        {
            try
            {

                Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlsBook = xlsApp.Workbooks.Open(Convert.ToString(FileNameNew));//, ReadOnly: false); //рабочая книга
                Worksheet xlsSheet;
                xlsSheet = (Worksheet)xlsBook.Sheets[1]; // раздел 1 (лист 1)
                xlsSheet.Activate();

                var MyRange = "G2:G9000";

                //разделяем  столбец Кол-во/Дата Образования
                xlsSheet.Range["G2:G9000"].TextToColumns(xlsSheet.get_Range(MyRange, Type.Missing), XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierDoubleQuote, true, Type.Missing, Type.Missing, true, false, Type.Missing, ",", Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Directory.CreateDirectory(@"D:\testNew");
                int i = 0;
                int ii = 0;

                var Podrazdelenie = "Аппарат*";
                var Doljnost = "";
                var Otdel = "=Руководство";
                var BlokPOdraz = "";
                var VidSlyg = "=Сотрудник";
                var NameNew = "Начальник Це**";
                                                                                                            
                do
                {
                    string pathSave = String.Format(@"D:\testNew\{0}_test_{1}_{2}_{3}_{4}_{5}_{6}_{7}_{8}", i, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond, ii);
                    if (i == 0) //Аппарат
                    {
                        xlsSheet.Range["L1"].Formula = "=SUBTOTAL(109,R[1]C[-5]:R[8999]C[-5])";

                        xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                        // xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, Doljnost, XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, Otdel, XlAutoFilterOperator.xlAnd);
                        //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(4, BlokPOdraz, XlAutoFilterOperator.xlAnd);
                      //  xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, VidSlyg, XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, NameNew, XlAutoFilterOperator.xlAnd);

                        arrNew[i,0] = xlsSheet.Range["L1"].Value;
                        
                        /*xlsSheet.Range["$A$1:$G$8521"].AutoFilter(7, "=**Заместитель начальника Це* - начальник**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=*Центр*", XlAutoFilterOperator.xlAnd);
                        var ZNY = xlsSheet.Range["E1411"].Value;
                        arr[i, 0] = ZNY;*/
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        /*
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        ii++;

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**служба**");
                        var OperatTech = xlsSheet.Range["E1411"].Value;
                        arr[i, 1] = OperatTech;
                        //xlsBook.SaveAs(pathSave, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                        ii++;

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Центр**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                        var IniyDoljnosti = xlsSheet.Range["E1411"].Value;
                        arr[i, 2] = IniyDoljnosti;
                        //xlsBook.SaveAs(pathSave, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                        xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);*/
                        
                        //xlsApp.Quit();
                    }
                    i++;
                    #region
                    /*  if (i == 1) //Белгород
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Бел* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Белг**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          arr[i, 0] = ZNYb;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          arr[i, 1] = OperatTechb;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Белг**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          arr[i, 2] = IniyDoljnostib;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          //xlsApp.Quit();
                      }
                      if (i == 2) //Брянск
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Бря* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Бря**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;

                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=*отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;

                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Бря**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;

                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          ii++;
                          // xlsApp.Quit();
                      }
                      if (i == 3) // Владимир
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Влад** - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Влад**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Влад**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          //xlsApp.Quit();
                      }
                      if (i == 4) //Воронеж
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Ворон* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ворон**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ворон**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();
                      }
                      if (i == 5) //Калуга
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Калу* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Калу**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Калу**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=*корруп*", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();
                      }
                      if (i == 6) //Курск
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Кур* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Кур**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Кур**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();

                      }
                      if (i == 7) //Липецк
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Лип* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Лип**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=*отделения*", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>служба");
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Лип**", XlAutoFilterOperator.xlAnd);
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Лип**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();
                      }
                      if (i == 8) //Москва
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Моск* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Моск**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Моск**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();
                      }
                      if (i == 9) //Смоленск
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Смол* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Смол**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Смол**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();
                      }
                      if (i == 10) //Тверь
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Тв* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Твер**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Твер**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          //xlsApp.Quit();
                      }
                      if (i == 11) //Тула
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Тул* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Тул**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Тул**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();

                      }
                      if (i == 12) //Ярослав
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Ярос* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ярос**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Ярос**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();
                      }
                      if (i == 13) //Центральна
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Центральной* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Центральная**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          // xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Центральная**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          // xlsApp.Quit();
                      }
                      if (i == 14) //Москва
                      {
                          xlsSheet.Range["E1411"].Formula = "=SUBTOTAL(109,R[-1409]C:R[-11]C)";

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**Заместитель начальника Приокс* - начальник**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Приокс**", XlAutoFilterOperator.xlAnd);
                          var ZNYb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          //xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);


                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7, "=**отделения**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**служба**");
                          var OperatTechb = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(7);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);

                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1, "=**Приокс**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2, "<>**ачальн**", XlAutoFilterOperator.xlAnd);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlAnd, "<>**отделение От**");
                          var IniyDoljnostib = xlsSheet.Range["E1411"].Value;
                          xlsBook.SaveAs(pathSave + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                          ii++;
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(1);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(2);
                          xlsSheet.Range["$A$1:$G$1372"].AutoFilter(3);
                          arr[i, 0] = ZNYb;
                          arr[i, 1] = OperatTechb;
                          arr[i, 2] = IniyDoljnostib;
                          //Directory.CreateDirectory(@"D:\");

                          //xlsApp.Quit();

                      }*/
                    #endregion

                }
                while (i <= 1);
                listBox1.Items.Add("Comlete");
                //Directory.CreateDirectory(@"D:\");
                //string pathSave = String.Format(@"D:\test_{0}_{1}_{2}_{3}_{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                /*xlsBook.SaveAs(pathSave, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlsApp.Quit();
                */
                //xlsSheet.Cells[411, 5].Fo
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /* label1.Visible = true;
             listBox1.Visible = true;*/
            loadingNew1.Visible = true;
            try
            {

                OpenFileDialog dlg = new OpenFileDialog
                {
                    Multiselect = true,
                    Title = "Выберите файлы",
                    InitialDirectory = @"D:\",
                    Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                };
                dlg.ShowDialog();
                if (dlg.FileName == String.Empty)
                    return;
                
               foreach (string file in dlg.FileNames)
                {
                    var sym = "Укомплект";
                    var indexOfChar = file.IndexOf(sym); // равно 4*/
                    if (indexOfChar < 1)
                    {
                        int Index = 0;
                        int ii = 0;
                        string FileNameNew = file;
                        listBox1.Items.Add(FileNameNew);
                        foreach (string podr in Podrazd)
                        {

                            string NamePodraz = podr;
                          // if ((podr == "=*Апп*") == true)
                         //   {
                                ZapolnenieVakansii(FileNameNew, Index, ii);
                           // }
                                listBox1.Items.Add(NamePodraz + "Заполнена");
                            listBox1.TopIndex = listBox1.Items.Count - 1; // log следует за добавлениями файлов
                            Index++;
                            ii++;
                            //progress.Value++;
                            
                        }
                    }


                    if (indexOfChar > 1)
                    {
                      
                        string FileNameShtat = file;
                        listBox1.Items.Add(FileNameShtat);
                        IzvlechenieShtat(FileNameShtat);
                    }
                }
                ZapolnenieYkomplektovannosti();

            }
          

            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
            {
                currentProcess.Kill();
                
            }
        }

        public void ZapolnenieVakansii(string FileNameNew, int i, int ii)
        {
            //какой-то костыль с датой
            string DataAkt =String.Format("<={1}/{0}/{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year);
            
            try
            {
                Directory.CreateDirectory(@"D:\NewTestSprav");

                string pathSaveNew = String.Format(@"D:\NewTestSprav\{0}_test_{1}_{2}_{3}_{4}_{5}_{6}_{7}_{8}", i, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond, ii);
                string Podrazdelenie = Podrazd[i];
                Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlsBook = xlsApp.Workbooks.Open(Convert.ToString(FileNameNew));
                Worksheet xlsSheet;
                xlsSheet = (Worksheet)xlsBook.Sheets[1]; // раздел 1 (лист 1)
                xlsSheet.Activate();
                var MyRange = "G2:G9000";
                //this.Dis
                //разделяем  столбец Кол-во/Дата Образования
                xlsApp.DisplayAlerts = false;
                xlsSheet.Range["G2:G9000"].TextToColumns(xlsSheet.get_Range(MyRange, Type.Missing), XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierDoubleQuote, true, Type.Missing, Type.Missing, true, false, Type.Missing, ",", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlsSheet.Columns[8].ColumnWidth = (17.43);
                xlsSheet.Cells[1, 8].value = "Дата";
                xlsSheet.Columns[8].Replace(" ", "", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Range["$H$1:$H$8521"].NumberFormat = "m/d/yyyy";
                xlsSheet.Columns[8].Replace("", "01.01.2001", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells[1, 8].AutoFilter();

                
                //xlsSheet.Range["G2:G9000"].TextToColumns(xlsSheet.get_Range(MyRange, Type.Missing), XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierDoubleQuote, true, Type.Missing, Type.Missing, true, false, false, ",", FieldInfo: new object[,] { { 1, XlColumnDataType.xlGeneralFormat}, { 2, XlColumnDataType.xlDMYFormat} }, TrailingMinusNumbers: true);

                xlsSheet.Range["L1"].Formula = "=SUBTOTAL(109,R[1]C[-5]:R[8999]C[-5])";
                Directory.CreateDirectory(@"D:\testNew");
                
                xlsSheet.Cells.Replace("-1", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-2", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-.5", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace(".5", "0,5", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                int iii = 1;

                xlsSheet.Cells[1, 8].AutoFilter(8,Criteria1: "01/01/2001", Operator: XlAutoFilterOperator.xlOr);// тот же костыль
                xlsSheet.Cells[1, 9].AutoFilter(9, Criteria1: "**Вакантна**", Operator: XlAutoFilterOperator.xlOr);
                //1
                xlsSheet.Range["$A$1:$H$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$H$8521"].AutoFilter(2, "=Начальник*", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$H$8521"].AutoFilter(3, "=**Руководство**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$H$8521"].AutoFilter(6, "**Начальник **", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["H2:H8521"].AutoFilter(8, DataAkt); // без этого костыля (с датой) почему то не работало
                
               // xlsBook.SaveAs(pathSaveNew + ii+iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
                arrNew[i, 0] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                //1
               // Thread.Sleep(1000);
               // iii++;
                //2
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**Руководство**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "*Первый заместитель*нач**", XlAutoFilterOperator.xlOr);
               // xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 1] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                //2
               // Thread.Sleep(1000);
                //iii++;
                //3
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=**Зам*нач**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=Руководство*", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "<>**службы**", XlAutoFilterOperator.xlAnd, "<>*первый**");
               //xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 2] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //3
               // Thread.Sleep(1000);
              //  iii++;
                //4
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=**Заместитель начальника**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**Руководство**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**зам*нач*уп*-*нач**", XlAutoFilterOperator.xlAnd, "<>*таможни*");
              //  xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 3] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //4
               // Thread.Sleep(1000);
              //  iii++;
                //5
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Начальник*", XlAutoFilterOperator.xlOr);
                // xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**Руководство**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=*нач*службы**", XlAutoFilterOperator.xlAnd, "<>**отдел**");
              //  xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 4] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //5
               // Thread.Sleep(1000);
               // iii++;
                //6
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Зам*нач*", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**служба**", XlAutoFilterOperator.xlAnd, "<>**отдел**");
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=*зам*нач*службы**", XlAutoFilterOperator.xlAnd, "<>**-*нач*служ**");
               // xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 5] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //6

               // Thread.Sleep(1000);
              //  iii++;
                //7
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Начальник", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "<>**отдел**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**нач*поста**", XlAutoFilterOperator.xlOr, "=**ТП*"); // или *ТП*
              //  xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 6] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //7

               // Thread.Sleep(1000);
              //  iii++;
                //8
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Зам*Начальник*", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "<>**отдел**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=*Зам*нач*пост**", XlAutoFilterOperator.xlOr, "=**ТП*");  // или *ТП*
                //xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 7] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //8
               // Thread.Sleep(1000);
               // iii++;
                //9
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Начальник", XlAutoFilterOperator.xlOr, "=Начальник от*");
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "<>**корруп**", XlAutoFilterOperator.xlAnd, "<>*опер*тех**");
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**нач*отдел**", XlAutoFilterOperator.xlAnd);
               // xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 8] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //9
               // Thread.Sleep(1000);
              //  iii++;
                //10
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Зам*Начальник*", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "<>**корруп**", XlAutoFilterOperator.xlAnd, "<>*пост*");
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=*зам*нач*отдел**", XlAutoFilterOperator.xlAnd, "<>поста"); // не содержит корруп
            // xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 9] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //10
              //  Thread.Sleep(1000);
             //   iii++;
                //11
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "<>**Начальник**", XlAutoFilterOperator.xlOr, "=**смен**");
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "<>**корруп**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "<>**корруп**", XlAutoFilterOperator.xlAnd, "<>**рук**");
               // xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 10] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //11
               // Thread.Sleep(1000);
              //  iii++;
                //12
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "<>**Начальник**", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "<>**корруп**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**опер*тех**", XlAutoFilterOperator.xlAnd);
             // xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 11] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //12
                Thread.Sleep(1000);
                iii++;
                //13
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Начальник", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**опер*тех**", XlAutoFilterOperator.xlAnd);
              //  xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 12] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //13
             //   Thread.Sleep(1000);
             //   iii++;
                //14
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "<>Начальник*", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Сотрудник**", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**опер*тех**", XlAutoFilterOperator.xlAnd);
            //   xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 13] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //14
             //  Thread.Sleep(1000);
            //    iii++;
                //15
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Начальник", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Гос**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**поста**", XlAutoFilterOperator.xlAnd, "<>**отдел**");
              //  xlsBook.SaveAs(pathSaveNew + ii + iii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 14] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //15
            //    Thread.Sleep(1000);
            //    iii++;
                //16
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Зам*начальник*", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Гос**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**поста**", XlAutoFilterOperator.xlAnd, "<>**отдел**");
              // xlsBook.SaveAs(pathSaveNew + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 15] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //16

                //17
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Начальник", XlAutoFilterOperator.xlOr, "=**отд**");
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Гос**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**нач*отдел**", XlAutoFilterOperator.xlAnd, "<>Заместитель*");
               // xlsBook.SaveAs(pathSaveNew + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 16] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //17

                //18
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Зам*начальник*", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Гос**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**нач*отдел**", XlAutoFilterOperator.xlAnd);
                //xlsBook.SaveAs(pathSaveNew + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 17] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //18

                //19
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "<>**начальник*", XlAutoFilterOperator.xlOr, "=**помощ**");
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Гос**", XlAutoFilterOperator.xlOr);
                // xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**нач*отдел**", XlAutoFilterOperator.xlAnd);
                //xlsBook.SaveAs(pathSaveNew + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 18] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);
                //19

                //20
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, Podrazdelenie, XlAutoFilterOperator.xlAnd);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, "=Зам*начальник*", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3, "=**корруп**", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "=**Работник**", XlAutoFilterOperator.xlOr);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5, "Работник", XlAutoFilterOperator.xlOr);
                //xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6, "=**нач*отдел**", XlAutoFilterOperator.xlAnd);
                xlsBook.SaveAs(pathSaveNew + ii, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                arrNew[i, 19] = xlsSheet.Range["L1"].Value;
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(3);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(6);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2);
                xlsSheet.Range["$A$1:$G$8521"].AutoFilter(5);

                //20

                xlsBook.Close(false);
                iii = 1;

            }

            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error в " + s.Source,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void IzvlechenieShtat(string FileNameShtat)
        {
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlsBook = xlsApp.Workbooks.Open(Convert.ToString(FileNameShtat));//, ReadOnly: false); //рабочая книга
            Worksheet xlsSheet;
            xlsSheet = (Worksheet)xlsBook.Sheets[1]; // раздел 1 (лист 1)
            xlsSheet.Activate();
            //Итого
            string yacheikaBykva = "F";
            int yacheikaShtatItog = 3;
            int yacheikaFaktItog = 4;
            int yacheikaVakansItog = 5;
            //Итого
            
            //Работники
            string BykvaRab = "E";
            int yacheikaShtatRab = 3;
            int yacheikaFaktIRab = 4;
            int yacheikaVakansRab = 5;
            //Работники


            //Гос
            string BykvaGos = "D";
            int yacheikaShtatGos = 3;
            int yacheikaFaktGos = 4;
            int yacheikaVakansGos = 5;
            //Гос

            //Сотрудники
            string BykvaSotr = "C";
            int yacheikaShtatSotr = 3;
            int yacheikaFaktSotr = 4;
            int yacheikaVakansSotr = 5;
            //Сотрудники

            //Итого
            for (int i = 0; i <= 14; i++)
            {
                string ShtatItog = (yacheikaBykva + yacheikaShtatItog);
                string FaktItog = (yacheikaBykva + yacheikaFaktItog);
                string VakansItog = (yacheikaBykva + yacheikaVakansItog);
                arrShtat[i] = xlsSheet.Range[ShtatItog].Value;
                arrFakt[i] = xlsSheet.Range[FaktItog].Value;
                arrVakan[i] = xlsSheet.Range[VakansItog].Value;

                yacheikaShtatItog += 3;
                yacheikaFaktItog += 3;
                yacheikaVakansItog += 3;
            }

            //Работники
            for (int i = 0; i <= 14; i++)
            {
                string RabShtat = (BykvaRab + yacheikaShtatRab);
                string RabFakt = (BykvaRab + yacheikaFaktIRab);
                string RabVakans = (BykvaRab + yacheikaVakansRab);
                arrShtatRab[i] = xlsSheet.Range[RabShtat].Value;
                arrFaktRab[i] = xlsSheet.Range[RabFakt].Value;
                arrVakansRab[i] = xlsSheet.Range[RabVakans].Value;

                yacheikaShtatRab += 3;
                yacheikaFaktIRab += 3;
                yacheikaVakansRab += 3;
            }

            //Гос
            for (int i = 0; i <= 14; i++)
            {
                string GosShtat = (BykvaGos + yacheikaShtatGos);
                string GosFakt = (BykvaGos + yacheikaFaktGos);
                string GosVakans = (BykvaGos + yacheikaVakansGos);
                arrShtatGos[i] = xlsSheet.Range[GosShtat].Value;
                arrFaktGos[i] = xlsSheet.Range[GosFakt].Value;
                arrVakansGos[i] = xlsSheet.Range[GosVakans].Value;

                yacheikaShtatGos += 3;
                yacheikaFaktGos += 3;
                yacheikaVakansGos += 3;
            }

            //Сотрудники
            for (int i = 0; i <= 14; i++)
            {
                string SotrShtat = (BykvaSotr + yacheikaShtatSotr);
                string SotrFakt = (BykvaSotr + yacheikaFaktSotr);
                string SotrVakans = (BykvaSotr + yacheikaVakansSotr);
                arrShtatSotr[i] = xlsSheet.Range[SotrShtat].Value;
                arrFaktSotr[i] = xlsSheet.Range[SotrFakt].Value;
                arrVakansSotr[i] = xlsSheet.Range[SotrVakans].Value;

                yacheikaShtatSotr += 3;
                yacheikaFaktSotr += 3;
                yacheikaVakansSotr += 3;
            }

        }


        public void ZapolnenieYkomplektovannosti ()
        {
           
            var structPodr = 5;
            int i = 0;
           

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            //Книга. 
            
            ObjWorkBook = ObjExcel.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory + @"Templates\Укомплектованность.xls");
            
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            do
            {
                ObjWorkSheet.Cells[structPodr, 3] = arrNew[i, 0]; // НАЧАЛЬНИК УПРАВЛЕНИЯ
                ObjWorkSheet.Cells[structPodr, 4].value = arrNew[i, 1]; // Первый заместитель
                ObjWorkSheet.Cells[structPodr, 5].value = arrNew[i, 2]; // Заместитель НУ
                ObjWorkSheet.Cells[structPodr, 6].value = arrNew[i, 3]; // ЗНУ
                ObjWorkSheet.Cells[structPodr, 7].value = arrNew[i, 4]; //Начальник службы
                ObjWorkSheet.Cells[structPodr, 8].value = arrNew[i, 5]; // Заместитель начальника службы
                ObjWorkSheet.Cells[structPodr, 9].value = arrNew[i, 6]; // Начальник поста
                ObjWorkSheet.Cells[structPodr, 10].value = arrNew[i, 7]; //Заместитель начальника поста
                ObjWorkSheet.Cells[structPodr, 11].value = arrNew[i, 8]; //НО (отделения)
                ObjWorkSheet.Cells[structPodr, 12].value = arrNew[i, 9]; // ЗНО
                ObjWorkSheet.Cells[structPodr, 13].value = arrNew[i, 10]; // Иные должности
                ObjWorkSheet.Cells[structPodr, 14].value = arrNew[i, 11]; //Оперативно-технические (все должности)
                ObjWorkSheet.Cells[structPodr, 15].value = arrNew[i, 12]; // НО противКорруп
                ObjWorkSheet.Cells[structPodr, 16].value = arrNew[i, 13];// Иные должности ПротивКорруп
                ObjWorkSheet.Cells[structPodr, 17].value = arrNew[i, 14];// Начальник поста ФГГС
                ObjWorkSheet.Cells[structPodr, 18].value = arrNew[i, 15];// Заместитель начальника поста ФГГС
                ObjWorkSheet.Cells[structPodr, 19].value = arrNew[i, 16]; // НО ФГГС
                ObjWorkSheet.Cells[structPodr, 20].value = arrNew[i, 17]; // ЗНО ФГГС
                ObjWorkSheet.Cells[structPodr, 21].value = arrNew[i, 18]; // Иные должности ФГГС
                ObjWorkSheet.Cells[structPodr, 22].value = arrNew[i, 19]; // Работники ФГГС
                ObjWorkSheet.Cells[structPodr, 24].value = arrShtat[i]; //Штат




                //Доп
                ObjWorkSheet.Cells[structPodr, 32].value = arrShtatSotr[i]; //Штат
                ObjWorkSheet.Cells[structPodr, 33].value = arrFaktSotr[i]; //Штат
                ObjWorkSheet.Cells[structPodr, 34].value = arrVakansSotr[i]; //Штат

                ObjWorkSheet.Cells[structPodr, 35].value = arrShtatGos[i]; //Штат
                ObjWorkSheet.Cells[structPodr, 36].value = arrFaktGos[i]; //Штат
                ObjWorkSheet.Cells[structPodr, 37].value = arrVakansGos[i]; //Штат

                ObjWorkSheet.Cells[structPodr, 38].value = arrShtatRab[i]; //Штат
                ObjWorkSheet.Cells[structPodr, 39].value = arrFaktRab[i]; //Штат
                ObjWorkSheet.Cells[structPodr, 40].value = arrVakansRab[i]; //Штат

                //Доп
                //listBox1.Items.Add(structPodr + " " + i);
                structPodr++;
                i++;
                //listBox1.Items.Add(structPodr +" "+ i);
            }

            while (structPodr <= 19);

            string DatePoSostoyaniu = (Date1.ToString("dd MMMM yyyy"));
            ObjWorkSheet.Cells[1, 2] = ("Справка об укомплектованности Центрального таможенного управления по состоянию на: " + DatePoSostoyaniu + "г.");

            ObjExcel.DisplayAlerts = false;
            Directory.CreateDirectory(@"D:\Справки");
            string pathSave = String.Format(@"D:\Справки\Справка_Об_Укомплектованности_{0}_{1}_{2}_{3}_{4}_New", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);

            ObjWorkBook.SaveAs(pathSave, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            DialogResult dialogResult = MessageBox.Show("Открыть созданный файл?", "Открытие файла", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ObjExcel.Visible = true;
                ObjExcel.UserControl = true;
            }
            loadingNew1.Visible = false;
            MessageBox.Show("Сохранён по адресу: \nD:\\Справки\\", "Файл успешно сформирован");
            ObjExcel.Quit();
         
            //xlsBookOpen.Close(0);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //progress.Maximum = 1350;
            try
            {
                Doljnosti doljnosti = new Doljnosti();
                OpenFileDialog dlg = new OpenFileDialog
                {
                    Multiselect = true,
                    Title = "Выберите файлы",
                    InitialDirectory = @"D:\",
                    Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                };
                dlg.ShowDialog();

                if (dlg.FileName == String.Empty)
                    return;

                var FileName = dlg.FileName;
                doljnosti.FailZapolnenieDoljnostiAsync(FileName);
                backgroundWorker1.RunWorkerAsync();
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
            {
                currentProcess.Kill();
            }

        }



        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            //progress.Value = e.ProgressPercentage;
            // Set the text.
            this.Text = e.ProgressPercentage.ToString();
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            for (int i = 1; i <= 1350; i++)
            {
                // Wait 100 milliseconds.
                Thread.Sleep(8);
                // Report progress.
                backgroundWorker1.ReportProgress(i);
            }
        }

        private void YkomplektovannostCTY_Click(object sender, EventArgs e)
        {
            try
            {
                Loading ld = new Loading();
                YkomplektovannostCTY ykompCTY = new YkomplektovannostCTY();
                

                OpenFileDialog dlg = new OpenFileDialog
                {
                    Multiselect = true,
                    Title = "Выберите файлы",
                    InitialDirectory = @"D:\",
                    Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                };
                dlg.ShowDialog();

                if (dlg.FileName == String.Empty)
                    return;
                    
                ykompCTY.Copy16KadYkomplektovannostAsync(dlg.FileNames);

                //loadingNew1.Visible = false;

            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
            foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
            {
                currentProcess.Kill();
               
            }
            
        }

        private void StatFact_Click(object sender, EventArgs e)
        {


            try
            {
                Shtat_Fakt shtat_Fakt = new Shtat_Fakt();
                OpenFileDialog dlg = new OpenFileDialog
                {
                    Multiselect = true,
                    Title = "Выберите файлы",
                    InitialDirectory = @"D:\",
                    Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                };
                dlg.ShowDialog();

                if (dlg.FileName == String.Empty)
                    return;


                shtat_Fakt.FormirovanieStatFactAsync(dlg.FileName);
                
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
            {
                currentProcess.Kill();

            }



        }

        private void SvodKad1_Click(object sender, EventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                try
                {
                    Loading ld = new Loading();
                    SvodKad1 svod1kad = new SvodKad1();


                    OpenFileDialog dlg = new OpenFileDialog
                    {
                        Multiselect = true,
                        Title = "Выберите файлы",
                        InitialDirectory = @"D:\",
                        Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                    };
                    dlg.ShowDialog();

                    if (dlg.FileName == String.Empty)
                        return;


                    loadingNew2.Visible = true;
                    svod1kad.Copy1KadAsync(dlg.FileNames);
                    //loadingNew1.Visible = false;

                    //loadingNew1.Visible = false;

                }
                catch (Exception s)
                {
                    MessageBox.Show(s.Message,
                        "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally

                {
                    loadingNew2.Visible = false;
                }

                foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
                {
                    currentProcess.Kill();
                    //loadingNew1.Visible = false;
                }
            }));
            //loadingNew1.Visible = false;
        }

        private void Kad16Vigruzka_Click(object sender, EventArgs e)
        {
           
            _16KADVigruzka Vigruzka16kad = new _16KADVigruzka();
            BeginInvoke(new MethodInvoker(delegate
            {
                try
                {
                    loadingNew2.Visible = true;
                    Kad16Vigruzka.Enabled = false;
                    Vigruzka16kad.F16KadAsync(dat);
                }

                catch (Exception s)
                {
                    MessageBox.Show(s.Message,
                        "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                

               
            }));

        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                SelectDatee.Text = e.End.ToString();
            dat = SelectDatee.Text;
            monthCalendar1.Visible = false;
        }));
        }

        private void SelectDatee_MouseClick(object sender, MouseEventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate
            {
                monthCalendar1.Visible = true;
        }));
        }


        public void CloseLoad()
        {
            loadingNew2.Visible = false;
            loadingNew2.Hide();
            Controls.Remove(loadingNew2);
            this.Refresh();
        }

        private void KommandirovanieMOT_Click(object sender, EventArgs e)
        {
            MOT mt = new MOT();
            mt.ShowDialog();
        }

        private void ShtatFaktNew_Click(object sender, EventArgs e)
        {
            ShtatFaktView shtatFaktView = new ShtatFaktView();
            shtatFaktView.ShowDialog();
        }

        private void Form_6_2_CAD_Click(object sender, EventArgs e)
        {
            CAD_2_6 cad_2_6 = new CAD_2_6();
            cad_2_6.ShowDialog();
        }
    }

}
