using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelR
{
    class YkomplektovannostCTY
    {
        int count = 0;
        Excel.Application ObjWorkExcel = new Excel.Application(); //сам эксель

        Excel.Workbook ObjWorkBooks;// из этой книги будем копировать
        Excel.Worksheet ObjWorkSheets;//с этого листа
        //Excel.Workbook WB; //в эту книгу будем копировать
        Excel.Worksheet WS;// в этот лист
        Form1 fr = new Form1();

        public async void Copy16KadYkomplektovannostAsync(string[] FileNam)
        {

           

            await Task.Run(() => Copy16KadYkomplektovannost(FileNam));



        }


        public  void Copy16KadYkomplektovannost(string[] FileNam)
        {
            
            try
            {
                
                //string pathSaveYkomplektCTY = String.Format(@"D:\Справки\Укомплектованность_ЦТУ_{0}-{1}-{2}_{3}%{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //сам эксель
                string pathSaveYkomplektCTY = String.Format(@"D:\Справки\Укомплектованность_ЦТУ_{0}-{1}-{2}_{3}%{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                
                Workbook WB = ObjWorkExcel.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory + @"Templates\Укомплектованность_ЦТУ.xls");//создаем новую книгу
                int RowCount = 0;
                WS = (Excel.Worksheet)WB.Sheets[1];
                WS.Range["A2"].Value = String.Format("{0}.{1}.{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year);



                foreach (string FileName in FileNam)
                {
                    ObjWorkBooks = ObjWorkExcel.Workbooks.Open(Convert.ToString(FileName)); //открываем существующую книгу
                    Worksheet xlsSheet;
                    ObjWorkExcel.DisplayAlerts = false;
                    xlsSheet = (Worksheet)ObjWorkBooks.Sheets[1]; // раздел 1 (лист 1)
                    xlsSheet.Activate();
                    //Считывание с определенных ячеек
                    int KodOrgana = Convert.ToInt32(xlsSheet.Range["S15"].Value);

                    ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                    int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                    string Name = "A1:EP" + LastRow.ToString(); //берем диапазон ячеек от а1 до еP+последняя строка
                    string Adress = "A1" + ":EP" + (RowCount + LastRow - 1).ToString();//формируем адрес куда в существующей книге мы это скопируем

                    if ((KodOrgana == 10100000) == true) //аппарат
                    {
                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10101000) == true) //белогород
                    {
                        WS = (Excel.Worksheet)WB.Sheets[4];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст);
                        count++;
                    }
                    if ((KodOrgana == 10119000) == true) //ЦОТ
                    {
                        WS = (Excel.Worksheet)WB.Sheets[16];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10102000) == true) //брянск
                    {
                        WS = (Excel.Worksheet)WB.Sheets[5];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10103000) == true) //владимир
                    {
                        WS = (Excel.Worksheet)WB.Sheets[6];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10104000) == true) //воронеж
                    {
                        WS = (Excel.Worksheet)WB.Sheets[7];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10106000) == true) //калуга
                    {
                        WS = (Excel.Worksheet)WB.Sheets[8];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10108000) == true) //курск
                    {
                        WS = (Excel.Worksheet)WB.Sheets[9];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10109000) == true) //липецк
                    {
                        WS = (Excel.Worksheet)WB.Sheets[10];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        count++;
                    }
                    if ((KodOrgana == 10113000) == true) //смоленск
                    {
                        WS = (Excel.Worksheet)WB.Sheets[12];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;;
                        count++;
                    }
                    if ((KodOrgana == 10115000) == true) //тверь
                    {
                        WS = (Excel.Worksheet)WB.Sheets[13];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;;
                        count++;
                    }
                    if ((KodOrgana == 10116000) == true) //тула
                    {
                        WS = (Excel.Worksheet)WB.Sheets[14];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;;
                        count++;
                    }
                    if ((KodOrgana == 10117000) == true) //ярославль
                    {
                        WS = (Excel.Worksheet)WB.Sheets[15];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;;
                        count++;
                    }
                    if ((KodOrgana == 10120010) == true) //приокск
                    {
                        WS = (Excel.Worksheet)WB.Sheets[17];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;;
                        count++;
                    }
                    if ((KodOrgana == 10129000) == true) //москва
                    {
                        WS = (Excel.Worksheet)WB.Sheets[11];
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;;
                        count++;
                    }
                    ObjWorkBooks.Close(false, Type.Missing, Type.Missing);//закрываем книгу из которой копировали
                }
                
                 WB.SaveAs(pathSaveYkomplektCTY, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                WB.Close(false, Type.Missing, Type.Missing);
                MessageBox.Show("Сохранён по адресу: " +"\n" + pathSaveYkomplektCTY, "Файл успешно сформирован");

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
