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
    class SvodKad1
    {
        int count = 0;
        Excel.Application ObjWorkExcel = new Excel.Application(); //сам эксель

        Excel.Workbook ObjWorkBooks;// из этой книги будем копировать
        Excel.Worksheet ObjWorkSheets;//с этого листа
        //Excel.Workbook WB; //в эту книгу будем копировать
        Excel.Worksheet WS;// в этот лист
        Form1 fr = new Form1();
        //LoadingNew 

        public async void Copy1KadAsync(string[] FileNam)
        {

      

            await Task.Run(() => CopyKad1(FileNam));

           
        }


        public void CopyKad1(string[] FileNam)
        {
            
            try
            {

                //string pathSaveYkomplektCTY = String.Format(@"D:\Справки\Укомплектованность_ЦТУ_{0}-{1}-{2}_{3}%{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //сам эксель
                string pathSaveSvod1Kad = String.Format(@"D:\Справки\1-KAD_{0}-{1}-{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year);

                Workbook WB = ObjWorkExcel.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory + @"Templates\1-KAD.xlsx");//создаем новую книгу
                int RowCount = 0;



                foreach (string FileName in FileNam)
                {
                    ObjWorkBooks = ObjWorkExcel.Workbooks.Open(Convert.ToString(FileName)); //открываем существующую книгу
                    Worksheet xlsSheet;
                    ObjWorkExcel.DisplayAlerts = false;
                    xlsSheet = (Worksheet)ObjWorkBooks.Sheets[1]; // раздел 1 (лист 1)
                    xlsSheet.Activate();
                    //Считывание с определенных ячеек
                    /*var KodOrgana = Convert.ToString(xlsSheet.Range["K20"].Value);

                    KodOrgana= KodOrgana.Substring(24);*/


                    fr.Refresh();

                    if (FileName.Contains("АУ")) //Аппарат
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C36" + ":AJ36";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C13" + ":Z13";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C13" + ":Q13";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C13" + ":BS13";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C13" + ":AQ13";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C13" + ":N13";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                                                                                          

                    if (FileName.Contains("БеТ")) //Белгород
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                       //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C37" + ":AJ37";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                         Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C14" + ":Z14";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C14" + ":Q14";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C14" + ":BS14";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C14" + ":AQ14";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C14" + ":N14";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }


                    if (FileName.Contains("БрТ")) //Брянск
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C38" + ":AJ38";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C15" + ":Z15";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C15" + ":Q15";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C15" + ":BS15";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C15" + ":AQ15";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C15" + ":N15";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }


                    if (FileName.Contains("ВлТ")) //Владимир
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C39" + ":AJ39";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C16" + ":Z16";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C16" + ":Q16";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C16" + ":BS16";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C16" + ":AQ16";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C16" + ":N16";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                    if (FileName.Contains("ВоТ")) //Воронеж
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C40" + ":AJ40";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C17" + ":Z17";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C17" + ":Q17";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C17" + ":BS17";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C17" + ":AQ17";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C17" + ":N17";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }


                    if (FileName.Contains("КаТ")) //Калуга
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C41" + ":AJ41";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C18" + ":Z18";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C18" + ":Q18";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C18" + ":BS18";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C18" + ":AQ18";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C18" + ":N18";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                    if (FileName.Contains("КуТ")) //Курск
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C42" + ":AJ42";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C19" + ":Z19";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C19" + ":Q19";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C19" + ":BS19";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C19" + ":AQ19";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C19" + ":N19";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }


                    if (FileName.Contains("ЛиТ")) //Липецк
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C43" + ":AJ43";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C20" + ":Z20";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C20" + ":Q20";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C20" + ":BS20";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C20" + ":AQ20";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C20" + ":N20";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }


                    if (FileName.Contains("МТ")) //Москва
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C44" + ":AJ44";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C21" + ":Z21";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C21" + ":Q21";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C21" + ":BS21";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C21" + ":AQ21";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C21" + ":N21";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                    if (FileName.Contains("СмТ")) //Смоленск
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C45" + ":AJ45";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C22" + ":Z22";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C22" + ":Q22";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C22" + ":BS22";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C22" + ":AQ22";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C22" + ":N22";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                                       
                    if (FileName.Contains("ТвТ")) //Тверь
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C46" + ":AJ46";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C23" + ":Z23";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C23" + ":Q23";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C23" + ":BS23";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C23" + ":AQ23";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C23" + ":N23";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                    if (FileName.Contains("ТуТ")) //Тула
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C47" + ":AJ47";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C24" + ":Z24";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C24" + ":Q24";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C24" + ":BS24";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C24" + ":AQ24";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C24" + ":N24";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }


                    if (FileName.Contains("ЯрТ")) //Ярославль
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C48" + ":AJ48";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C25" + ":Z25";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C25" + ":Q25";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C25" + ":BS25";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C25" + ":AQ25";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C25" + ":N25";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                    if (FileName.Contains("ЦОТ")) //ЦОТ
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C47" + ":AJ47";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C26" + ":Z26";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C26" + ":Q26";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C26" + ":BS26";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C26" + ":AQ26";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C26" + ":N26";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                    }

                    if (FileName.Contains("ПТТП")) //Приокский
                    {
                        //ЛИСТ 1
                        ObjWorkSheets = ObjWorkBooks.Worksheets[1];//берем 1 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        string Name = "C33:AJ33"; //берем диапазон ячеек от C37 до AJ37
                        string Adress = "C50" + ":AJ50";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ2
                        ObjWorkSheets = ObjWorkBooks.Worksheets[2];//берем 2 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Z12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C27" + ":Z27";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[2];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;


                        //ЛИСТ3
                        ObjWorkSheets = ObjWorkBooks.Worksheets[3];//берем 3 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:Q12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C27" + ":Q27";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[3];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ4
                        ObjWorkSheets = ObjWorkBooks.Worksheets[4];//берем 4 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C12:BS12"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C27" + ":BS27";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[4];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ5
                        ObjWorkSheets = ObjWorkBooks.Worksheets[5];//берем 5 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:AQ13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C27" + ":AQ27";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[5];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;

                        //ЛИСТ6
                        ObjWorkSheets = ObjWorkBooks.Worksheets[6];//берем 6 лист
                                                                   //int LastRow = ObjWorkBooks.Sheets[2].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку, можно заменить на известный номер
                        Name = "C13:N13"; //берем диапазон ячеек от C37 до AJ37
                        Adress = "C27" + ":N27";//формируем адрес куда в существующей книге мы это скопируем

                        WS = (Excel.Worksheet)WB.Sheets[6];//выбираем 1 лист
                        ObjWorkSheets.get_Range(Name).Copy(); // копи
                        WS.get_Range(Adress).PasteSpecial(); // паст;
                        fr.Refresh();
                    }

                    ObjWorkBooks.Close(false, Type.Missing, Type.Missing);//закрываем книгу из которой копировали*/
                    fr.Refresh();
                }

                WB.SaveAs(pathSaveSvod1Kad);
                //, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                WB.Close(false, Type.Missing, Type.Missing);
                
                fr.Refresh();
                fr.loadingNew2.Visible = false;
                MessageBox.Show("Сохранён по адресу: " + "\n" + pathSaveSvod1Kad, "Файл успешно сформирован");
                fr.loadingNew2.Visible = false;
            }

            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                fr.loadingNew2.Visible = false;
            }
          
        }
               

    }
}
