using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelR
{
    class Shtat_Fakt
    {
        private string[] Struct = { "Центральная оперативная таможня", "Аппарат Центрального таможенного управления", "Курская таможня" };
        private string[] VidRab = { "Госслужащий", "Сотрудник",  "Работник" };
        private string[] Komplect = { "Штат", "Факт" };

        private string[] Otdel = {
                                            "Руководство",
                                            "Отдельная должность",
                                            "Отдел защиты государственной тайны и специальной документальной связи",
                                            "Отдел документационного обеспечения",
                                            "Отдел бухгалтерского учета и финансового мониторинга",
                                            "Организационно - инспекторский отдел",
                                            "Отдел оперативных учетов",
                                            "Отдел оперативно-дежурной службы и таможенной охраны",
                                            "Отдел организации и контроля за деятельностью правоохранительных подразделений",
                                            "Отдел по борьбе с экономическими таможенными правонарушениями",
                                            "Отдел по борьбе с особо опасными видами контрабанды",
                                            "Отдел по борьбе с контрабандой наркотиков",
                                            "Оперативно-аналитический отдел",
                                            "Отделение сотрудничества с правоохранительными органами зарубежных стран",
                                            "Отдел  организации и контроля деятельности СОБР",
                                            "**Специальный отряд**",
                                            "Служба организации кинологической деятельности",
                                            "**Информационно-аналитический отдел**",
                                            "Отдел кинологической подготовки**",
                                            "Отдел организации административных расследований",
                                            "Отдел административных расследований",
                                            "Отдел организации дознания",
                                            "Отдел дознания",
                                            "Отдел исполнения поручений по уголовным делам и делам об административных правонарушениях",
                                            "Учетно-регистрационный отдел",
                                            "Отдел контроля соблюдения законности при привлечении к административной ответственности",
                                            "Отдел распоряжения имуществом и исполнения постановлений уполномоченных органов"
                                         };

        private string[] CatDoljnostei ={
                                            "Высший начальствующий состав",
                                            "Старший начальствующий состав",
                                            "Средний начальствующий состав",
                                            "Младший состав"
                                         };

        private string[] ZvanieChin = {
                                        "=*СГГС**",
                                        "=*РГГС**",
                                        "=*СрГГС**",
                                      };

        private string[] RabDolj = {
                                        "Руководители",
                                        "Специалисты и служащие",
                                        "Рабочие"
                                    };
                   
            


        Excel.Application ObjWorkExcel = new Excel.Application(); //сам эксель


        public async  void FormirovanieStatFactAsync(string FileNam)
        {
            await Task.Run(() => FormirovanieStatFact(FileNam));

        }
        public  void FormirovanieStatFact(string FileNam)
        {
            try
            {
               


                Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //открыть эксель
                Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory + @"Templates\ШтатФакт.xls");
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
                var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);//1 ячейку
                int numCol = 1;

                Range usedColumn = ObjWorkSheet.UsedRange.Range["A5:A31"];
                System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
                string[] arrOtdel = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

                Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlsBook = xlsApp.Workbooks.Open(Convert.ToString(FileNam));
                Worksheet xlsSheet;
                xlsSheet = (Worksheet)xlsBook.Sheets[1]; // раздел 1 (лист 1)
                xlsSheet.Activate();
                var MyRange = "G:G";

                xlsApp.DisplayAlerts = false;
                xlsSheet.Range["G:G"].TextToColumns(xlsSheet.get_Range(MyRange, Type.Missing), XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierDoubleQuote, true, Type.Missing, Type.Missing, true, false, Type.Missing, ",", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlsSheet.Cells.Replace(".5", "0,5", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-1", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-2", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-.5", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Range["G:G"].Replace("0,5", "0,5", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Range["M1:N1"].UnMerge();

                xlsSheet.Range["P1"].Formula = "=SUBTOTAL(109,R[1]C[-10]:R[9086]C[-10])";
                xlsSheet.Range["O1"].Formula = "=SUBTOTAL(109,RC[-8]:R[9000]C[-8])";
                Directory.CreateDirectory(@"D:\NewTestSpravDol\Otladka");
                Directory.CreateDirectory(@"D:\Справки");
                /*int shtatSotr = 0;
                int faktSotr = 0;
                int strokaShtat = 5;
                int stolbecShtat = 2;
                int strokaFact = 5;
                int stolbecFact = 3;*/
               // arrOtdel[15] = ;
                foreach (string struc in Struct)
                {
                    int shtatSotr = 0;
                    int faktSotr = 0;
                    int strokaShtat = 5;
                    int stolbecShtat = 2;
                    int strokaFact = 5;
                    int stolbecFact = 3;
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(1);
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(10);
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(11);
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(3);
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(14);
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(12);
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(13);
                    foreach (string vid in VidRab)
                    {

                        foreach (string otd in Otdel)
                        {

                            xlsSheet.Range["$A$1:$N$8521"].AutoFilter(1, struc, XlAutoFilterOperator.xlOr);
                            xlsSheet.Range["$A$1:$N$8521"].AutoFilter(10, vid, XlAutoFilterOperator.xlOr);
                            xlsSheet.Range["$A$1:$N$8521"].AutoFilter(11, otd, XlAutoFilterOperator.xlOr);
                            int shtat = Convert.ToInt32(xlsSheet.Range["P1"].Value);
                            int fakt = Convert.ToInt32(xlsSheet.Range["O1"].Value);


                            if ((otd == "Служба организации кинологической деятельности")&(vid == "Сотрудник"))
                            {
                                xlsSheet.Range["$A$1:$N$8521"].AutoFilter(11, "**Служба организации кинологической деятельности**", XlAutoFilterOperator.xlOr);
                                xlsSheet.Range["$A$1:$N$8521"].AutoFilter(3, "<>*отдел**", XlAutoFilterOperator.xlOr);
                                shtat = Convert.ToInt32(xlsSheet.Range["P1"].Value);
                                fakt = Convert.ToInt32(xlsSheet.Range["O1"].Value);
                                ObjWorkSheet.Cells[21, stolbecShtat] = shtat;
                                ObjWorkSheet.Cells[21, stolbecFact] = shtat - fakt;

                            }

                            xlsSheet.Range["$A$1:$N$8521"].AutoFilter(3);
                            ObjWorkSheet.Cells[strokaShtat, stolbecShtat] = shtat;
                            ObjWorkSheet.Cells[strokaFact, stolbecFact] = shtat- fakt;
                            strokaShtat++;
                            strokaFact++;
                        }
                        stolbecShtat += 2;
                        stolbecFact += 2;
                        strokaFact = 5;
                        strokaShtat = 5;
                    }

                    //xlsSheet.Range["$A$1:$K$8521"].AutoFilter(1);
                    xlsSheet.Range["$A$1:$H$8521"].AutoFilter(10);
                    xlsSheet.Range["$A$1:$H$8521"].AutoFilter(11);
                    xlsSheet.Range["$A$1:$K$8521"].AutoFilter(3);


                    
                    int StrokaSotr = 41;

                    foreach (string catdolj in CatDoljnostei)
                    {
                        xlsSheet.Range["P1"].Formula = "=SUBTOTAL(109,R[1]C[-10]:R[9086]C[-10])";
                        xlsSheet.Range["O1"].Formula = "=SUBTOTAL(109,RC[-8]:R[9000]C[-8])";
                        xlsSheet.Range["$A$1:$N$8521"].AutoFilter(10, "Сотрудник", XlAutoFilterOperator.xlOr);
                        xlsSheet.Range["$A$1:$N$8521"].AutoFilter(14, catdolj, XlAutoFilterOperator.xlOr);

                        shtatSotr = Convert.ToInt32(xlsSheet.Range["P1"].Value);
                        faktSotr = Convert.ToInt32(xlsSheet.Range["O1"].Value);
                       // ObjWorkSheet.Cells[StrokaSotr, 4] = shtatSotr;
                        ObjWorkSheet.Cells[StrokaSotr, 5] = shtatSotr - faktSotr;
                        
                        StrokaSotr++;
                    }


                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(14);
                    int shtatGos = 0;
                    int faktGos = 0;
                    int StrokaGos = 36;

                    foreach (string zvanchin in ZvanieChin)
                    {
                        xlsSheet.Range["$A$1:$N$8521"].AutoFilter(10, "Госслужащий", XlAutoFilterOperator.xlOr);
                        xlsSheet.Range["$A$1:$N$8521"].AutoFilter(12, zvanchin, XlAutoFilterOperator.xlOr);

                        shtatGos = Convert.ToInt32(xlsSheet.Range["P1"].Value);
                        faktGos = Convert.ToInt32(xlsSheet.Range["O1"].Value);
                        //ObjWorkSheet.Cells[StrokaGos, 2] = shtatGos;
                        ObjWorkSheet.Cells[StrokaGos, 3] = shtatGos - faktGos;
                        
                        StrokaGos++;

                    }
                    xlsSheet.Range["$A$1:$N$8521"].AutoFilter(12);
                    int shtatRab = 0;
                    int faktRab = 0;
                    int StrokaRab = 47;
                    foreach (string rabdoljnos in RabDolj)
                    {
                        xlsSheet.Range["$A$1:$N$8521"].AutoFilter(10, "Работник", XlAutoFilterOperator.xlOr);
                        xlsSheet.Range["$A$1:$N$8521"].AutoFilter(13, rabdoljnos, XlAutoFilterOperator.xlOr);

                        shtatRab = Convert.ToInt32(xlsSheet.Range["P1"].Value);
                        faktRab = Convert.ToInt32(xlsSheet.Range["O1"].Value);
                        //ObjWorkSheet.Cells[StrokaRab, 6] = shtatRab;
                        ObjWorkSheet.Cells[StrokaRab, 7] = shtatRab - faktRab;

                        StrokaRab++;

                    }




                    string pathSaveShtatFact = String.Format(@"D:\Справки\Штат-Факт_{0}_{1}-{2}-{3}_{4}%{5}", struc, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                    ObjWorkBook.SaveAs(pathSaveShtatFact, XlFileFormat.xlExcel8, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);




                }


                MessageBox.Show("Сохранён по адресу: " + "\n" +"D:\\Справки", "Файл успешно сформирован");

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
