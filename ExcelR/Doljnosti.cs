using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelR
{
    class Doljnosti
    {
       

        public double[,] TestMass = new double[89, 30];

        DateTime Date1 = DateTime.Now;
        Form1 form1 = new Form1();

        public async void  FailZapolnenieDoljnostiAsync(string FileName)
        {
            await Task.Run(() => ZapolnenieDoljnosti(FileName));
            
        }
              
        public void ZapolnenieDoljnosti(string FileName)
        {
           
          //  form1.progress.Maximum = (90 * 15);
            
            try
            {
                               
                //=======для заполнения======

                Microsoft.Office.Interop.Excel.Application ObjExcelZapol = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookZapol;
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetZapol;
                ObjWorkBookZapol = ObjExcelZapol.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory + @"Templates\ЦТУ_Работники.xlsx");
                ObjWorkSheetZapol = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookZapol.Sheets[1];
                var nametamoznStroka = 3;
                var nametamoznStolbec = 1;

                var namedoljnostStroka = 3;
                var namedoljnostStolbec = 2;

                var kolvoshtatStroka = 3;
                var kolvoshtatStolbec = 3;

                var kolvofaktStroka = 3;
                var kolvofaktStolbec = 4;

                
                ObjWorkSheetZapol.Columns["A:A"].ColumnWidth = 23.86;
                ObjWorkSheetZapol.Columns["B:B"].ColumnWidth = 26.43;
                ObjWorkSheetZapol.Columns["C:C"].ColumnWidth = 11.43;
                ObjWorkSheetZapol.Range["A1"].Formula = String.Format("Сведения о штатной численности работников таможенных органов и аппарата Управления по состоянию на {0} и фактической численности на {1}", Date1.ToString("dd MMMM yyyy"), Date1.ToString("dd MMMM yyyy"));
                

                //=======для заполнения====





                Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //открыть эксель
                Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory + @"Templates\Должности.xlsx");
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
                var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);//1 ячейку
                int numCol = 1;

                Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
                System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
                string[] arrDoljnosti = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();


                int numRow = 1;
                Range usedRow = ObjWorkSheet.UsedRange.Rows[numRow];
                System.Array myvaluesRow = (System.Array)usedRow.Cells.Value2;
                string[] arrPodrazdel = myvaluesRow.OfType<object>().Select(o => o.ToString()).ToArray();
                
                Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlsBook = xlsApp.Workbooks.Open(Convert.ToString(FileName));
                Worksheet xlsSheet;
                xlsSheet = (Worksheet)xlsBook.Sheets[1]; // раздел 1 (лист 1)
                xlsSheet.Activate();
                var MyRange = "I:I";
                //this.Dis
                //разделяем  столбец Кол-во/Дата Образования
                xlsApp.DisplayAlerts = false;
                xlsSheet.Range["I:I"].TextToColumns(xlsSheet.get_Range(MyRange, Type.Missing), XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierDoubleQuote, true, Type.Missing, Type.Missing, true, false, Type.Missing, ",", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlsSheet.Cells.Replace(".5", "0,5", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-1", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-2", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Cells.Replace("-.5", "0", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);
                xlsSheet.Range["G:G"].Replace("0,5", "0,5", XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false);

                xlsSheet.Range["P2"].Formula = "=SUBTOTAL(109,RC[-7]:R[1594]C[-7])";
                xlsSheet.Range["O2"].Formula = "=SUBTOTAL(109,RC[-8]:R[1825]C[-8])";
                Directory.CreateDirectory(@"D:\NewTestSpravDol\Otladka");
                Directory.CreateDirectory(@"D:\Справки");

                double obshSumShtat = 0;
                double obshSumFakt = 0;


                int KolonStat = 2;
                int KolonFakt = 3;
                int Stroka = 3;
                string pathSaveNewRezultTy = String.Format(@"D:\Справки\ЦТУ_Работники_ШТАТ_ФАКТ_{0}-{1}-{2}_{3}%{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);
                string pathSaveNewRezult = String.Format(@"D:\Справки\Результат_Должности_{0}-{1}-{2}_{3}%{4}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute);

                int arrdlinaDolj =  (arrDoljnosti.Length)+1;
                ObjWorkExcel.DisplayAlerts = false;
                
                for (int i=0; i < arrPodrazdel.LongLength; i++)
                {
                    
                    xlsSheet.Range["$A$1:$G$8521"].AutoFilter(1, arrPodrazdel[i], XlAutoFilterOperator.xlOr);
                    Stroka = 3;
                    double Shtat = 0;
                    double Fakt = 0;
                    double SumShtat = 0;
                    double SumFakt = 0;
                    

                    for (int j = 0; j < arrDoljnosti.Length; j++)
                    {
                       // string pathSaveNew = String.Format(@"D:\NewTestSpravDol\Otladka\{0}_{1}_{2}_{3}_{4}_{5}_{6}",Convert.ToString(arrDoljnosti[j]), DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond);


                        xlsSheet.Range["$A$1:$G$8521"].AutoFilter(2, arrDoljnosti[j], XlAutoFilterOperator.xlOr);
                        
                        Shtat = xlsSheet.Range["O2"].Value;
                        Fakt = xlsSheet.Range["P2"].Value;
                        ObjWorkSheet.Cells[Stroka, KolonStat] = Shtat;
                        ObjWorkSheet.Cells[Stroka, KolonFakt] = Fakt;
                        Stroka++;

                     
                        var NameTamozn = Convert.ToString(arrPodrazdel[i]);
                        var NameDoljnost = Convert.ToString(arrDoljnosti[j]);

                        
                        if (Shtat != 0)
                        {
                            ObjWorkSheetZapol.Cells[nametamoznStroka, nametamoznStolbec] = NameTamozn;
                            ObjWorkSheetZapol.Cells[namedoljnostStroka, namedoljnostStolbec] = NameDoljnost;

                            ObjWorkSheetZapol.Cells[kolvoshtatStroka, kolvoshtatStolbec] = Shtat;
                            ObjWorkSheetZapol.Cells[kolvofaktStroka, kolvofaktStolbec] = Shtat- Fakt;
                            
                            nametamoznStroka++;
                            namedoljnostStroka++;
                            kolvoshtatStroka++;
                            kolvofaktStroka++;

                            SumShtat = SumShtat + Shtat;
                            SumFakt = SumFakt + (Shtat - Fakt);

                        }
                        

                    }
                    obshSumShtat = obshSumShtat + SumShtat;
                    obshSumFakt = obshSumFakt + SumFakt;
                    ObjWorkSheetZapol.Cells[nametamoznStroka, nametamoznStolbec] = "Итого по таможне:";
                    ObjWorkSheetZapol.Cells[nametamoznStroka, nametamoznStolbec].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;

                    ObjWorkSheetZapol.Cells[kolvoshtatStroka, kolvoshtatStolbec] = SumShtat;
                    ObjWorkSheetZapol.Cells[kolvofaktStroka, kolvofaktStolbec] = SumFakt;
                    ObjWorkSheetZapol.Cells[kolvofaktStroka, kolvofaktStolbec].Interior.ColorIndex = 6; // жёлтый
                    ObjWorkSheetZapol.Cells[kolvoshtatStroka, kolvoshtatStolbec].Interior.ColorIndex = 6; // жёлтый
                    
                    SumShtat = 0;
                    SumFakt = 0;
                    nametamoznStroka++;
                    namedoljnostStroka++;
                    kolvoshtatStroka++;
                    kolvofaktStroka++;
                    

                    KolonStat += 2;
                    KolonFakt+=2;
                    
                }
                ObjWorkSheetZapol.Cells[nametamoznStroka, nametamoznStolbec] = "Общий итог:";
                ObjWorkSheetZapol.Cells[nametamoznStroka, nametamoznStolbec].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                ObjWorkSheetZapol.Cells[kolvoshtatStroka, kolvoshtatStolbec] = obshSumShtat;
                ObjWorkSheetZapol.Cells[kolvofaktStroka, kolvofaktStolbec] = obshSumFakt;
                ObjWorkSheetZapol.Cells[kolvofaktStroka, kolvofaktStolbec].Interior.ColorIndex = 46; // оранж
                ObjWorkSheetZapol.Cells[kolvoshtatStroka, kolvoshtatStolbec].Interior.ColorIndex = 46; // оранж
                ObjExcelZapol.DisplayAlerts = false;
                ObjWorkBookZapol.SaveAs(pathSaveNewRezultTy, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                ObjWorkBook.SaveAs(pathSaveNewRezult, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
                DialogResult dialogResult = MessageBox.Show("Открыть созданный файл?", "Открытие файла", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    ObjWorkExcel.Visible = true;
                    ObjWorkExcel.UserControl = true;
                }
                MessageBox.Show("Сохранён по адресу: \nD:\\Справки\\", "Файл успешно сформирован");

                /*form1.progress.Value = (90*15);
                form1*/
                ObjWorkExcel.Quit();
                GC.Collect(); // убрать за собой


           }
           catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            /*foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
            {
                currentProcess.Kill();
            }*/
        }

      
    }
    
}
