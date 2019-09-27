using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelR
{
    public partial class ShtatFaktView : Form
    {
        //  public static List<COkFriendListDb> _LstOkFriendListDb = new List<COkFriendListDb>();
        static MassivDannyh massiv = new MassivDannyh();
        static public int KOLSHTAT_SOTR;
        static public int KOLSHTAT_RAB;
        static public int KOLSHTAT_GOS;

        static public int KOLFAKT_SOTR;
        static public int KOLFAKT_RAB;
        static public int KOLFAKT_GOS;

        static public double KOLVAK_SOTR;
        static public double KOLVAK_RAB;
        static public double KOLVAK_GOS;


        static public double SLUJ_KOLSHTAT_SOTR;
        static public double SLUJ_KOLSHTAT_RAB;
        static public double SLUJ_KOLSHTAT_GOS;

        static public int SLUJ_KOLVAK_SOTR;
        static public int SLUJ_KOLVAK_RAB;
        static public int SLUJ_KOLVAK_GOS;
        static public int SLUJ_KOLFAKT_SOTR;
        static public int SLUJ_KOLFAKT_RAB;
        static public int SLUJ_KOLFAKT_GOS;


        static public double SUM_SHTAT_GOS;
        static public double SUM_FAKT_GOS;
        static public double SUM_SHTAT_SOTR;
        static public double SUM_FAKT_SOTR;
        static public double SUM_SHTAT_RAB;
        static public double SUM_FAKT_RAB;

        static public double FULL_SUM_GOS_SHTAT =0;
        static public double FULL_SUM_GOS_FAKT =0;
        static public double FULL_SUM_RAB_SHTAT =0;
        static public double FULL_SUM_RAB_FAKT =0;
        static public double FULL_SUM_SOTR_SHTAT =0;
        static public double FULL_SUM_SOTR_FAKT =0;
        static public double FULL_SUM_SHTAT =0;
        private static Range Excelcells;

        static public string str;
        static public int XZ;
        static public bool tt;


        static public int SUM_SOVETNIK=0;
        static public int SUM_REFERENT=0;
        static public int SUM_SEKRETAR=0;
        static public int SUM_Vysh_Nach_sostav = 0;
        static public int SUM_Starsh_Nach_sostav = 0;
        static public int SUM_Srednii_Nach_sostav = 0;
        static public int SUM_Mladshii_sostav = 0;
        static public int SUM_Rykovod = 0;
        static public int SUM_Spec_Slyj = 0;
        static public int SUM_Rabochie = 0;
        

        static public int SUM_SOVETNIK_FAKT = 0;
        static public int SUM_REFERENT_FAKT = 0;
        static public int SUM_SEKRETAR_FAKT = 0;
        static public int SUM_Vysh_Nach_sostav_FAKT = 0;
        static public int SUM_Starsh_Nach_sostav_FAKT = 0;
        static public int SUM_Srednii_Nach_sostav_FAKT = 0;
        static public int SUM_Mladshii_sostav_FAKT = 0;
        static public int SUM_Rykovod_FAKT = 0;
        static public int SUM_Spec_Slyj_FAKT = 0;
        static public int SUM_Rabochie_FAKT = 0;

        static public double FULL_SUMM_DOP_GOS=0;
        static public double FULL_SUMM_DOP_SOTR=0;
        static public double FULL_SUMM_DOP_RAB=0;

        static public int arrTM_depart;
        static public string arrTM_TO;
        static public int arrTM_RN;

        static public double sum_fakt_full;
        //ShtatFaktView sh = new ShtatFaktView();

        public ShtatFaktView()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            massiv._Dolj.Clear();
            massiv._Otdel.Clear();
            massiv._Slujba.Clear();
            massiv._SlujDolj.Clear();
            massiv._Svedenia.Clear();
            massiv._SotrudnikiNizBlok.Clear();
            massiv._RabotnikiBlok.Clear();
            massiv._RabotnikiBlokFact.Clear();


             KOLSHTAT_SOTR = 0;
             KOLSHTAT_RAB = 0;
             KOLSHTAT_GOS = 0;

             KOLFAKT_SOTR = 0;
             KOLFAKT_RAB = 0;
             KOLFAKT_GOS = 0;

            KOLVAK_SOTR = 0;
            KOLVAK_RAB = 0;
            KOLVAK_GOS = 0;


            SLUJ_KOLSHTAT_SOTR = 0;
            SLUJ_KOLSHTAT_RAB = 0;
            SLUJ_KOLSHTAT_GOS = 0;

             SLUJ_KOLVAK_SOTR = 0;
             SLUJ_KOLVAK_RAB = 0;
             SLUJ_KOLVAK_GOS = 0;
             SLUJ_KOLFAKT_SOTR = 0;
             SLUJ_KOLFAKT_RAB = 0;
             SLUJ_KOLFAKT_GOS = 0;


            SUM_SHTAT_GOS = 0;
            SUM_FAKT_GOS = 0;
            SUM_SHTAT_SOTR = 0;
            SUM_FAKT_SOTR = 0;
            SUM_SHTAT_RAB = 0;
            SUM_FAKT_RAB = 0;

            FULL_SUM_GOS_SHTAT = 0;
            FULL_SUM_GOS_FAKT = 0;
            FULL_SUM_RAB_SHTAT = 0;
            FULL_SUM_RAB_FAKT = 0;
            FULL_SUM_SOTR_SHTAT = 0;
            FULL_SUM_SOTR_FAKT = 0;
            FULL_SUM_SHTAT = 0;
        

         str = null;
         XZ = 0;
         tt = false;


         SUM_SOVETNIK = 0;
         SUM_REFERENT = 0;
         SUM_SEKRETAR = 0;
         SUM_Vysh_Nach_sostav = 0;
         SUM_Starsh_Nach_sostav = 0;
         SUM_Srednii_Nach_sostav = 0;
         SUM_Mladshii_sostav = 0;
         SUM_Rykovod = 0;
         SUM_Spec_Slyj = 0;
         SUM_Rabochie = 0;


         SUM_SOVETNIK_FAKT = 0;
         SUM_REFERENT_FAKT = 0;
         SUM_SEKRETAR_FAKT = 0;
         SUM_Vysh_Nach_sostav_FAKT = 0;
         SUM_Starsh_Nach_sostav_FAKT = 0;
         SUM_Srednii_Nach_sostav_FAKT = 0;
         SUM_Mladshii_sostav_FAKT = 0;
         SUM_Rykovod_FAKT = 0;
         SUM_Spec_Slyj_FAKT = 0;
         SUM_Rabochie_FAKT = 0;

        FULL_SUMM_DOP_GOS = 0;
        FULL_SUMM_DOP_SOTR = 0;
        FULL_SUMM_DOP_RAB = 0;
                   

        sum_fakt_full=0;


        MassivDannyh md = new MassivDannyh();
            var rr = md.RN;
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"C:\Для ОШ для БАС.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            //-------------------------------------
            int lastColumn = (int)lastCell.Column;//!сохраним непосредственно требующееся в дальнейшем
            int lastRow = (int)lastCell.Row;
            //-------------------------------------
            string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу

            /*  for (int i = 0; i < (int)lastCell.Column; i++) //по всем колонкам
                  for (int j = 0; j < (int)lastCell.Row; j++) // по всем строкам
                      list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку
                      */

            int iLastRow = ObjWorkSheet.Cells[ObjWorkSheet.Rows.Count, "E"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
            var arrData = (object[,])ObjWorkSheet.Range["E2:E" + iLastRow].Value;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой -- в том числе не используемые явно объекты !
            for (int i = 1; i < lastColumn; i++) //по всем колонкам
                for (int j = 1; j < lastRow; j++) // по всем строкам 
                    Console.Write(list[i, j]);//выводим строку
            Console.ReadLine();





            
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            
            DBOracleUtils db = new DBOracleUtils();
          OracleConnection conn =  db.GetDBConnection();

            


            try
            {
                conn.Open();
                sum_fakt_full = 0;
                load.Visible = true;
                ojidan.Text = "Ожидание . . .";
                ojidan.Visible = true;
                var rr = await Task.Run(() => Query(conn));
               if (rr == 1)
                {
                    load.Visible = false;
                    ojidan.Text = "Выполнено!";
                   
                }
                
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                    "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           /* finally
            {
                conn.Close();
                conn.Dispose();
            }*/
        }
       /* public async int QueryAsync(OracleConnection conn)
        {

            //var result = Task.Run(async () => { return await Query(conn); }).Result;
            var res = await Task.Run(() => Query(conn));
            //var tt = res;
            return res;
        }*/

        private  int Query (OracleConnection conn)
        {
            massiv._Dolj.Clear();
            massiv._Otdel.Clear();
            massiv._Slujba.Clear();
            massiv._SlujDolj.Clear();
            massiv._Svedenia.Clear();
            massiv._SotrudnikiNizBlok.Clear();
            massiv._RabotnikiBlok.Clear();
            massiv._RabotnikiBlokFact.Clear();
            //LoadingNew load = new LoadingNew();
            // load.Visible = true;

            SUM_SOVETNIK = 0;
            SUM_REFERENT = 0;
            SUM_SEKRETAR = 0;
            SUM_Vysh_Nach_sostav = 0;
            SUM_Starsh_Nach_sostav = 0;
            SUM_Srednii_Nach_sostav = 0;
            SUM_Mladshii_sostav = 0;
            SUM_Rykovod = 0;
            SUM_Spec_Slyj = 0;
            SUM_Rabochie = 0;
            Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //сам эксель
            string pathSaveSpisokLicz = String.Format(@"D:\Справки\ОргШтатФакт_" + arrTM_TO);
            string FilNam = (AppDomain.CurrentDomain.BaseDirectory + @"Templates\ШтатФакт.xlsx");
            Excel.Worksheet WS;// в этот лист
            Workbook WB = ObjWorkExcel.Workbooks.Open(Convert.ToString(FilNam));
            int RowCount = 0;
            WS = (Excel.Worksheet)WB.Sheets[1];
            WS.Activate();
            ObjWorkExcel.DisplayAlerts = false;
            //Считывание с определенных ячеек
            string NameTamozn = Convert.ToString(WS.Range["A1"].Value);

            int LastRow;
            int LastRowOtd;
            string addr;
             string sqlSlj = "SELECT IE.dep, IE.urlev,id.NAME, id.NAME_NOM, id.SHORTNAME_NOM, id.CODE, id.RN, id.DEPART_DISP FROM UK_PARUS.INS_DEPARTMENT id INNER JOIN (SELECT id3.RN,  LEVEL AS urlev, id3.PRN AS dep, id3.DEPART_DISP  FROM UK_PARUS.INS_DEPARTMENT id3 START WITH id3.RN =" + arrTM_RN + "CONNECT BY PRIOR id3.RN = id3.PRN) IE ON id.RN = IE.RN  WHERE id.BGNDATE <= TRUNC(SYSDATE) AND (id.ENDDATE >= TRUNC(SYSDATE) OR id.ENDDATE IS NULL) AND IE.urlev = 2 ORDER BY ID.NAME asc";
            // Создать объект Command.
            OracleCommand cmd = new OracleCommand();

            // Сочетать Command с Connection.
            cmd.Connection = conn;
            cmd.CommandText = sqlSlj;
            


            try
            {
                int stroka=0;
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        massiv._Slujba.Add(new Slujba()
                        {

                            dep = (reader["dep"].ToString()),
                            urlev = int.Parse((reader["urlev"].ToString())),
                            NAME = (reader["NAME"].ToString()),
                            NAME_NOM = (reader["NAME_NOM"].ToString()),
                            SHORTNAME_NOM = (reader["SHORTNAME_NOM"].ToString()),
                            CODE = (reader["CODE"].ToString()),
                            RN = int.Parse((reader["RN"].ToString())),
                            DEPART_DISP = (reader["DEPART_DISP"].ToString()),

                        });
                    }
                    conn.Close();

                }

                foreach (Slujba sl in massiv._Slujba) // пробегаем по службам
                {
                    WS = (Excel.Worksheet)WB.Sheets[1];//выбираем 1 лист
                    LastRow = WB.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                    string Proverka = "A" + LastRow;
                    string ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                    if (ZnachVal != null)
                    {
                        LastRow++;
                        addr = "A" + LastRow + ":" + "G" + LastRow;
                        WS.get_Range(addr).EntireRow.Insert(XlDirection.xlDown);
                        //adr = "A" + stroka + ":" + "G" + stroka;
                        WS.get_Range(addr).Font.Name = "Times New Roman";
                        WS.get_Range(addr).Font.Size = 12;
                        WS.get_Range(addr).Font.Bold = true;
                        //WS.get_Range(addr).Font.Bold = true;
                    }

                    if (sl.RN == 453968139)
                    { var ii = 0; }
                    WS.Cells[LastRow, 1] = sl.NAME_NOM;
                    QuSvedenia(conn, arrTM_depart);
                    //if (sl.RN == 453968139)
                        QuOtdel(sl.RN, conn);
                    QuSluDolj(conn, sl.DEPART_DISP, arrTM_depart);
                   
                    foreach (Otdel otd in massiv._Otdel) // пробегаем по отделам
                    {

                        QuDolj(conn, otd.DEPART_DISP, arrTM_depart);
                        LastRowOtd = WB.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // находим в ней последнюю строку
                        Proverka = "A" + LastRowOtd;
                        ZnachVal = Convert.ToString(WS.Range[Proverka].Value);

                        if (ZnachVal != null)
                        {
                            LastRowOtd++;
                            //addr = "A" + LastRowOtd;
                            addr = "A" + LastRowOtd + ":" + "G" + LastRowOtd;
                            WS.get_Range(addr).EntireRow.Insert(XlDirection.xlDown);
                            WS.get_Range(addr).Font.Name = "Times New Roman";
                            WS.get_Range(addr).Font.Size = 12;
                            //WS.get_Range(addr).Font.Bold = true;
                            WS.get_Range(addr).Font.Bold = false;
                        }

                        foreach (Dolj dol in massiv._Dolj)
                        {
                            if (dol.SLVID == "Работник")
                            {
                                KOLSHTAT_RAB += dol.MAIN_KOLSHTAT;
                                tt = dol.MAIN_PIPEC.Contains("1,");
                                if (tt)
                                {
                                    XZ++;
                                }
                                                                
                                                               
                                    var Polovinki = dol.MAIN_PIPEC.Contains(",5,");


                                if (Polovinki & (dol.DEP_CODE == "Вакантна"))
                                {
                                    var pol = 0.5;
                                    KOLVAK_RAB = KOLVAK_RAB + pol;
                                }

                                if ((dol.DEP_CODE == "Вакантна") & (dol.MAIN_KOLSHTAT > 0) & (dol.AGN_RN == ""))
                                {

                                        KOLVAK_RAB++;

                                }

                            }
                            if (dol.SLVID == "Госслужащий")
                            {
                                KOLSHTAT_GOS += dol.MAIN_KOLSHTAT;

                                if ((dol.DEP_CODE == "Вакантна") & (dol.MAIN_KOLSHTAT > 0))
                                {
                                    KOLVAK_GOS++;
                                }
                            }
                            if (dol.SLVID == "Сотрудник")
                            {
                                KOLSHTAT_SOTR += dol.MAIN_KOLSHTAT;

                                if ((dol.DEP_CODE == "Вакантна") & (dol.MAIN_KOLSHTAT > 0))
                                {
                                    KOLVAK_SOTR++;
                                }

                            }

                        }
                        WS.Cells[LastRowOtd, 1] = otd.NAME_NOM;

                        WS.Cells[LastRowOtd, 2] = KOLSHTAT_GOS;
                        WS.Cells[LastRowOtd, 4] = KOLSHTAT_SOTR;
                        WS.Cells[LastRowOtd, 6] = KOLSHTAT_RAB;

                        WS.Cells[LastRowOtd, 3] = KOLSHTAT_GOS - KOLVAK_GOS;
                        WS.Cells[LastRowOtd, 5] = KOLSHTAT_SOTR - KOLVAK_SOTR;
                        WS.Cells[LastRowOtd, 7] = KOLSHTAT_RAB - KOLVAK_RAB;

                        SUM_SHTAT_GOS += KOLSHTAT_GOS;
                        SUM_SHTAT_RAB += KOLSHTAT_RAB;
                        SUM_SHTAT_SOTR += KOLSHTAT_SOTR;

                        SUM_FAKT_GOS += KOLSHTAT_GOS - KOLVAK_GOS;
                        SUM_FAKT_RAB += KOLSHTAT_RAB - KOLVAK_RAB;
                        SUM_FAKT_SOTR += KOLSHTAT_SOTR - KOLVAK_SOTR;


                        KOLSHTAT_GOS = 0;
                        KOLSHTAT_SOTR = 0;
                        KOLSHTAT_RAB = 0;
                        KOLVAK_SOTR = 0;
                        KOLVAK_GOS = 0;
                        KOLVAK_RAB = 0;
                        stroka = LastRowOtd;

                    }

                    
                    foreach (SlujDolj dol in massiv._SlujDolj)
                    {
                        if (dol.SLVID == "Работник")
                        {
                            SLUJ_KOLSHTAT_RAB += dol.MAIN_KOLSHTAT;
                            string patt = "1*";
                            //   var str = dol.MAIN_PIPEC.ToString();
                            //  var res = dol.MAIN_PIPEC.Where(str => Regex.IsMatch(Convert.ToString(str), patt));
                            tt = dol.MAIN_PIPEC.Contains("1,");
                            if (tt)
                            {
                                XZ++;
                            }
                            if ((dol.DEP_CODE == "Вакантна") & (dol.MAIN_KOLSHTAT > 0) & (dol.AGN_RN==""))
                            {
                                SLUJ_KOLVAK_RAB++;
                            }

                        }
                        if (dol.SLVID == "Госслужащий")
                        {
                            SLUJ_KOLSHTAT_GOS += dol.MAIN_KOLSHTAT;

                            if ((dol.DEP_CODE == "Вакантна") & (dol.MAIN_KOLSHTAT > 0))
                            {
                                SLUJ_KOLVAK_GOS++;
                            }
                        }
                        if (dol.SLVID == "Сотрудник")
                        {
                            SLUJ_KOLSHTAT_SOTR += dol.MAIN_KOLSHTAT;

                            if ((dol.DEP_CODE == "Вакантна") & (dol.MAIN_KOLSHTAT > 0))
                            {
                                SLUJ_KOLVAK_SOTR++;
                            }

                        }

                    }
                    
                    WS.Cells[LastRow, 2] = SLUJ_KOLSHTAT_GOS + SUM_SHTAT_GOS;
                    WS.Cells[LastRow, 4] = SLUJ_KOLSHTAT_SOTR + SUM_SHTAT_SOTR;
                    WS.Cells[LastRow, 6] = SLUJ_KOLSHTAT_RAB + SUM_SHTAT_RAB;

                    WS.Cells[LastRow, 3] = (SLUJ_KOLSHTAT_GOS - SLUJ_KOLVAK_GOS)+ SUM_FAKT_GOS;
                    WS.Cells[LastRow, 5] = (SLUJ_KOLSHTAT_SOTR - SLUJ_KOLVAK_SOTR) + SUM_FAKT_SOTR;
                    WS.Cells[LastRow, 7] = (SLUJ_KOLSHTAT_RAB - SLUJ_KOLVAK_RAB) + SUM_FAKT_RAB;




                    FULL_SUM_GOS_SHTAT+= SLUJ_KOLSHTAT_GOS + SUM_SHTAT_GOS;
                    FULL_SUM_GOS_FAKT += (SLUJ_KOLSHTAT_GOS - SLUJ_KOLVAK_GOS) + SUM_FAKT_GOS;
                    FULL_SUM_RAB_SHTAT += SLUJ_KOLSHTAT_RAB + SUM_SHTAT_RAB;
                    FULL_SUM_RAB_FAKT += (SLUJ_KOLSHTAT_RAB - SLUJ_KOLVAK_RAB) + SUM_FAKT_RAB;
                    FULL_SUM_SOTR_SHTAT += SLUJ_KOLSHTAT_SOTR + SUM_SHTAT_SOTR;
                    FULL_SUM_SOTR_FAKT+= (SLUJ_KOLSHTAT_SOTR - SLUJ_KOLVAK_SOTR) + SUM_FAKT_SOTR;

                    

                    SLUJ_KOLSHTAT_GOS = 0;
                    SLUJ_KOLSHTAT_SOTR = 0;
                    SLUJ_KOLSHTAT_RAB = 0;
                    SLUJ_KOLVAK_SOTR = 0;
                    SLUJ_KOLVAK_GOS = 0;
                    SLUJ_KOLVAK_RAB = 0;
                    SUM_SHTAT_GOS = 0;
                    SUM_SHTAT_RAB = 0;
                    SUM_SHTAT_SOTR = 0;

                    SUM_FAKT_GOS = 0;
                    SUM_FAKT_RAB = 0;
                    SUM_FAKT_SOTR = 0;

                                 

                }

                sum_fakt_full = FULL_SUM_SOTR_FAKT + FULL_SUM_RAB_FAKT + FULL_SUM_GOS_FAKT;

                stroka = WB.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                string adr = "A" + stroka + ":" + "G" + stroka;
                WS.get_Range(adr).Font.Name = "Times New Roman";
                WS.get_Range(adr).Font.Size = 12;
                WS.get_Range(adr).Font.Bold = true;

                WS.Cells[stroka, 1] = "ИТОГО:";
                WS.Cells[stroka, 2] = FULL_SUM_GOS_SHTAT;
                WS.Cells[stroka, 4] = FULL_SUM_SOTR_SHTAT;
                WS.Cells[stroka, 6] = FULL_SUM_RAB_SHTAT;

                WS.Cells[stroka, 3] = FULL_SUM_GOS_FAKT;
                WS.Cells[stroka, 5] = FULL_SUM_SOTR_FAKT;
                WS.Cells[stroka, 7] = FULL_SUM_RAB_FAKT;
                stroka=WB.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                stroka++;
                adr = "A" + stroka + ":" + "G" + stroka;
                WS.get_Range(adr).Font.Name = "Times New Roman";
                WS.get_Range(adr).Font.Size = 12;
                WS.get_Range(adr).Font.Bold = true;
                WS.Cells[stroka, 1] = "ВСЕГО штат|факт:";
                adr = "B" + stroka + ":" + "G" + stroka;
                WS.get_Range(adr).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                WS.Cells[stroka, 2] = FULL_SUM_GOS_SHTAT + FULL_SUM_SOTR_SHTAT + FULL_SUM_RAB_SHTAT;
                WS.Cells[stroka, 3] = sum_fakt_full;
                WS.Cells[stroka, 4] = "";
                WS.Cells[stroka, 5] = "";
                WS.Cells[stroka, 6] = "";
                WS.Cells[stroka, 7] = "";
                stroka++;
                //заполнение другого блока
                QuSvedRab(conn, arrTM_depart);
                QuSvedRabotniki(conn, arrTM_RN);
                QuSvedRabotnikiFakt(conn, arrTM_depart);

               foreach (SotrudnikiNizBlok svedSotr in massiv._SotrudnikiNizBlok)
                {
                    

                         if ((svedSotr.SLVID == "Сотрудник") & (svedSotr.MAIN_KOLSHTAT == 1))
                         {

                            //Высший
                            if ((svedSotr.CLNZVAN == "Генерал-лейтенант таможенной службы") | (svedSotr.CLNZVAN == "Генерал-майор таможенной службы") | (svedSotr.CLNZVAN == "Генерал-полковник таможенной службы"))
                            {
                                SUM_Vysh_Nach_sostav++;
                            }


                            //Старший
                            if ((svedSotr.CLNZVAN == "Майор таможенной службы") | (svedSotr.CLNZVAN == "Подполковник таможенной службы") | (svedSotr.CLNZVAN == "Полковник таможенной службы"))
                            {
                                SUM_Starsh_Nach_sostav++;
                            }

                            //Средний
                            if ((svedSotr.CLNZVAN == "Капитан таможенной службы") | (svedSotr.CLNZVAN == "Старший лейтенант таможенной службы") | (svedSotr.CLNZVAN == "Лейтенант таможенной службы") | (svedSotr.CLNZVAN == "Младший лейтенант таможенной службы"))
                            {
                                SUM_Srednii_Nach_sostav++;
                            }


                            //младший
                            if ((svedSotr.CLNZVAN == "Старший прапорщик таможенной службы") | (svedSotr.CLNZVAN == "Прапорщик таможенной службы"))
                            {
                                SUM_Mladshii_sostav++;
                            }
                         }



                    if ((svedSotr.SLVID == "Сотрудник") & (svedSotr.DEP_CODE != "Вакантна"))
                    {

                        //Высший
                        if ((svedSotr.CLNZVAN == "Генерал-лейтенант таможенной службы") | (svedSotr.CLNZVAN == "Генерал-майор таможенной службы") | (svedSotr.CLNZVAN == "Генерал-полковник таможенной службы"))
                        {
                            SUM_Vysh_Nach_sostav_FAKT++;
                        }


                        //Старший
                        if ((svedSotr.CLNZVAN == "Майор таможенной службы") | (svedSotr.CLNZVAN == "Подполковник таможенной службы") | (svedSotr.CLNZVAN == "Полковник таможенной службы"))
                        {
                            SUM_Starsh_Nach_sostav_FAKT++;
                        }

                        //Средний
                        if ((svedSotr.CLNZVAN == "Капитан таможенной службы") | (svedSotr.CLNZVAN == "Старший лейтенант таможенной службы") | (svedSotr.CLNZVAN == "Лейтенант таможенной службы") | (svedSotr.CLNZVAN == "Младший лейтенант таможенной службы"))
                        {
                            SUM_Srednii_Nach_sostav_FAKT++;
                        }


                        //младший
                        if ((svedSotr.CLNZVAN == "Старший прапорщик таможенной службы") | (svedSotr.CLNZVAN == "Прапорщик таможенной службы"))
                        {
                            SUM_Mladshii_sostav_FAKT++;
                        }
                    }


                }
                
               foreach (RabotnikiBlok rabblolk in massiv._RabotnikiBlok)
                {


                    if (rabblolk.CatDolj == "Руководители")//Руководители
                    {
                        SUM_Rykovod++;

                    }
                    if (rabblolk.CatDolj == "Специалисты и служащие") //Спец и Служ
                    {
                        SUM_Spec_Slyj++;
                    }
                    if (rabblolk.CatDolj == "Рабочие")//Рабочие
                    {
                        SUM_Rabochie++;
                    }

                    foreach (RabotnikiBlokFact rabblokfact in massiv._RabotnikiBlokFact)
                    {
                        if ((rabblokfact.AGN_RN != "") & (rabblokfact.NUMB == rabblolk.NUMB) & (rabblolk.CatDolj == "Руководители"))
                        {
                            SUM_Rykovod_FAKT++;
                        }
                        if ((rabblokfact.AGN_RN != "") & (rabblokfact.NUMB == rabblolk.NUMB) & (rabblolk.CatDolj == "Специалисты и служащие"))
                        {
                            SUM_Spec_Slyj_FAKT++;
                        }

                        if ((rabblokfact.AGN_RN != "") & (rabblokfact.NUMB == rabblolk.NUMB) & (rabblolk.CatDolj == "Рабочие"))
                        {
                            SUM_Rabochie_FAKT++;
                        }
                    }

                }

                    foreach (Svedenia sved in massiv._Svedenia)
                {
                    var zvanSovetnik = "советник государственной гражданской";
                    var zvanReferent = "референт государственной гражданской";
                    var zvanSekretar = "секретарь государственной гражданской";

                    var zvanVish = "Высший начальствующий состав";
                    var zvanStarsh = "Старший начальствующий состав";
                    var zvanSrednii = "Средний начальствующий состав";
                    var zvanMladsh = "Младший состав";

                    var zvanRykov = "Руководители";
                    var zvanSpecSlyj = "Специалисты и служащие";
                    var zvanRab = "Рабочие";

                    var indexOfCharSovetnik = sved.CLNZVAN.IndexOf(zvanSovetnik);
                    var indexOfCharReferent = sved.CLNZVAN.IndexOf(zvanReferent);
                    var indexOfCharSecretar = sved.CLNZVAN.IndexOf(zvanSekretar);

                    var indexOfCharVish = sved.SOSTAV.IndexOf(zvanVish);
                    var indexOfCharStarsh = sved.SOSTAV.IndexOf(zvanStarsh);
                    var indexOfCharSrednii = sved.SOSTAV.IndexOf(zvanSrednii);
                    var indexOfCharMladsh = sved.SOSTAV.IndexOf(zvanMladsh);

                    var indexOfCharRykov = sved.CLNKAD.IndexOf(zvanRykov);
                    var indexOfCharSpecSlyj = sved.CLNKAD.IndexOf(zvanSpecSlyj);
                    var indexOfCharRab = sved.CLNKAD.IndexOf(zvanRab);

                    if (sved.SLVID == "Госслужащий")
                    {

                        if (indexOfCharSovetnik > -1)  //советник
                        {
                            SUM_SOVETNIK_FAKT++;
                        }
                        if (indexOfCharReferent > -1)  //референт
                        {
                            SUM_REFERENT_FAKT++;
                        }
                        if (indexOfCharSecretar > -1) //секретарь
                        {
                            SUM_SEKRETAR_FAKT++;
                        }

                        
                        if ((indexOfCharSovetnik > -1) & (sved.DEP_CODE != "Вакантна")) //советник
                        {
                            SUM_SOVETNIK++;
                        }
                        if ((indexOfCharReferent > -1) & (sved.DEP_CODE != "Вакантна")) //референт
                        {
                            SUM_REFERENT++;
                        }
                        if ((indexOfCharSecretar > -1) & (sved.DEP_CODE != "Вакантна")) //секретарь
                        {
                            SUM_SEKRETAR++;
                        }
                    }
                   
                   
                    

                }

                stroka = WB.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                stroka++;
                adr = "A" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).Font.Name = "Times New Roman";
                WS.get_Range(adr).Font.Size = 12;
                WS.get_Range(adr).Font.Bold = true;

                WS.Cells[stroka, 1] = "Государственные служащие";
                //WS.Cells[stroka, 2] = 0;
                adr = "B" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).HorizontalAlignment = XlHAlign.xlHAlignRight;
                WS.Cells[stroka, 3] = SUM_SOVETNIK+SUM_REFERENT+SUM_SEKRETAR;
                WS.Cells[stroka, 2] = SUM_SOVETNIK_FAKT + SUM_REFERENT_FAKT + SUM_SEKRETAR_FAKT;
                WS.Cells[stroka, 4] = "";
                WS.Cells[stroka, 5] = "";
                WS.Cells[stroka, 6] = "";
                WS.Cells[stroka, 7] = "";

                stroka++;
                adr = "A" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).Font.Bold = false;
                WS.Cells[stroka, 1] = "в том числе: ";
                WS.Cells[stroka, 2] = "";
                WS.Cells[stroka, 3] = "";
                //WS.Cells[stroka, 3] = "";
                WS.Cells[stroka, 4] = "";
                WS.Cells[stroka, 5] = "";
                WS.Cells[stroka, 6] = "";
                WS.Cells[stroka, 7] = "";
                stroka++;
                adr = "A" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).Font.Bold = false;
                adr = "B" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).HorizontalAlignment = XlHAlign.xlHAlignRight;
                WS.Cells[stroka, 1] = "Советник государственной гражданской службы  Российской Федерации 1 - 3 класса";
                WS.Cells[stroka, 1].WrapText = true;
                WS.Cells[stroka, 2] = SUM_SOVETNIK_FAKT;
                WS.Cells[stroka, 3] = SUM_SOVETNIK;
                //WS.Cells[stroka, 3] = "";
                WS.Cells[stroka, 4] = "";
                WS.Cells[stroka, 5] = "";
                WS.Cells[stroka, 6] = "";
                WS.Cells[stroka, 7] = "";
                stroka++;
                adr = "A" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).Font.Bold = false;
                adr = "B" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).HorizontalAlignment = XlHAlign.xlHAlignRight;
                WS.Cells[stroka, 1] = "Референт государственной гражданской службы  Российской Федерации 1 - 3 класса";
                WS.Cells[stroka, 1].WrapText = true;
                WS.Cells[stroka, 3] = SUM_REFERENT;
                WS.Cells[stroka, 2] = SUM_REFERENT_FAKT;
              //  WS.Cells[stroka, 3] = "";
                WS.Cells[stroka, 4] = "";
                WS.Cells[stroka, 5] = "";
                WS.Cells[stroka, 6] = "";
                WS.Cells[stroka, 7] = "";
                stroka++;
                adr = "A" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).Font.Bold = false;
                adr = "B" + stroka + ":" + "C" + stroka;
                WS.get_Range(adr).HorizontalAlignment = XlHAlign.xlHAlignRight;
                WS.Cells[stroka, 1] = "Секретарь государственной гражданской службы  Российской Федерации 1 - 3 класса";
                WS.Cells[stroka, 1].WrapText = true;
                WS.Cells[stroka, 2] = SUM_SEKRETAR_FAKT;
                WS.Cells[stroka, 3] = SUM_SEKRETAR;
               // WS.Cells[stroka, 3] = "";
                WS.Cells[stroka, 4] = "";
                WS.Cells[stroka, 5] = "";
                WS.Cells[stroka, 6] = "";
                WS.Cells[stroka, 7] = "";



                stroka++;
                adr = "A" + stroka + ":" + "E" + stroka;
                WS.get_Range(adr).Font.Name = "Times New Roman";
                WS.get_Range(adr).Font.Size = 12;
                WS.get_Range(adr).Font.Bold = true;

                WS.Cells[stroka, 1] = "Сотрудники";
               // WS.Cells[stroka, 4] = 0;
                WS.Cells[stroka, 4] = SUM_Vysh_Nach_sostav+SUM_Starsh_Nach_sostav+SUM_Srednii_Nach_sostav+SUM_Mladshii_sostav;
                WS.Cells[stroka, 5] = SUM_Vysh_Nach_sostav_FAKT + SUM_Starsh_Nach_sostav_FAKT + SUM_Srednii_Nach_sostav_FAKT + SUM_Mladshii_sostav_FAKT;

                stroka++;
                WS.Cells[stroka, 1] = "в том числе: ";
                stroka++;
                WS.Cells[stroka, 1] = "высший начальствующий состав";
                WS.Cells[stroka, 4] = SUM_Vysh_Nach_sostav;
                WS.Cells[stroka, 5] = SUM_Vysh_Nach_sostav_FAKT;
                stroka++;
                WS.Cells[stroka, 1] = "старший начальствующий состав";
                WS.Cells[stroka, 4] = SUM_Starsh_Nach_sostav;
                WS.Cells[stroka, 5] = SUM_Starsh_Nach_sostav_FAKT;
                stroka++;
                WS.Cells[stroka, 1] = "средний начальствующий состав";
                WS.Cells[stroka, 4] = SUM_Srednii_Nach_sostav;
                WS.Cells[stroka, 5] = SUM_Srednii_Nach_sostav_FAKT;
                stroka++;
                WS.Cells[stroka, 1] = "младший состав";
                WS.Cells[stroka, 4] = SUM_Mladshii_sostav;
                WS.Cells[stroka, 5] = SUM_Mladshii_sostav_FAKT;




                stroka++;
                adr = "A" + stroka + ":" + "G" + stroka;
                WS.get_Range(adr).Font.Name = "Times New Roman";
                WS.get_Range(adr).Font.Size = 12;
                WS.get_Range(adr).Font.Bold = true;

                WS.Cells[stroka, 1] = "Работники";
                // WS.Cells[stroka, 4] = 0;
                WS.Cells[stroka, 6] = SUM_Rabochie+SUM_Rykovod+SUM_Spec_Slyj;
                WS.Cells[stroka, 7] = SUM_Rabochie_FAKT + SUM_Rykovod_FAKT + SUM_Spec_Slyj_FAKT;

                stroka++;
                WS.Cells[stroka, 1] = "в том числе: ";
                stroka++;
                WS.Cells[stroka, 1] = "руководители";
                WS.Cells[stroka, 7] = SUM_Rykovod_FAKT;
                WS.Cells[stroka, 6] = SUM_Rykovod;
                stroka++;
                WS.Cells[stroka, 1] = "специалисты и служащие";
                WS.Cells[stroka, 7] = SUM_Spec_Slyj_FAKT;
                WS.Cells[stroka, 6] = SUM_Spec_Slyj;
                stroka++;
                WS.Cells[stroka, 1] = "рабочие";
                WS.Cells[stroka, 7] = SUM_Rabochie_FAKT;
                WS.Cells[stroka, 6] = SUM_Rabochie;
                


                #region
                //заполнение рамок
                string adrRamok = "A5:" + "G" + stroka;

                Excelcells = WS.get_Range(adrRamok);
                Microsoft.Office.Interop.Excel.XlBordersIndex BorderIndex;

                BorderIndex = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft;
                Excelcells.Borders[BorderIndex].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Excelcells.Borders[BorderIndex].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Excelcells.Borders[BorderIndex].ColorIndex = 0;


                BorderIndex = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop;
                Excelcells.Borders[BorderIndex].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Excelcells.Borders[BorderIndex].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Excelcells.Borders[BorderIndex].ColorIndex = 0;


                BorderIndex = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom;
                Excelcells.Borders[BorderIndex].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Excelcells.Borders[BorderIndex].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Excelcells.Borders[BorderIndex].ColorIndex = 0;

                BorderIndex = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight;
                Excelcells.Borders[BorderIndex].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Excelcells.Borders[BorderIndex].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Excelcells.Borders[BorderIndex].ColorIndex = 0;

                BorderIndex = Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical;
                Excelcells.Borders[BorderIndex].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Excelcells.Borders[BorderIndex].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Excelcells.Borders[BorderIndex].ColorIndex = 0;

                BorderIndex = Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal;
                Excelcells.Borders[BorderIndex].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Excelcells.Borders[BorderIndex].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Excelcells.Borders[BorderIndex].ColorIndex = 0;
                WS.Range["A1"].Value = arrTM_TO;
                WS.Range["B1"].Value = String.Format("{0}.{1}.{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year); ;
                #endregion
                var rr= XZ;
                WB.SaveAs(pathSaveSpisokLicz, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    WB.Close(true);
                    //File.Copy(@"D:\Справки\Список_лиц_командированных_в_МОТ.xlsx", AppDomain.CurrentDomain.BaseDirectory + @"Templates\Список_лиц_командированных_в_МОТ.xlsx", true);
                    ObjWorkExcel.Quit();

                FULL_SUM_GOS_SHTAT = 0 ;
                FULL_SUM_SOTR_SHTAT = 0 ;
                FULL_SUM_RAB_SHTAT =0;
                FULL_SUM_GOS_FAKT=0;
                FULL_SUM_SOTR_FAKT=0;
                FULL_SUM_RAB_FAKT=0;


            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }

            return 1;
        }

        static private void QuOtdel(int RN, OracleConnection conn)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                string sqlOtd = "SELECT id.DEPART_DISP,id.PRN,id.NAME, id.NAME_NOM, id.RN, id.CODE, ie.urlev FROM UK_PARUS.INS_DEPARTMENT id INNER JOIN (SELECT id3.RN, LEVEL AS urlev FROM UK_PARUS.INS_DEPARTMENT id3 START WITH id3.RN =" + RN + "CONNECT BY PRIOR id3.RN = id3.PRN) ie   ON id.RN = ie.RN WHERE id.BGNDATE <= TRUNC(SYSDATE)  AND (id.ENDDATE >= TRUNC(SYSDATE)  OR id.ENDDATE IS NULL)  AND ie.urlev <> 1 AND id.ORG_SIGN = 0 ORDER BY id.NAME_NOM";

                // Сочетать Command с Connection.
                cmd.Connection = conn;
                conn.Open();
                cmd.CommandText = sqlOtd;
                massiv._Otdel.Clear();
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        massiv._Otdel.Add(new Otdel()
                        {
                            DEPART_DISP = (reader["DEPART_DISP"].ToString()),
                            PRN = int.Parse((reader["PRN"].ToString())),
                            NAME = (reader["NAME"].ToString()),
                            NAME_NOM = (reader["NAME_NOM"].ToString()),
                            RN = int.Parse((reader["RN"].ToString())),
                            CODE = (reader["CODE"].ToString()),
                            urlev = int.Parse((reader["urlev"].ToString())),

                        });

                        
                    }
                    conn.Close();

                }
              
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }
        }

        static private void QuDolj (OracleConnection conn, string DEPART, int arrTM_depart)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                string sqlOtd = "SELECT MAIN.POST , MAIN.UPRAV , MAIN.KOLSHTAT , MAIN.PIPEC , MAIN.EXEC, MAIN.CLNKOD, MAIN.SLVID, MAIN.RN,  MAIN.DEP_CODE, MAIN.AGN_RN,  MAIN.CLNZVAN, MAIN.CRN FROM(SELECT dep.RN AS RN, FR_ADD_COL_PERSONS(dep.RN) AS DEP_CODE, dm_query.get_svyz(dep.RN) AS AGN_RN, sts.NAME_NOM AS EXEC, dep.PSDEP_NAME AS CLNCODE, dm_query.get_clnpsdep_zvan(dep.RN) AS CLNZVAN, dep.DO_ACT_FROM AS CLNDOLDAT, dep.DEPART_DISP AS CLNKOD, dm_query.get_clnpsdep_kad(dep.RN) AS CLNKAD, dm_query.get_clnpsdep_group(dep.RN) AS CLNGROUP, dm_query.get_departdisp_post(dep.RN) AS POST, dm_query.get_depart2(dep.RN) AS UPRAV, dep.RATEACC AS KOLSHTAT, DECODE(dep.ON_STAFF, 0, 'внештатная', 1, 'штатная') AS DOLVID, dm_query.get_block(dep.RN) AS BLOCK,dm_query.get_pipec(dep.RN) AS PIPEC, dm_query.get_pipec2(dep.RN) AS PIPEC2, dm_query.get_nomnaz(dep.RN) AS NOMNAZ, dm_query.get_slvid(dep.RN) AS SLVID, dm_query.get_specdol(dep.RN) AS SPECDOL, dm_query.get_sovdol(dep.RN) AS SOVDOL, dm_query.get_depart_uprav_ctu(dep.RN) AS UPRAVL, dep.CRN AS CRN, sts.CRN AS EXPR1, p.PRN, sts.RN AS EXPR2, dep.POSTRN, sts.POST_DISP,  sts.IS_ROT FROM UK_PARUS.CLNPSDEP dep INNER JOIN UK_PARUS.CLNPOSTS sts ON dep.POSTRN = sts.RN INNER JOIN UK_PARUS.CLNPSDEPPRS p ON dep.RN = p.PRN WHERE dep.DO_ACT_FROM <= SYSDATE AND(dep.DO_ACT_TO >= SYSDATE OR dep.DO_ACT_TO IS NULL)) MAIN WHERE EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP WHERE UP.CATALOG = MAIN.CRN) AND MAIN.POST =" + "'" + DEPART + "'" + " AND MAIN.CRN = " + arrTM_depart;

                // Сочетать Command с Connection.
                cmd.Connection = conn;
                conn.Open();
                cmd.CommandText = sqlOtd;
                massiv._Dolj.Clear();
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        massiv._Dolj.Add(new Dolj()
                        {

                            MAIN_POST = (reader["POST"].ToString()),
                            MAIN_UPRAV = (reader["UPRAV"].ToString()),
                            MAIN_KOLSHTAT = int.Parse((reader["KOLSHTAT"].ToString())),
                            MAIN_PIPEC = (reader["PIPEC"].ToString()),
                            EXEC = (reader["EXEC"].ToString()),
                            CLNKOD = int.Parse((reader["CLNKOD"].ToString())),
                            SLVID = (reader["SLVID"].ToString()),
                            RN = int.Parse((reader["RN"].ToString())),
                            DEP_CODE = (reader["DEP_CODE"].ToString()),
                            AGN_RN = ((reader["AGN_RN"].ToString())),
                            CLNZVAN = (reader["CLNZVAN"].ToString()),
                            CRN = int.Parse((reader["CRN"].ToString())),


                        });
                    }
                    conn.Close();

                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }
        }


        static private void QuSluDolj(OracleConnection conn, string DEPART, int arrTM_depart)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                string sqlOtd = "SELECT MAIN.POST,  MAIN.UPRAV, MAIN.KOLSHTAT, MAIN.PIPEC,MAIN.EXEC,  MAIN.CLNKOD, MAIN.SLVID, MAIN.RN, MAIN.DEP_CODE, MAIN.AGN_RN,  MAIN.CLNZVAN,MAIN.CRN, MAIN.CLNCODE,  MAIN.CLNGROUP FROM(SELECT dep.RN AS RN, FR_ADD_COL_PERSONS(dep.RN) AS DEP_CODE, dm_query.get_svyz(dep.RN) AS AGN_RN, sts.NAME_NOM AS EXEC, dep.PSDEP_NAME AS CLNCODE, dm_query.get_clnpsdep_zvan(dep.RN) AS CLNZVAN, dep.DO_ACT_FROM AS CLNDOLDAT, dep.DEPART_DISP AS CLNKOD, dm_query.get_clnpsdep_kad(dep.RN) AS CLNKAD, dm_query.get_clnpsdep_group(dep.RN) AS CLNGROUP, dm_query.get_departdisp_post(dep.RN) AS POST, dm_query.get_depart2(dep.RN) AS UPRAV, dep.RATEACC AS KOLSHTAT, DECODE(dep.ON_STAFF, 0, 'внештатная', 1, 'штатная') AS DOLVID, dm_query.get_block(dep.RN) AS BLOCK, dm_query.get_pipec(dep.RN) AS PIPEC, dm_query.get_pipec2(dep.RN) AS PIPEC2, dm_query.get_nomnaz(dep.RN) AS NOMNAZ, dm_query.get_slvid(dep.RN) AS SLVID, dm_query.get_specdol(dep.RN) AS SPECDOL, dm_query.get_sovdol(dep.RN) AS SOVDOL, dm_query.get_depart_uprav_ctu(dep.RN) AS UPRAVL, dep.CRN AS CRN,sts.CRN AS EXPR1,p.PRN,sts.RN AS EXPR2,dep.POSTRN,sts.POST_DISP,sts.IS_ROT,sts.CODE AS EXPR3,dep.DEPTRN, p.PREDPROF1,p.DUMMY,p.PRRANKTYPE3,p.IS_REPLRNK FROM UK_PARUS.CLNPSDEP dep        INNER JOIN UK_PARUS.CLNPOSTS sts          ON dep.POSTRN = sts.RN        INNER JOIN UK_PARUS.CLNPSDEPPRS p          ON dep.RN = p.PRN      WHERE dep.DO_ACT_FROM <= SYSDATE    AND(dep.DO_ACT_TO >= SYSDATE        OR dep.DO_ACT_TO IS NULL)) MAIN WHERE EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY      FROM UK_PARUS.V_USERPRIV UP      WHERE UP.CATALOG = MAIN.CRN) AND MAIN.POST =" + "'" + DEPART + "'" + "  AND MAIN.CRN = " + arrTM_depart;
                // Сочетать Command с Connection.
                cmd.Connection = conn;
                conn.Open();
                cmd.CommandText = sqlOtd;
                massiv._SlujDolj.Clear();
                
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        var clnkod = 0;
                        var cl = reader["CLNKOD"].ToString();
                        if (cl == "")
                        {
                           clnkod = 0;
                        }
                        else
                        {
                            clnkod = int.Parse(reader["CLNKOD"].ToString());
                        }
                        massiv._SlujDolj.Add(new SlujDolj()
                        {

                            MAIN_POST = (reader["POST"].ToString()),
                            MAIN_UPRAV = (reader["UPRAV"].ToString()),
                            MAIN_KOLSHTAT = double.Parse((reader["KOLSHTAT"].ToString())),
                            MAIN_PIPEC = (reader["PIPEC"].ToString()),
                            EXEC = (reader["EXEC"].ToString()),

                           // CLNKOD = int.Parse((reader["CLNKOD"].ToString())),

                            CLNKOD = clnkod,
                            SLVID = (reader["SLVID"].ToString()),
                            RN = int.Parse((reader["RN"].ToString())),
                            DEP_CODE = (reader["DEP_CODE"].ToString()),
                            AGN_RN = ((reader["AGN_RN"].ToString())),
                            CLNZVAN = (reader["CLNZVAN"].ToString()),
                            CRN = int.Parse((reader["CRN"].ToString())),
                            CLNCODE = (reader["CLNCODE"].ToString()),
                            CLNGROUP = (reader["CLNGROUP"].ToString()),
                        });
                    }
                    conn.Close();

                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }
        }

        static private void QuSvedenia (OracleConnection conn, int arrTM_depart)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                string sqlSved = "SELECT MAIN_AGNRN.CLNGROUP AS SOSTAV,  MAIN_AGNRN.FULLNAME AS FIO, MAIN.AGN_RN AS Numb, MAIN.CRN, MAIN.DEP_CODE, MAIN.CLNZVAN,  MAIN_AGNRN.CLNKAD, MAIN.SLVID FROM (SELECT dep.RN AS RN, FR_ADD_COL_PERSONS(dep.RN) AS DEP_CODE, dm_query.get_svyz(dep.RN) AS AGN_RN, dm_query.get_clnpsdep_group(dep.RN) AS CLNGROUP,  dm_query.get_clnpsdep_zvan(dep.RN) AS CLNZVAN, dm_query.get_slvid(dep.RN) AS SLVID, dep.CRN AS CRN FROM UK_PARUS.CLNPSDEP dep INNER JOIN UK_PARUS.CLNPOSTS sts ON dep.POSTRN = sts.RN INNER JOIN UK_PARUS.CLNPSDEPPRS p ON dep.RN = p.PRN  WHERE dep.DO_ACT_FROM <= SYSDATE  AND (dep.DO_ACT_TO >= SYSDATE  OR dep.DO_ACT_TO IS NULL)) MAIN  LEFT OUTER JOIN (SELECT a.AGNFAMILYNAME AS agnfamilyname,  a.AGNFIRSTNAME AS agnfirstname, a.AGNLASTNAME AS agnlastname,  a.AGNFAMILYNAME || ' ' || a.AGNFIRSTNAME || ' ' || a.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(a.RN) AS INICIALI, a.RN AS RN,   dm_query.get_clnpsdep_kod(a.RN) AS CLNKOD,   dm_query.get_clnpsdep_kad(a.RN) AS CLNKAD,     dm_query.get_clnpsdep_group(a.RN) AS CLNGROUP      FROM UK_PARUS.AGNLIST a  WHERE a.AGNTYPE = " + "'1'" + " AND a.RN IN (SELECT pp.AGNLIST  FROM UK_PARUS.PREMPLFLS pp)) MAIN_AGNRN  ON MAIN_AGNRN.RN = MAIN.AGN_RN WHERE EXISTS (SELECT UP.RN, UP.AUTHID,  UP.ROLEID, UP.COMPANY, UP.VERSION,  UP.UNITCODE, UP.CATALOG, UP.JUR_PERS,   UP.HIERARCHY  FROM UK_PARUS.V_USERPRIV UP  WHERE UP.CATALOG = MAIN.CRN)  AND MAIN.CRN =" + arrTM_depart ;
                // Сочетать Command с Connection.
                cmd.Connection = conn;
                conn.Open();
                cmd.CommandText = sqlSved;
                massiv._Svedenia.Clear();
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        massiv._Svedenia.Add(new Svedenia()
                        {

                            SOSTAV = (reader["SOSTAV"].ToString()),
                            FIO = (reader["FIO"].ToString()),
                            Numb = (reader["Numb"].ToString()),
                            CLNKAD = ((reader["CLNKAD"].ToString())),
                            DEP_CODE = (reader["DEP_CODE"].ToString()),
                            CLNZVAN = (reader["CLNZVAN"].ToString()),
                            CRN = (reader["CRN"].ToString()),
                            SLVID = (reader["SLVID"].ToString()),
                        });
                    }
                    conn.Close();

                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }

        }

        static private void QuSvedRab(OracleConnection conn, int  arrTM_depart)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                string sqlOtd = "SELECT MAIN.POST,  MAIN.UPRAV, MAIN.KOLSHTAT, MAIN.PIPEC,MAIN.EXEC,  MAIN.CLNKOD, MAIN.SLVID, MAIN.RN, MAIN.DEP_CODE, MAIN.AGN_RN,  MAIN.CLNZVAN,MAIN.CRN, MAIN.CLNCODE,  MAIN.CLNGROUP FROM(SELECT dep.RN AS RN, FR_ADD_COL_PERSONS(dep.RN) AS DEP_CODE, dm_query.get_svyz(dep.RN) AS AGN_RN, sts.NAME_NOM AS EXEC, dep.PSDEP_NAME AS CLNCODE, dm_query.get_clnpsdep_zvan(dep.RN) AS CLNZVAN, dep.DO_ACT_FROM AS CLNDOLDAT, dep.DEPART_DISP AS CLNKOD, dm_query.get_clnpsdep_kad(dep.RN) AS CLNKAD, dm_query.get_clnpsdep_group(dep.RN) AS CLNGROUP, dm_query.get_departdisp_post(dep.RN) AS POST, dm_query.get_depart2(dep.RN) AS UPRAV, dep.RATEACC AS KOLSHTAT, DECODE(dep.ON_STAFF, 0, 'внештатная', 1, 'штатная') AS DOLVID, dm_query.get_block(dep.RN) AS BLOCK, dm_query.get_pipec(dep.RN) AS PIPEC, dm_query.get_pipec2(dep.RN) AS PIPEC2, dm_query.get_nomnaz(dep.RN) AS NOMNAZ, dm_query.get_slvid(dep.RN) AS SLVID, dm_query.get_specdol(dep.RN) AS SPECDOL, dm_query.get_sovdol(dep.RN) AS SOVDOL, dm_query.get_depart_uprav_ctu(dep.RN) AS UPRAVL, dep.CRN AS CRN,sts.CRN AS EXPR1,p.PRN,sts.RN AS EXPR2,dep.POSTRN,sts.POST_DISP,sts.IS_ROT,sts.CODE AS EXPR3,dep.DEPTRN, p.PREDPROF1,p.DUMMY,p.PRRANKTYPE3,p.IS_REPLRNK FROM UK_PARUS.CLNPSDEP dep        INNER JOIN UK_PARUS.CLNPOSTS sts          ON dep.POSTRN = sts.RN        INNER JOIN UK_PARUS.CLNPSDEPPRS p          ON dep.RN = p.PRN      WHERE dep.DO_ACT_FROM <= SYSDATE    AND(dep.DO_ACT_TO >= SYSDATE        OR dep.DO_ACT_TO IS NULL)) MAIN WHERE EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY      FROM UK_PARUS.V_USERPRIV UP      WHERE UP.CATALOG = MAIN.CRN)  AND MAIN.CRN = " + arrTM_depart;
                // Сочетать Command с Connection.
                cmd.Connection = conn;
                conn.Open();
                cmd.CommandText = sqlOtd;
                massiv._SotrudnikiNizBlok.Clear();
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    
                    while (reader.Read())
                    {
                        var clnkod = 0;
                        var cl = reader["CLNKOD"].ToString();
                        if (cl == "")
                        {
                            clnkod = 0;
                        }
                        else
                        {
                            clnkod = int.Parse(reader["CLNKOD"].ToString());
                        }
                        massiv._SotrudnikiNizBlok.Add(new SotrudnikiNizBlok()
                        {

                            MAIN_POST = (reader["POST"].ToString()),
                            MAIN_UPRAV = (reader["UPRAV"].ToString()),
                            MAIN_KOLSHTAT = double.Parse((reader["KOLSHTAT"].ToString())),
                            MAIN_PIPEC = (reader["PIPEC"].ToString()),
                            EXEC = (reader["EXEC"].ToString()),

                           // CLNKOD = int.Parse((reader["CLNKOD"].ToString())),

                            CLNKOD = clnkod,
                            SLVID = (reader["SLVID"].ToString()),
                            RN = int.Parse((reader["RN"].ToString())),
                            DEP_CODE = (reader["DEP_CODE"].ToString()),
                            AGN_RN = ((reader["AGN_RN"].ToString())),
                            CLNZVAN = (reader["CLNZVAN"].ToString()),
                            CRN = int.Parse((reader["CRN"].ToString())),
                            CLNCODE = (reader["CLNCODE"].ToString()),
                            CLNGROUP = (reader["CLNGROUP"].ToString()),
                        });
                    }
                    conn.Close();

                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }
        }


        static private void QuSvedRabotniki(OracleConnection conn, int arrTM_RN)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                string sqlOtd = "SELECT SHD.CODE AS Dolj, TKN.CODE AS VidSljub,SHD.NUMB, OFC.NAME AS CatDolj from CLNPSDEP shd, CLNPSDEPPRS shds,   CLNPSDEPHS shdh, PRDTKN tkn,  OFFICERCLS ofc  where shd.rn = shds.prn and shd.deptrn in  (select tt.rn from(select it.rn, it.org_sign from INS_DEPARTMENT it start with it.RN = " + arrTM_RN + "  connect by prior it.RN = it.PRN) tt  ) and shd.officercls = ofc.rn   and shds.prdtkn = tkn.rn  and upper(tkn.code) = upper('Работник')   and shd.on_staff = 1 and shdh.prn = shd.rn and shd.do_act_from <= TRUNC(SYSDATE) and((shd.do_act_to >= TRUNC(SYSDATE)) or(shd.do_act_to is null)) and shdh.do_act_from <= TRUNC(SYSDATE)    and((shdh.do_act_to >= TRUNC(SYSDATE)) or(shdh.do_act_to is null))";// Сочетать Command с Connection.
                cmd.Connection = conn;
                conn.Open();
                cmd.CommandText = sqlOtd;
                massiv._RabotnikiBlok.Clear();
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        massiv._RabotnikiBlok.Add(new RabotnikiBlok()
                        {

                            CatDolj = (reader["CatDolj"].ToString()),
                            Dolj = (reader["Dolj"].ToString()),
                            NUMB = (reader["NUMB"].ToString()),
                            VidSljub = (reader["VidSljub"].ToString()),

                        });
                    }
                    conn.Close();

                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }
        }

        static private void QuSvedRabotnikiFakt(OracleConnection conn, int arrTM_depart)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                string sqlOtd = "SELECT  MAIN.UPRAV, MAIN.KOLSHTAT, MAIN.PIPEC,  MAIN.EXEC,  MAIN.SLVID, MAIN.DEP_CODE,  MAIN.AGN_RN,       MAIN.NUMB FROM(SELECT dep.RN AS RN, FR_ADD_COL_PERSONS(dep.RN) AS DEP_CODE, dm_query.get_svyz(dep.RN) AS AGN_RN, sts.NAME_NOM AS EXEC, dep.PSDEP_NAME AS CLNCODE, dm_query.get_clnpsdep_zvan(dep.RN) AS CLNZVAN, dep.DO_ACT_FROM AS CLNDOLDAT, dep.DEPART_DISP AS CLNKOD, dm_query.get_clnpsdep_kad(dep.RN) AS CLNKAD, dm_query.get_clnpsdep_group(dep.RN) AS CLNGROUP, dm_query.get_departdisp_post(dep.RN) AS POST, dm_query.get_depart2(dep.RN) AS UPRAV, dep.RATEACC AS KOLSHTAT, DECODE(dep.ON_STAFF, 0, 'внештатная', 1, 'штатная') AS DOLVID, dm_query.get_block(dep.RN) AS BLOCK, dm_query.get_pipec(dep.RN) AS PIPEC, dm_query.get_pipec2(dep.RN) AS PIPEC2, dm_query.get_nomnaz(dep.RN) AS NOMNAZ, dm_query.get_slvid(dep.RN) AS SLVID, dm_query.get_specdol(dep.RN) AS SPECDOL, dm_query.get_sovdol(dep.RN) AS SOVDOL, dm_query.get_depart_uprav_ctu(dep.RN) AS UPRAVL, dep.CRN AS CRN,sts.CRN AS EXPR1, p.PRN,sts.RN AS EXPR2, dep.POSTRN, sts.POST_DISP,sts.IS_ROT, sts.CODE AS EXPR3,dep.DEPTRN, p.PREDPROF1,p.DUMMY, p.PRRANKTYPE3,p.IS_REPLRNK, dep.NUMB FROM UK_PARUS.CLNPSDEP dep INNER JOIN UK_PARUS.CLNPOSTS sts ON dep.POSTRN = sts.RN INNER JOIN UK_PARUS.CLNPSDEPPRS p ON dep.RN = p.PRN WHERE dep.DO_ACT_FROM <= SYSDATE AND(dep.DO_ACT_TO >= SYSDATE OR dep.DO_ACT_TO IS NULL)) MAIN WHERE EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP WHERE UP.CATALOG = MAIN.CRN) AND MAIN.CRN =" + arrTM_depart + " AND MAIN.SLVID = 'Работник'";
                cmd.Connection = conn;
                conn.Open();
                cmd.CommandText = sqlOtd;
                massiv._RabotnikiBlokFact.Clear();
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        massiv._RabotnikiBlokFact.Add(new RabotnikiBlokFact()
                        {

                            AGN_RN = (reader["AGN_RN"].ToString()),
                            DEP_CODE = (reader["DEP_CODE"].ToString()),
                            EXEC = (reader["EXEC"].ToString()),
                            KOLSHTAT = (reader["KOLSHTAT"].ToString()),
                            NUMB = (reader["NUMB"].ToString()),
                            PIPEC = (reader["PIPEC"].ToString()),
                            SLVID = (reader["SLVID"].ToString()),
                            UPRAV = (reader["UPRAV"].ToString()),

                        });
                    }
                    conn.Close();

                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var n = comboBox1.SelectedIndex.ToString();
            arrTM_RN = Convert.ToInt32(massiv.TM[0, Convert.ToInt32(n)]);
            arrTM_TO = massiv.TM[1, Convert.ToInt32(n)];
            arrTM_depart = Convert.ToInt32(massiv.TM[2, Convert.ToInt32(n)]);
            button2.Enabled = true;
          
        }
    }
}
