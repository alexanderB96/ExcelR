using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelR
{
    class _16KADVigruzka
    {
        Form1 fr = new Form1();
        DBOracleUtils db = new DBOracleUtils();
        OracleCommand cmd = new OracleCommand();
        OracleCommand cmdN = new OracleCommand();
        OracleConnection conn = new OracleConnection();
        OracleDataReader dr;
        Excel.Application ObjWorkExcel = new Excel.Application(); //сам эксель

        private int[] RN ={ 222318,221665,221668,221671,221674,221683,221689,221692,451367509,221737,221719,221725,221728,221652,221734};

/*
        #region
        string pprim3, pprim6, pprim9, pprim12, pprim15, pprim18, pprim21, pprim24, pprim27, pprim30, pprim33, pprim36, pprim39, pprim42, pprim45, pprim48, pprim51;
        string gprim3,gprim6,gprim9,gprim12,gprim15,gprim18,gprim21,gprim24,gprim27,gprim30,gprim33,gprim36,gprim39,gprim42,gprim45,gprim48,gprim51;
        string pprim54,pprim57,pprim60,pprim63,pprim66,pprim69,pprim72,pprim75,pprim78,pprim81,pprim84;
        string gprim54,gprim57,gprim60,gprim63,gprim66,gprim69,gprim72,gprim75,gprim78,gprim81,gprim84;
        #endregion



        Excel.Workbook ObjWorkBooks;// КНИГА
        Excel.Worksheet ObjWorkSheets;//ЛИСТ*/
        public async void F16KadAsync(string dat)
        {
            await Task.Run(() => Formirovanie16Kad(dat));
        }

        public void Formirovanie16Kad(string dat)
        {
            fr.Refresh();
            int ppcol, pprow;
            try
            {
                foreach (int RN_INS in RN)
                {
                    // int RN_INS = 222318;
                    conn = db.GetDBConnection();
                    conn.Open();
                    cmd.Connection = conn;


                    cmd.CommandText = "select trim(code) code, TRIM(name_nom) name_nom, trim(depart_code) depart_code, name_gen, shortname_nom  from ins_department  where RN =" + RN_INS;


                    dr = cmd.ExecuteReader();
                    dr.Read();
                    var ins_code = dr["code"].ToString();
                    var ins_gen = dr["name_gen"].ToString();
                    var ins_name = dr["name_gen"].ToString();
                    var depart_code = dr["depart_code"].ToString();



                    FileStream filestream = new FileStream(@"D:\Справки\KAD_16\test1.xls", FileMode.OpenOrCreate);
                    cmd.CommandText = "select t.data from tatable t where t.file_name = 'KAD16_2013.XLS' ";
                    using (dr = cmd.ExecuteReader())
                    {
                        dr.Read();
                        byte[] b = new byte[(dr.GetBytes(0, 0, null, 0, int.MaxValue))];
                        dr.GetBytes(0, 0, b, 0, b.Length);
                        filestream.Write(b, 0, b.Length);
                    }

                    Workbook WB = ObjWorkExcel.Workbooks.Add(@"D:\Справки\KAD_16\test1.xls");//создаем новую книгу
                                                                                             //ObjWorkBooks = ObjWorkExcel.Workbooks.Open(Convert.ToString(FileName)); //открываем существующую книгу
                    Worksheet xlsSheet;
                    ObjWorkExcel.DisplayAlerts = false;
                    xlsSheet = (Worksheet)WB.Sheets[1]; // раздел 1 (лист 1)
                    xlsSheet.Activate();
                    xlsSheet.Cells[15, 7] = ins_name;
                    xlsSheet.Cells[15, 19] = ins_code;

                    xlsSheet = (Worksheet)WB.Sheets[2]; // раздел 1 (лист 1)
                    xlsSheet.Activate();
                    xlsSheet.Cells[2, 1] = "Штатная численность сотрудников, государственных гражданских служащих и работников таможенных органов " + ins_gen;
                    xlsSheet.Cells[3, 1] = "по состоянию на " + (Convert.ToDateTime(dat)).ToString("dd/MM/yyyy");

                    cmdN.Connection = conn;
                    cmdN.CommandText = "PR_KAD16_2013";
                    cmdN.CommandType = System.Data.CommandType.StoredProcedure;
                    cmdN.Parameters.Add("nDEPT", OracleDbType.Int32).Value = 0;
                    cmdN.Parameters.Add("dDATE", OracleDbType.Date).Value = Convert.ToDateTime(dat);
                    cmdN.ExecuteNonQuery();
                    cmdN.Cancel();
                    conn.Close();
                    cmdN.Parameters.Clear();
                    conn.Open();

                    cmd.CommandText = "PR_KAD16_2013";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.Add("nDEPT", OracleDbType.Int32).Value = RN_INS;
                    cmd.Parameters.Add("dDATE", OracleDbType.Date).Value = Convert.ToDateTime(dat);
                    cmd.ExecuteNonQuery();
                    cmd.Prepare();
                    cmd.Cancel();
                    cmd.Parameters.Clear();
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = "select t.column_numb, t.row_numb, t.quant, t.fam from CUR_KAD16 t where t.authid = 'CTU_NESTERKINA'";
                    dr = cmd.ExecuteReader();
                    //dr.Read();
                    while (dr.Read())
                    {
                        ppcol = Convert.ToInt32(dr["column_numb"].ToString()) - 1;
                        pprow = Convert.ToInt32(dr["row_numb"].ToString());
                        xlsSheet.Cells[pprow, ppcol] = dr["quant"].ToString();
                        if (dr["fam"].ToString() != "")
                        {
                            var tt = dr["fam"].ToString();
                            xlsSheet.Cells[pprow, ppcol].Cells.AddComment(dr["fam"].ToString());
                          

                        }
                        #region
                        /*
                        if(ppcol == 3)
                        {
                            if ((pprow >=10) & (pprow<= 36))
                                {
                                pprim3 += pprim3 + dr["fam"].ToString();
                                }
                            if ((pprow >= 39) & (pprow <= 64))
                            {
                                gprim3 += gprim3 + dr["fam"].ToString();
                            }
                        }

                        if (ppcol == 6)
                        {
                            if ((pprow >= 10) & (pprow <= 36))
                            {
                                pprim6 += pprim6 + dr["fam"].ToString();
                            }
                            if ((pprow >= 39) & (pprow <= 64))
                            {
                                gprim6 += gprim6 + dr["fam"].ToString();
                            }
                        }

                        if (ppcol == 9)
                        {
                            if ((pprow >= 10) & (pprow <= 36))
                            {
                                pprim9 += pprim9 + dr["fam"].ToString();
                            }
                            if ((pprow >= 38) & (pprow <= 64))
                            {
                                gprim9 += gprim9 + dr["fam"].ToString();
                            }
                        }
                        if (ppcol == 12)
                        {
                            if ((pprow >= 10) & (pprow <= 36))
                            {
                                pprim12 += pprim12 + dr["fam"].ToString();
                            }
                            if ((pprow >= 39) & (pprow <= 64))
                            {
                                gprim12 += gprim12 + dr["fam"].ToString();
                            }
                        }
                        if (ppcol == 15)
                        {
                            if ((pprow >= 10) & (pprow <= 36))
                            {
                                pprim15 += pprim15 + dr["fam"].ToString();
                            }
                            if ((pprow >= 39) & (pprow <= 64))
                            {
                                gprim15 += gprim15 + dr["fam"].ToString();
                            }
                        }
                        if (ppcol == 18)
                        {
                            if ((pprow >= 10) & (pprow <= 36))
                            {
                                pprim18 += pprim18 + dr["fam"].ToString();
                            }
                            if ((pprow >= 39) & (pprow <= 64))
                            {
                                gprim18 += gprim18 + dr["fam"].ToString();
                            }
                        }
                        if (ppcol == 21)
                        {
                            if ((pprow >= 10) & (pprow <= 36))
                            {
                                pprim21 += pprim21 + dr["fam"].ToString();
                            }
                            if ((pprow >= 39) & (pprow <= 64))
                            {
                                gprim21 += gprim21 + dr["fam"].ToString();
                            }
                        }
                        if (ppcol == 24)
                        {
                            if ((pprow >= 10) & (pprow <= 36))
                            {
                                pprim24 += pprim24 + dr["fam"].ToString();
                            }
                            if ((pprow >= 39) & (pprow <= 64))
                            {
                                gprim24 += gprim24+ dr["fam"].ToString();
                            }
                        }
                        */
                        #endregion
                    }
                    string pathSave16KAD = String.Format(@"D:\Справки\KAD16_{0}_{1}-{2}-{3}", depart_code, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year);
                    WB.SaveAs(pathSave16KAD);
                    filestream.Close();
                    WB.Close(false, Type.Missing, Type.Missing);
                }



                fr.CloseLoad();
                fr.Kad16Vigruzka.Enabled = true;
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message,
                     "Error",
                     MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            finally
            {
                MessageBox.Show("Успешно!");
                conn.Close();
                fr.loadingNew2.Hide();
                
                fr.CloseLoad();
                foreach (Process currentProcess in Process.GetProcessesByName("EXCEL"))
                {
                    currentProcess.Kill();
                    
                }
                fr.Refresh();
            }
        }

    }
}
