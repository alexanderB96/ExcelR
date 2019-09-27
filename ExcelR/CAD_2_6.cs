using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace ExcelR
{
    public partial class CAD_2_6 : Form
    {
        public CAD_2_6()
        {
            InitializeComponent();
        }


        static MassivDannyh massiv = new MassivDannyh();
        static public int arrTM_depart = -5;
        static public string arrTM_TO;
        static public int arrTM_RN;
        static public int arrTM_mnemo;
        static public int items_period = -1;
        static public string Local_ADR_TEST;
        int Kol_vo=0;
        string Primechanie_Cells;
        string items_period_text;
        static public string Period_S;
        static public string Period_Po;
        private async void CAD_6_Click(object sender, EventArgs e)
        {
            DBOracleUtils db = new DBOracleUtils();
            OracleConnection conn = db.GetDBConnection();

            //if ((Period_S != "") || (Period_Po != ""))
            //{
            //    MessageBox.Show("Не выбран период!",
            //        "Error");
            //}
            //else
            //{

                try
                {
                    conn.Open();
                    // sum_fakt_full = 0;
                    load.Visible = true;
                    ojidan.Text = "Ожидание . . .";
                    ojidan.Visible = true;
                    var rr = await Task.Run(() => Query_6_CAD(conn, arrTM_depart));
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
            
        }

        private int Query_6_CAD(OracleConnection conn, int arrTM_depart)
        {
            massiv._Sotrudniki6KAD.Clear();
            string Sotrudnik_Prinyat = "";
            string YvolSQL = "";
            string Peroid_Parametr = " ((MAIN.LAST >= TO_DATE('"+ Period_S + "', 'yyyy-mm-dd')) AND (MAIN.LAST <= TO_DATE('" + Period_Po + "', 'yyyy-mm-dd'))) and ((MAIN.DATEEXECUTION >= TO_DATE('" + Period_S + "', 'yyyy-mm-dd')) AND (MAIN.DATEEXECUTION <= TO_DATE('" + Period_Po + "', 'yyyy-mm-dd'))) ";
            // "((TRUNC(MAIN_AGNRN.LAST, 'Q') = TRUNC(ADD_MONTHS(SYSDATE, -3), 'q')) OR (TRUNC(MAIN_AGNRN.LAST, 'Q')=ADD_MONTHS(TRUNC(ADD_MONTHS(SYSDATE, -3), 'q'),3)-1))";
            
            #region
            ////периоды выборки
            //if (items_period == 0) //текущая неделя
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN_AGNRN.LAST, 'DAY') = TRUNC(SYSDATE, 'DAY')) or (TRUNC(MAIN_AGNRN.LAST, 'DAY')=TRUNC(SYSDATE, 'DAY')+6)) ";
            //}
            //if (items_period == 1)//прошлая неделя
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN_AGNRN.LAST, 'DAY') = TRUNC(SYSDATE, 'DAY')-7) or (TRUNC(MAIN_AGNRN.LAST, 'DAY')=TRUNC(SYSDATE, 'DAY')-1)) ";
            //}
            //if (items_period == 2) //текущий месяц
            //{
            //    Peroid_Parametr = " TRUNC(MAIN_AGNRN.LAST, 'MM') = TRUNC(SYSDATE, 'MM') ";
            //}
            //if (items_period == 3)//прошлый месяц
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN_AGNRN.LAST, 'MM') = (TRUNC(ADD_MONTHS(SYSDATE, -1), 'MM'))) OR (TRUNC(MAIN_AGNRN.LAST, 'MM')=TRUNC(SYSDATE, 'MM')-1))  ";
            //}
            //if (items_period == 4)//текущий квартал
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN_AGNRN.LAST, 'Q') = TRUNC(SYSDATE, 'Q')) OR (TRUNC(MAIN_AGNRN.LAST, 'Q')=ADD_MONTHS(TRUNC(SYSDATE, 'q'), 3)-1))  ";
            //}
            //if (items_period == 5)//прошлый квартал
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN_AGNRN.LAST, 'Q') = TRUNC(ADD_MONTHS(SYSDATE, -3), 'q')) OR (TRUNC(MAIN_AGNRN.LAST, 'Q')=ADD_MONTHS(TRUNC(ADD_MONTHS(SYSDATE, -3), 'q'),3)-1)) ";
            //}
            //if (items_period == 6)//текущий год
            //{
            //    Peroid_Parametr = " TRUNC(MAIN_AGNRN.LAST, 'Y')= TRUNC(SYSDATE, 'Y')  ";
            //}
            //if (items_period == 7)//прошлый год
            //{
            //    Peroid_Parametr = " TRUNC(MAIN_AGNRN.LAST, 'Y') = ADD_MONTHS(TRUNC(SYSDATE, 'YEAR'),-12) ";
            //}
            #endregion

            
            #region

            //периоды выборки
            //if (items_period == 0) //текущая неделя
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN.LAST, 'DAY') = TRUNC(SYSDATE, 'DAY')) or (TRUNC(MAIN.LAST, 'DAY')=TRUNC(SYSDATE, 'DAY')+6)) and ((TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'DAY') = TRUNC(SYSDATE, 'DAY')) or (TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'DAY')=TRUNC(SYSDATE, 'DAY')+6))  ";
            //}
            //if (items_period == 1)//прошлая неделя
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN.LAST, 'DAY') = TRUNC(SYSDATE, 'DAY')-7) or (TRUNC(MAIN.LAST, 'DAY')=TRUNC(SYSDATE, 'DAY')-1)) and  ((TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'DAY') = TRUNC(SYSDATE, 'DAY')-7) or (TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'DAY')=TRUNC(SYSDATE, 'DAY')-1)) ";
            //}
            //if (items_period == 2) //текущий месяц
            //{
            //    Peroid_Parametr = " TRUNC(MAIN.LAST, 'MM') = TRUNC(SYSDATE, 'MM') AND TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'MM') = TRUNC(SYSDATE, 'MM') ";
            //}
            //if (items_period == 3)//прошлый месяц
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN.LAST, 'MM') = (TRUNC(ADD_MONTHS(SYSDATE, -1), 'MM'))) OR (TRUNC(MAIN.LAST, 'MM')=TRUNC(SYSDATE, 'MM')-1))  AND ((TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'MM') = (TRUNC(ADD_MONTHS(SYSDATE, -1), 'MM'))) OR (TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'MM')=TRUNC(SYSDATE, 'MM')-1))  ";
            //}
            //if (items_period == 4)//текущий квартал
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN.LAST, 'Q') = TRUNC(SYSDATE, 'Q')) OR (TRUNC(MAIN.LAST, 'Q')=ADD_MONTHS(TRUNC(SYSDATE, 'q'), 3)-1)) AND ((TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'Q') = TRUNC(SYSDATE, 'Q')) OR (TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'Q')=ADD_MONTHS(TRUNC(SYSDATE, 'q'), 3)-1))  ";
            //}
            //if (items_period == 5)//прошлый квартал
            //{
            //    Peroid_Parametr = " ((TRUNC(MAIN.LAST, 'Q') = TRUNC(ADD_MONTHS(SYSDATE, -3), 'q')) OR (TRUNC(MAIN.LAST, 'Q')=ADD_MONTHS(TRUNC(ADD_MONTHS(SYSDATE, -3), 'q'),3)-1)) AND ((TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'Q') = TRUNC(ADD_MONTHS(SYSDATE, -3), 'q')) OR (TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'Q')=ADD_MONTHS(TRUNC(ADD_MONTHS(SYSDATE, -3), 'q'),3)-1)) ";
            //}
            //if (items_period == 6)//текущий год
            //{
            //    Peroid_Parametr = " TRUNC(MAIN.LAST, 'Y') = TRUNC(SYSDATE, 'Y')   "; //and TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'Y') = TRUNC(SYSDATE, 'Y')
            //}
            //if (items_period == 7)//прошлый год
            //{
            //    Peroid_Parametr = " TRUNC(MAIN.LAST, 'Y') = ADD_MONTHS(TRUNC(SYSDATE, 'YEAR'),-12) and TRUNC(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'Y') = ADD_MONTHS(TRUNC(SYSDATE, 'YEAR'),-12) ";
            //}
            #endregion

            if(arrTM_TO == "Аппарат Центрального таможенного управления")
            {
                arrTM_TO = "Аппарат управления Центрального таможенного управления";
            }
            
            if (arrTM_depart == 0)
            {
                Sotrudnik_Prinyat = "SELECT MAIN_crn.NAME AS Name_TO,   MAIN.MOVORG AS Mnemo_TO,          MAIN_crn.PARENT AS Verhnii_Level,          MAIN.INICIALI AS FIO,           MAIN.SEX AS Pol,          MAIN.AGE AS Vozrast,           MAIN.EXEC AS Doljnost,       MAIN.CLNGROUP AS Gruppa_Doljnost,           MAIN.PRDTKN AS Status_Slujb,           MAIN.SVCSLUJ AS Vid_Sluj,           TO_CHAR(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'dd.MM.yyyy') AS Data_Naznachenia, MAIN.SHORTEXEC AS Mnemo_Dolj, MAIN.CLNKAD AS Kat_Dolj,           MAIN.EDVID AS Obrazov, MAIN.EDPROF AS Vid_Obrazov,  MAIN.EDSPEC AS Specialnost,TO_CHAR(MAIN.LAST, 'dd.MM.yyyy') AS Date_Priema, MAIN.SVCKAT AS Otkyda FROM(SELECT a.AGNFAMILYNAME || ' ' || a.AGNFIRSTNAME || ' ' || a.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(a.RN) AS INICIALI, TRUNC(MONTHS_BETWEEN(SYSDATE, a.AGNBURN) / 12) AS AGE, a.CRN AS crn,               dm_query.get_premplsts(a.RN) AS PRDTKN, DECODE(a.SEX, 1, 'муж', 2, 'жен') AS SEX, a.PRFMLSTS AS SEM,                dm_query.get_latter_date(a.RN) AS LATTER, dm_query.get_min_datebg(a.RN) AS LAST, dm_query.get_agneduc_vidob(a.RN) AS EDVID, dm_query.get_agneduc_prof(a.RN) AS EDPROF, dm_query.get_agneduc_spec(a.RN) AS EDSPEC, dm_query.get_agnsvcprd_sluj(a.RN) AS SVCSLUJ, dm_query.get_agnsvcprd_kat(a.RN) AS SVCKAT, dm_query.get_agnsvcmov_org(a.RN) AS MOVORG, dm_query.get_max_execution(a.RN) AS EXEC, dm_query.get_max_shortexecution(a.RN) AS SHORTEXEC, dm_query.get_max_dateexecution(a.RN) AS DATEEXECUTION, dm_query.get_clnpsdep_kad(a.RN) AS CLNKAD, dm_query.get_clnpsdep_group(a.RN) AS CLNGROUP       FROM UK_PARUS.AGNLIST a      WHERE a.AGNTYPE = 1        AND a.RN IN(SELECT pp.AGNLIST FROM UK_PARUS.PREMPLFLS pp)) MAIN LEFT OUTER JOIN(SELECT T.RN AS RN, T.NAME AS NAME, DECODE(T.IS_ROOT, 0, 'нет', 1, 'да') AS IS_ROOT,                            (SELECT A.NAME                                FROM UK_PARUS.ACATALOG A                                WHERE LEVEL = 2                              CONNECT BY PRIOR a.crn = a.rn                              START WITH a.rn = t.rn) AS PARENT        FROM UK_PARUS.ACATALOG T) MAIN_crn ON MAIN_crn.RN = MAIN.crn  WHERE  " + Peroid_Parametr + "AND MAIN.SVCKAT <> 'Внутри т/о'    AND MAIN.SVCKAT <> 'из др.там.органа'    AND MAIN.SVCSLUJ <> 'Работник'    AND EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP        WHERE UP.CATALOG = MAIN.crn)  ORDER BY MAIN.LATTER";
                YvolSQL = "SELECT SubQuery.DATA_YVOLN AS Date_Yvol, SubQuery.FIO AS FIO, SubQuery.KATALOG AS Katalog, SUBQUERY.TORG AS TORG, SUBQUERY.AGENT_RN AS Agen_RN, SUBQUERY.VOZRAST AS Vozrast, SUBQUERY.SEX AS POL, SUBQUERY.DATA_NACH AS Date_Nach, SUBQUERY.DATA_FIN AS Date_Fin, SUBQUERY.VID_SLUJB AS Vid_Slujb, SUBQUERY.VID_OBRAZ AS Vid_Obraz, SubQuery_1.NAME as Motiv, SUBQUERY.MNEMO_DOLJ AS Mnemo_Dolj, SUBQUERY.GRUPPA_DOLJNOST AS Gruppa_Dolj, SUBQUERY.KAT_DOLJ AS Kat_Dolj FROM (SELECT MAIN.MOVEDATE AS Data_Yvoln, MAIN.INICIALI AS FIO, MAIN_crn.NAME AS Katalog, MAIN_crn.PARENT AS TORG, MAIN.RN AS AGENT_RN, MAIN.AGE AS Vozrast, MAIN.SVCPOST AS Data_Nach, MAIN.SVCPOST2 AS Data_Fin, MAIN.SVCSLUJ AS Vid_Slujb, MAIN.SEX AS SEX , MAIN.EDVID AS Vid_Obraz, MAIN.SHORTEXEC AS Mnemo_Dolj, MAIN.CLNGROUP AS Gruppa_Doljnost, MAIN.CLNKAD AS Kat_Dolj FROM (SELECT a.AGNFAMILYNAME AS agnfamilyname, a.AGNFIRSTNAME AS agnfirstname, a.AGNLASTNAME AS agnlastname, a.AGNFAMILYNAME || ' ' || a.AGNFIRSTNAME || ' ' || a.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(a.RN) AS INICIALI, a.RN AS RN, a.AGNBURN AS agnburn, TRUNC(MONTHS_BETWEEN(SYSDATE, a.AGNBURN) / 12) AS AGE, dm_query.get_prem_ln(a.RN) AS licnumb, dm_query.get_depart_uprav_ca(a.RN) AS UPRAV, dm_query.get_depart_upravshort_ca(a.RN) AS UPRAV_SHORT, dm_query.get_discharge1(a.RN) AS discharge, dm_query.get_depart(a.RN) AS DEPART, a.AGNABBR AS AGNCODE, a.CRN AS crn, dm_query.get_premplsts(a.RN) AS PRDTKN, DECODE(a.SEX, 1, 'муж', 2, 'жен') AS SEX, a.PRFMLSTS AS SEM, a.PENSION_NBR AS NBR, a.AGNIDNUMB AS ID, a.PHONE AS PHONE, a.PHONE2 AS PHONE2, a.TELEX AS TELEX, DM_QUERY.GET_STAJ(a.RN) AS ssta, dm_query.GET_STAJ_KAD2(a.RN) AS dddss, dm_query.get_agnsvcprd(a.RN) AS TORG, dm_query.get_first_date(a.RN) AS FIRST, dm_query.get_latter_date(a.RN) AS LATTER, dm_query.get_min_datebg(a.RN) AS LAST, dm_query.get_isp_contr(a.RN) AS CONTR, dm_query.get_isp_contrnom(a.RN) AS CONNOM, dm_query.get_isp_contrvid(a.RN) AS CONVID, dm_query.get_isp_contrend(a.RN) AS DATEND, dm_query.get_isp_controsn(a.RN) AS OSN, dm_query.get_isp_contrdopnom(a.RN) AS DOPNOM, dm_query.get_isp_contrdopdat(a.RN) AS DOPDAT, dm_query.get_isp_contrdopopis(a.RN) AS DOPOPIS, dm_query.get_isp_contrdopnom40(a.RN) AS DOPNOM40, dm_query.get_isp_contrdopdat40(a.RN) AS DOPDAT40, dm_query.get_isp_contrdopopis40(a.RN) AS DOPOPIS40, dm_query.get_pris(a.RN) AS PRIS, dm_query.get_provdat(a.RN) AS PROV, dm_query.get_provnom(a.RN) AS PROVNOM, dm_query.get_dak(a.RN) AS DAK, dm_query.get_attdat(a.RN) AS ATTDAT, dm_query.get_attns_date(a.RN) AS ATTNSDATE, dm_query.get_attdatnext(a.RN) AS ATTDATNE, dm_query.get_attosn(a.RN) AS ATTOSN, dm_query.get_attviv(a.RN) AS ATTVIV, dm_query.get_attdoctype(a.RN) AS ATTDOCTYPE, dm_query.get_attwho(a.RN) AS ATTWHO, dm_query.get_attdocdat(a.RN) AS ATTDOCDAT, dm_query.get_attdocnom(a.RN) AS ATTDOCNOM, DECODE(dm_query.get_attrlz(a.RN), 1, 'Да', 0, 'нет') AS ATTRLZ, dm_query.get_attdoctype_rlz(a.RN) AS ATTRLZTYPE, dm_query.get_attdocwho_rlz(a.RN) AS ATTRLZWHO, dm_query.get_attdocdat_rlz(a.RN) AS ATTRLZDAT, dm_query.get_attdocnom_rlz(a.RN) AS ATTRLZNOM, dm_query.get_agnrwrd_dat(a.RN) AS WRDDAT, dm_query.get_agnrwrd_who(a.RN) AS WRDWHO, dm_query.get_agnrwrd_numb(a.RN) AS WRDNUMB, dm_query.get_agnrwrd_type(a.RN) AS WRDTYPE, dm_query.get_agnrwrd_typ(a.RN) AS WRDTYP, dm_query.get_agnrwrd_vid(a.RN) AS WRDVID, dm_query.get_agnrwrd_reas(a.RN) AS WRDREAS, dm_query.get_agnpnlts_knd(a.RN) AS KND, dm_query.get_calc_length(a.RN) AS CALCLEN, dm_query.get_agnpnlts_fault(a.RN) AS FOULT, dm_query.get_agnpnlts_nlk(a.RN) AS NLK, dm_query.get_agnpnlts_css(a.RN) AS CSS, dm_query.get_agnpnlts_ini(a.RN) AS INI, dm_query.get_agnpnlts_pnl(a.RN) AS PNL, dm_query.get_agnpnlts_typeon(a.RN) AS TYPEON, dm_query.get_agnpnlts_whonal(a.RN) AS WHONAL, dm_query.get_agnpnlts_docon(a.RN) AS DOCON, dm_query.get_agnpnlts_docon2(a.RN) AS DOCON2, dm_query.get_agnpnlts_cntrldat(a.RN) AS CNTR, dm_query.get_agnpnlts_appl(a.RN) AS APPL, dm_query.get_agnpnlts_remov(a.RN) AS REMOV, dm_query.get_agnpnlts_osn(a.RN) AS OSN1, dm_query.get_agnpnlts_typeoff(a.RN) AS TYPEOFF, dm_query.get_agnpnlts_whosnal(a.RN) AS WHOSNAL, dm_query.get_agnpnlts_datoff(a.RN) AS DATOFF, dm_query.get_agnpnlts_docoff(a.RN) AS DOCOFF, dm_query.get_agnmedcer_beg(a.RN) AS MEDBEG, dm_query.get_agnmedcer_end(a.RN) AS MEDEND, dm_query.get_agnmedcer_pay(a.RN) AS MEDPAY, DECODE(dm_query.get_agnmedcer_hosp(a.RN), 1, 'Да', 0, 'нет') AS HOSP, dm_query.get_agnmedcer_ser(a.RN) AS MEDSER, dm_query.get_agnmedcer_nom(a.RN) AS MEDNOM, dm_query.get_agnmedcer_dat(a.RN) AS MEDDAT, dm_query.get_agnmedcer_med(a.RN) AS MED, dm_query.get_clnpersvac_vid(a.RN) AS OTPVID, DECODE(dm_query.get_clnpersvac_type(a.RN), 1, 'Дополнительный', 0, 'Основной', 2, 'За свой счет') AS OTPTYPE, dm_query.get_clnpersvac_beg(a.RN) AS OTPBEG, dm_query.get_clnpersvac_end(a.RN) AS OTPEND, dm_query.get_clnpersvac_periodon(a.RN) AS OTPPERON, dm_query.get_clnpersvac_periodoff(a.RN) AS OTPPEROFF, dm_query.get_clnpersvac_planon(a.RN) AS OTPPLANON, dm_query.get_clnpersvac_planoff(a.RN) AS OTPPLANOFF, dm_query.get_clnpersvac_major(a.RN) AS OTPMAJOR, dm_query.get_clnpersvac_minor(a.RN) AS OTPMINOR, dm_query.get_clnpersvac_hol(a.RN) AS OTPHOL, dm_query.get_clnpersvac_trav(a.RN) AS OTPTRAV, dm_query.get_clnpersvac_doctype(a.RN) AS OTPDOCTYPE, dm_query.get_clnpersvac_docwho(a.RN) AS OTPDOCWHO, dm_query.get_clnpersvac_docdat(a.RN) AS OTPDOCDAT, dm_query.get_clnpersvac_docnumb(a.RN) AS OTPDOCNUMB, dm_query.get_clnpersvac_prim(a.RN) AS OTPPRIM, DECODE(dm_query.get_clnpersvac_censel(a.RN), 1, 'Да', 0, 'Нет') AS OTPCENSEL, DECODE(dm_query.get_clnpersvac_hold(a.RN), 1, 'Да', 0, 'Нет') AS OTPHOLD, dm_query.get_clnpersvac_ret(a.RN) AS OTPRET, dm_query.get_clnpersvac_hnote(a.RN) AS OTPHNOTE, dm_query.get_clnpersvac_rettype(a.RN) AS OTPRETTYPE, dm_query.get_clnpersvac_retwho(a.RN) AS OTPRETWHO, dm_query.get_clnpersvac_retdat(a.RN) AS OTPRETDAT, dm_query.get_clnpersvac_retnumb(a.RN) AS OTPRETNUMB, dm_query.get_clnperstrip_beg(a.RN) AS TRIPBEG, dm_query.get_clnperstrip_end(a.RN) AS TRIPEND, dm_query.get_clnperstrip_cel(a.RN) AS TRIPCEL, DECODE(dm_query.get_clnperstrip_cansel(a.RN), 1, 'Да', 0, 'Нет') AS TRIPCANSEL, dm_query.get_clnperstrip_goal(a.RN) AS TRIPGOAL, dm_query.get_clnperstrip_type(a.RN) AS TRIPTYPE, dm_query.get_clnperstrip_who(a.RN) AS TRIPWHO, dm_query.get_clnperstrip_dat(a.RN) AS TRIPDAT, dm_query.get_clnperstrip_numb(a.RN) AS TRIPNUMB, dm_query.get_clnperstrip_adres(a.RN) AS TRIPADRES, dm_query.get_clnperstrip_dopadres(a.RN) AS TRIPDOPADRES, dm_query.get_agnfrntrp_geog(a.RN) AS TRPGEOG, dm_query.get_agnfrntrp_trp(a.RN) AS TRPTRP, dm_query.get_agnfrntrp_beg(a.RN) AS TRPBEG, dm_query.get_agnfrntrp_end2(a.RN) AS TRPEND, dm_query.get_agnaddinf_type(a.RN) AS DDTYPE, dm_query.get_agnaddinf_datbeg(a.RN) AS DDDATBEG, dm_query.get_agnaddinf_datend(a.RN) AS DDDATEND, dm_query.get_agnaddinf_sod(a.RN) AS DDSOD, dm_query.get_sppvd_vid(a.RN) AS SPPVID, dm_query.get_sppvd_numb(a.RN) AS SPPNUMB, dm_query.get_sppvd_docdat(a.RN) AS SPPDOCDAT, dm_query.get_sppvd_docwho(a.RN) AS SPPDOCWHO, dm_query.get_sppvd_vidpro(a.RN) AS SPPVIDPRO, dm_query.get_sppvd_datpro(a.RN) AS SPPDATPRO, dm_query.get_sppvd_opis(a.RN) AS SPPOPIS, dm_query.get_sppvd_stat(a.RN) AS SPPSTAT, dm_query.get_sppvd_mesto(a.RN) AS SPPMESTO, DECODE(dm_query.get_sppvd_vina(a.RN), 1, 'Да', 0, 'Нет') AS SPPVINA, dm_query.get_sppvd_oznak(a.RN) AS SPPOZNAK, DECODE(dm_query.get_sppvd_noznak(a.RN), 1, 'Да', 0, 'Нет') AS SPPNOZNAK, dm_query.get_sppvd_prim(a.RN) AS SPPPRIM, dm_query.get_agnacclev_form(a.RN) AS CCFORM, dm_query.get_agnacclev_s(a.RN) AS CCS, dm_query.get_agnacclev_po(a.RN) AS CCPO, dm_query.get_agnacclev_prim(a.RN) AS CCPRIM, dm_query.get_agnacclev_type(a.RN) AS CCTYPE, dm_query.get_agnacclev_who(a.RN) AS CCWHO, dm_query.get_agnacclev_dat(a.RN) AS CCDAT, dm_query.get_agnacclev_numb(a.RN) AS CCNUMB, dm_query.get_profpsih_dat(a.RN) AS SIHDAT, dm_query.get_profpsih_kat(a.RN) AS SIHKAT, dm_query.get_profpsih_zak(a.RN) AS SIHZAK, dm_query.get_profpsih_note(a.RN) AS SIHNOTE, dm_query.get_profvalid_dat(a.RN) AS LIDDAT, dm_query.get_profvalid_kat(a.RN) AS LIDKAT, dm_query.get_profvalid_who(a.RN) AS LIDWHO, dm_query.get_profvalid_nom(a.RN) AS LIDNOM, dm_query.get_profvalid_note(a.RN) AS LIDNOTE, dm_query.get_agnranks_zvan(a.RN) AS NKSZVAN, dm_query.get_agnranks_vid(a.RN) AS NKSVID, dm_query.get_agnranks_dat(a.RN) AS NKSDAT, dm_query.get_agnranks_ndat(a.RN) AS NKSNDAT, dm_query.get_agnranks_type(a.RN) AS NKSTYPE, dm_query.get_agnranks_who(a.RN) AS NKSWHO, dm_query.get_agnranks_prdat(a.RN) AS NKSPRDAT, dm_query.get_agnranks_nom(a.RN) AS NKSNOM, dm_query.get_agneduc_vidob(a.RN) AS EDVID, dm_query.get_agneduc_prof(a.RN) AS EDPROF, dm_query.get_agneduc_begdat(a.RN) AS EDBEGDAT, dm_query.get_agneduc_enddat(a.RN) AS EDENDDAT, dm_query.get_agneduc_educ(a.RN) AS EDEDUC, dm_query.get_agneduc_form(a.RN) AS EDFORM, DECODE(dm_query.get_agneduc_otm(a.RN), 1, 'Учится', 0, 'Не задана', 2, 'Учится на последнем курсе') AS EDOTM, dm_query.get_agneduc_dtype(a.RN) AS EDDTYPE, dm_query.get_agneduc_ser(a.RN) AS EDSER, dm_query.get_agneduc_nom(a.RN) AS EDNOM, dm_query.get_agneduc_spec(a.RN) AS EDSPEC, dm_query.get_agneduc_kval(a.RN) AS EDKVAL, dm_query.get_agneduc_type(a.RN) AS EDTYPE, dm_query.get_agneduc_who(a.RN) AS EDWHO, dm_query.get_agneduc_ddoc(a.RN) AS EDDDOC, dm_query.get_agneduc_dnom(a.RN) AS EDDNOM, dm_query.get_agneducperv_mesto(a.RN) AS EDPERVMESTO, dm_query.get_agneducperv_tipdoc(a.RN) AS EDPERVTIPDOC, dm_query.get_agneducperv_ser(a.RN) AS EDPERVSER, dm_query.get_agneducperv_nom(a.RN) AS EDPERVNOM, dm_query.get_agneducperv_bdate(a.RN) AS EDPERVBDATE, dm_query.get_agneducperv_edate(a.RN) AS EDPERVEDATE, dm_query.get_agneducperv_type(a.RN) AS EDPERVTYPE, dm_query.get_agneducperv_who(a.RN) AS EDPERVWHO, dm_query.get_agneducperv_dat(a.RN) AS EDPERVDAT, dm_query.get_agneducperv_numb(a.RN) AS EDPERVNUMB, dm_query.get_agneducpov_vidobr(a.RN) AS POVVIDOBR, dm_query.get_agneducpov_uch(a.RN) AS POVUCH, dm_query.get_agneducpov_kurs(a.RN) AS POVKURS, dm_query.get_agneducpov_form(a.RN) AS POVFORM, dm_query.get_agneducpov_fin(a.RN) AS POVFIN, dm_query.get_agneducpov_prof(a.RN) AS POVPROF, dm_query.get_agneducpov_bdate(a.RN) AS POVBDATE, dm_query.get_agneducpov_edate(a.RN) AS POVEDATE, dm_query.get_agneducpov_chas(a.RN) AS POVCHAS, dm_query.get_agneducpov_doctype(a.RN) AS POVDOCTYPE, dm_query.get_agneducpov_docser(a.RN) AS POVDOCSER, dm_query.get_agneducpov_docnom(a.RN) AS POVDOCNOM, dm_query.get_agneducpov_type(a.RN) AS POVTYPE, dm_query.get_agneducpov_who(a.RN) AS POVWHO, dm_query.get_agneducpov_dat(a.RN) AS POVDAT, dm_query.get_agneducpov_nom(a.RN) AS POVNOM, dm_query.get_agneducper_vid(a.RN) AS PERVID, dm_query.get_agneducper_uch(a.RN) AS PERUCH, dm_query.get_agneducper_kurs(a.RN) AS PERKURS, dm_query.get_agneducper_form(a.RN) AS PERFORM, dm_query.get_agneducper_fin(a.RN) AS PERFIN, dm_query.get_agneducper_prof(a.RN) AS PERPROF, dm_query.get_agneducper_bdate(a.RN) AS PERBDATE, dm_query.get_agneducper_edate(a.RN) AS PEREDATE, dm_query.get_agneducper_chas(a.RN) AS PERCHAS, dm_query.get_agneducper_doctype(a.RN) AS PERDOCTYPE, dm_query.get_agneducper_docser(a.RN) AS PERDOCSER, dm_query.get_agneducper_docnom(a.RN) AS PERDOCNOM, dm_query.get_agneducper_type(a.RN) AS PERTYPE, dm_query.get_agneducper_who(a.RN) AS PERWHO, dm_query.get_agneducper_dat(a.RN) AS PERDAT, dm_query.get_agneducper_nom(a.RN) AS PERNOM, dm_query.get_agneducstaj_mesto(a.RN) AS STAJMESTO, dm_query.get_agneducstaj_fin(a.RN) AS STAJFIN, dm_query.get_agneducstaj_bdate(a.RN) AS STAJBDATE, dm_query.get_agneducstaj_edate(a.RN) AS STAJEDATE, dm_query.get_agneducstaj_type(a.RN) AS STAJTYPE, dm_query.get_agneducstaj_who(a.RN) AS STAJWHO, dm_query.get_agneducstaj_dat(a.RN) AS STAJDAT, dm_query.get_agneducstaj_nom(a.RN) AS STAJNOM, dm_query.get_agneducstep_uch(a.RN) AS STEPUCH, dm_query.get_agneducstep_spec(a.RN) AS STEPSPEC, dm_query.get_agneducstep_dat(a.RN) AS STEPDAT, dm_query.get_agneducstep_type(a.RN) AS STEPTYPE, dm_query.get_agneducstep_ser(a.RN) AS STEPSER, dm_query.get_agneducstep_nom(a.RN) AS STEPNOM, dm_query.get_agneducstep_note(a.RN) AS STEPNOTE, dm_query.get_agneduczvan_uch(a.RN) AS ZVANUCH, dm_query.get_agneduczvan_dat(a.RN) AS ZVANDAT, dm_query.get_agneduczvan_type(a.RN) AS ZVANTYPE, dm_query.get_agneduczvan_ser(a.RN) AS ZVANSER, dm_query.get_agneduczvan_nom(a.RN) AS ZVANNOM, dm_query.get_agneduczvan_note(a.RN) AS ZVANNOTE, dm_query.get_agnlangs_in(a.RN) AS LANIN, dm_query.get_agnlangs_step(a.RN) AS LANSTEP, dm_query.get_agnlangs_note(a.RN) AS LANNOTE, dm_query.get_agnatnms_vid(a.RN) AS NMSVID, dm_query.get_agnatnms_step(a.RN) AS NMSSTEP, dm_query.get_docsprops_agnatnms(a.RN) AS DPNM, dm_query.get_agnaddresses_full(a.RN) AS ADRESFULL, dm_query.get_agnaddresses_index(a.RN) AS ADRESINDEX, dm_query.get_agnaddresses_dom(a.RN) AS ADRESDOM, dm_query.get_agnaddresses_blok(a.RN) AS ADRESBLOK, dm_query.get_agnaddresses_flat(a.RN) AS ADRESFLAT, dm_query.get_agnaddresses_note(a.RN) AS ADRESNOTE, dm_query.get_agndocums_type(a.RN) AS CUMSTYPE, dm_query.get_agndocums_ser(a.RN) AS CUMSSER, dm_query.get_agndocums_numb(a.RN) AS CUMSNUMB, dm_query.get_agndocums_vida(a.RN) AS CUMSVIDA, dm_query.get_agndocums_po(a.RN) AS CUMSPO, dm_query.get_agndocums_who(a.RN) AS CUMSWHO, dm_query.get_agndocums_osn(a.RN) AS CUMSOSN, dm_query.get_agndocums_doctype(a.RN) AS CUMSDOCTYPE, dm_query.get_agndocums_docwho(a.RN) AS CUMSDOCWHO, dm_query.get_agndocums_docdat(a.RN) AS CUMSDOCDAT, dm_query.get_agndocums_docnumb(a.RN) AS CUMSDOCNUMB, dm_query.get_agnsprtsk_kval(a.RN) AS TSKKVAL, dm_query.get_agnsprtsk_vid(a.RN) AS TSKVID, dm_query.get_agnsprtsk_god(a.RN) AS TSKGOD, dm_query.get_agnlbractivityprs_bdate(a.RN) AS TRUDBDATE, dm_query.get_agnlbractivityprs_edate(a.RN) AS TRUDEDATE, dm_query.get_agnlbractivityprs_gos(a.RN) AS TRUDGOS, dm_query.get_agnlbractivityprs_work(a.RN) AS TRUDWORK, dm_query.get_agnlbractivityprs_dolj(a.RN) AS TRUDDOLJ, dm_query.get_agnlbractivityprs_prich(a.RN) AS TRUDPRICH, dm_query.get_agnlbractivityprs_type(a.RN) AS TRUDTYPE, dm_query.get_agnlbractivityprs_who(a.RN) AS TRUDWHO, dm_query.get_agnlbractivityprs_date(a.RN) AS TRUDDATE, dm_query.get_agnlbractivityprs_numb(a.RN) AS TRUDNUMB, dm_query.get_agnlbractivityprs_note(a.RN) AS TRUDNOTE, dm_query.get_agnsvcprd_sluj(a.RN) AS SVCSLUJ, dm_query.get_agnsvcprd_post(a.RN) AS SVCPOST, dm_query.get_agnsvcprd_post2(a.RN) AS SVCPOST2, dm_query.get_agnsvcprd_vidpost(a.RN) AS SVCVIDPOST, dm_query.get_agnsvcprd_kat(a.RN) AS SVCKAT, dm_query.get_agnsvcprd_primtype(a.RN) AS SVCPRIMTYPE, dm_query.get_agnsvcprd_primwho(a.RN) AS SVCPRIMWHO, dm_query.get_agnsvcprd_primdate(a.RN) AS SVCPRIMDATE, dm_query.get_agnsvcprd_primnumb(a.RN) AS SVCPRIMNUMB, dm_query.get_agnsvcprd_ispdat(a.RN) AS SVCISPDAT, dm_query.get_agnsvcprd_ispedat(a.RN) AS SVCISPEDAT, dm_query.get_agnsvcprd_ispzak(a.RN) AS SVCISPZAK, dm_query.get_agnsvcprd_isptype(a.RN) AS SVCISPTYPE, dm_query.get_agnsvcprd_ispwho(a.RN) AS SVCISPWHO, dm_query.get_agnsvcprd_ispdate(a.RN) AS SVCISPDATE, dm_query.get_agnsvcprd_ispnumb(a.RN) AS SVCISPNUMB, dm_query.get_agnsvcmov_bdate(a.RN) AS MOVBDATE, dm_query.get_agnsvcmov_edate(a.RN) AS MOVEDATE, dm_query.get_agnsvcmov_pervid(a.RN) AS MOVPERVID, dm_query.get_agnsvcmov_perprich(a.RN) AS MOVPERPRICH, dm_query.get_agnsvcmov_permot(a.RN) AS MOVPERMOT, dm_query.get_agnsvcmov_mesto(a.RN) AS MOVMESTO, dm_query.get_agnsvcmov_org(a.RN) AS MOVORG, dm_query.get_agnsvcmov_block(a.RN) AS MOVBLOCK, dm_query.get_agnsvcmov_group(a.RN) AS MOVGROUP, dm_query.get_agnsvcmov_kod(a.RN) AS MOVKOD, dm_query.get_agnsvcmov_dolj(a.RN) AS MOVDOLJ, dm_query.get_agnsvcmov_fdolj(a.RN) AS MOVFDOLJ, dm_query.get_agnsvcmov_type(a.RN) AS MOVTYPE, dm_query.get_agnsvcmov_who(a.RN) AS MOVWHO, dm_query.get_agnsvcmov_date(a.RN) AS MOVDATE, dm_query.get_agnsvcmov_numb(a.RN) AS MOVNUMB, dm_query.get_agnsvcmov_prekprich(a.RN) AS MOVPRIKPRICH, dm_query.get_agnsvcmov_prektype(a.RN) AS MOVPRIKTYPE, dm_query.get_agnsvcmov_prekwho(a.RN) AS MOVPRIKWHO, dm_query.get_agnsvcmov_prekdate(a.RN) AS MOVPRIKDATE, dm_query.get_agnsvcmov_preknumb(a.RN) AS MOVPRIKNUMB, dm_query.get_agnsvcmov_note(a.RN) AS MOVNOTE, dm_query.get_max_execution(a.RN) AS EXEC, dm_query.get_max_shortexecution(a.RN) AS SHORTEXEC, dm_query.get_max_dateexecution(a.RN) AS DATEEXECUTION, dm_query.get_clnpsdep_code(a.RN) AS CLNCODE, dm_query.get_clnpsdep_zvan2(a.RN) AS CLNZVAN, dm_query.get_clnpsdep_doldat(a.RN) AS CLNDOLDAT, dm_query.get_clnpsdep_type(a.RN) AS CLNTYPE, dm_query.get_clnpsdep_who(a.RN) AS CLNWHO, dm_query.get_clnpsdep_date(a.RN) AS CLNDATE, dm_query.get_clnpsdep_numb(a.RN) AS CLNNUMB, dm_query.get_clnpsdep_kod(a.RN) AS CLNKOD, dm_query.get_clnpsdep_kad(a.RN) AS CLNKAD, dm_query.get_clnpsdep_group(a.RN) AS CLNGROUP FROM UK_PARUS.AGNLIST a WHERE a.AGNTYPE = 1 AND a.RN IN (SELECT pp.AGNLIST FROM UK_PARUS.PREMPLFLS pp)) MAIN LEFT OUTER JOIN (SELECT T.RN AS RN, T.NAME AS NAME, DECODE(T.IS_ROOT, 0, 'нет', 1, 'да') AS IS_ROOT, (SELECT A.NAME FROM UK_PARUS.ACATALOG A WHERE LEVEL = 2 CONNECT BY PRIOR a.crn = a.rn START WITH a.rn = t.rn) AS PARENT FROM UK_PARUS.ACATALOG T) MAIN_crn ON MAIN_crn.RN = MAIN.crn WHERE MAIN_crn.NAME LIKE 'Увол%' AND ((MAIN.MOVEDATE >= TO_DATE('" + Period_S + "', 'yyyy-mm-dd')) AND (MAIN.MOVEDATE <= TO_DATE('" + Period_Po + "', 'yyyy-mm-dd'))) AND EXISTS (SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP WHERE UP.CATALOG = MAIN.crn)) SubQuery INNER JOIN (SELECT SubQuery.PRN, SubQuery.NAME, SubQuery.STOP_DATE FROM (SELECT AGNSVCPRD.PRDISMTV, AGNSVCPRD.PRN, SubQuery.NAME, AGNSVCPRD.STOP_DATE FROM (SELECT AGNSVCPRD.PRN, AGNSVCPRD.PRDISMTV, AGNSVCPRD.STOP_DATE FROM UK_PARUS.AGNSVCPRD) AGNSVCPRD INNER JOIN (SELECT PRDISMTV.RN, PRDISMTV.NAME, PRDISMTV.RN AS EXPR1 FROM (SELECT PRDISMTV.RN, PRDISMTV.NAME FROM UK_PARUS.PRDISMTV) PRDISMTV) SubQuery ON AGNSVCPRD.PRDISMTV = SubQuery.RN) SubQuery) SubQuery_1 ON SubQuery.AGENT_RN = SubQuery_1.PRN AND SubQuery.Data_Yvoln = SubQuery_1.STOP_DATE";

                //Sotrudnik_Prinyat = "SELECT MAIN_CRN.NAME AS Name_TO, MAIN.EXEC AS Doljnost, MAIN_AGNRN.CLNGROUP AS Gruppa_Doljnost , MAIN_AGNRN.SEX AS Pol, MAIN.CLNKAD AS MAIN_CLNKAD, MAIN.SLVID AS Vid_Sluj, MAIN_AGNRN.SHORTEXEC AS Mnemo_Dolj, MAIN_AGNRN.CLNKAD AS Kat_Dolj, MAIN_AGNRN.EDVID AS Obrazov, MAIN_AGNRN.EDPROF AS Vid_Obrazov, MAIN_AGNRN.EDSPEC AS Specialnost, MAIN_AGNRN.INICIALI AS FIO, MAIN_AGNRN.AGE AS Vozrast, TO_CHAR(MAIN_AGNRN.LAST, 'dd.MM.yyyy') AS Date_Priema, MAIN_AGNRN.SVCKAT AS Otkyda, MAIN.AGN_RN AS Nomer_Agenta FROM(SELECT dep.RN AS RN, FR_ADD_COL_PERSONS(dep.RN) AS DEP_CODE, dm_query.get_svyz(dep.RN) AS AGN_RN, sts.NAME_NOM AS EXEC, dep.PSDEP_NAME AS CLNCODE, dm_query.get_clnpsdep_zvan(dep.RN) AS CLNZVAN, dep.DO_ACT_FROM AS CLNDOLDAT, dep.DEPART_DISP AS CLNKOD, dm_query.get_clnpsdep_kad(dep.RN) AS CLNKAD, dm_query.get_clnpsdep_group(dep.RN) AS CLNGROUP, dm_query.get_departdisp_post(dep.RN) AS POST, dm_query.get_depart2(dep.RN) AS UPRAV, dep.RATEACC AS KOLSHTAT, DECODE(dep.ON_STAFF, 0, 'внештатная', 1, 'штатная') AS DOLVID, dm_query.get_block(dep.RN) AS BLOCK, dm_query.get_pipec(dep.RN) AS PIPEC, dm_query.get_pipec2(dep.RN) AS PIPEC2, dm_query.get_nomnaz(dep.RN) AS NOMNAZ, dm_query.get_slvid(dep.RN) AS SLVID, dm_query.get_specdol(dep.RN) AS SPECDOL, dm_query.get_sovdol(dep.RN) AS SOVDOL, dm_query.get_depart_uprav_ctu(dep.RN) AS UPRAVL, dep.CRN AS CRN FROM UK_PARUS.CLNPSDEP dep INNER JOIN UK_PARUS.CLNPOSTS sts ON dep.POSTRN = sts.RN INNER JOIN UK_PARUS.CLNPSDEPPRS p ON dep.RN = p.PRN WHERE dep.DO_ACT_FROM <= SYSDATE AND(dep.DO_ACT_TO >= SYSDATE OR dep.DO_ACT_TO IS NULL)) MAIN LEFT OUTER JOIN(SELECT T.RN AS RN, T.NAME AS NAME, DECODE(T.IS_ROOT, 0, 'нет', 1, 'да') AS IS_ROOT, (SELECT A.NAME FROM UK_PARUS.ACATALOG A WHERE LEVEL = 2 CONNECT BY PRIOR a.crn = a.rn START WITH a.rn = t.rn) AS PARENT FROM UK_PARUS.ACATALOG T) MAIN_CRN ON MAIN_CRN.RN = MAIN.CRN LEFT OUTER JOIN(SELECT a.AGNFAMILYNAME AS agnfamilyname, a.AGNFIRSTNAME AS agnfirstname, a.AGNLASTNAME AS agnlastname, a.AGNFAMILYNAME || ' ' || a.AGNFIRSTNAME || ' ' || a.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(a.RN) AS INICIALI, a.RN AS RN, a.AGNBURN AS agnburn, TRUNC(MONTHS_BETWEEN(SYSDATE, a.AGNBURN) / 12) AS AGE, dm_query.get_prem_ln(a.RN) AS licnumb, dm_query.get_depart_uprav_ca(a.RN) AS UPRAV, dm_query.get_depart_upravshort_ca(a.RN) AS UPRAV_SHORT, dm_query.get_discharge1(a.RN) AS discharge, dm_query.get_depart(a.RN) AS DEPART, a.AGNABBR AS AGNCODE, a.CRN AS crn, dm_query.get_premplsts(a.RN) AS PRDTKN, DECODE(a.SEX, 1, 'муж', 2, 'жен') AS SEX, a.PRFMLSTS AS SEM, a.PENSION_NBR AS NBR, a.AGNIDNUMB AS ID, a.PHONE AS PHONE, a.PHONE2 AS PHONE2, a.TELEX AS TELEX, dm_query.get_agnsvcprd(a.RN) AS Name_TO, dm_query.get_first_date(a.RN) AS FIRST, dm_query.get_latter_date(a.RN) AS LATTER, dm_query.get_min_datebg(a.RN) AS LAST, dm_query.get_last_dateend(a.RN) AS To_END, dm_query.get_isp_contr(a.RN) AS CONTR, dm_query.get_isp_contrnom(a.RN) AS CONNOM, dm_query.get_isp_contrvid(a.RN) AS CONVID, dm_query.get_isp_contrend(a.RN) AS DATEND, dm_query.get_isp_controsn(a.RN) AS OSN, dm_query.get_isp_contrdopnom(a.RN) AS DOPNOM, dm_query.get_isp_contrdopdat(a.RN) AS DOPDAT, dm_query.get_isp_contrdopopis(a.RN) AS DOPOPIS, dm_query.get_isp_contrdopnom40(a.RN) AS DOPNOM40, dm_query.get_isp_contrdopdat40(a.RN) AS DOPDAT40, dm_query.get_isp_contrdopopis40(a.RN) AS DOPOPIS40, dm_query.get_pris(a.RN) AS PRIS, dm_query.get_provdat(a.RN) AS PROV, dm_query.get_provnom(a.RN) AS PROVNOM, dm_query.get_dak(a.RN) AS DAK, dm_query.get_attdat(a.RN) AS ATTDAT, dm_query.get_attns_date(a.RN) AS ATTNSDATE, dm_query.get_attdatnext(a.RN) AS ATTDATNE, dm_query.get_attosn(a.RN) AS ATTOSN, dm_query.get_attviv(a.RN) AS ATTVIV, dm_query.get_attdoctype(a.RN) AS ATTDOCTYPE, dm_query.get_attwho(a.RN) AS ATTWHO, dm_query.get_attdocdat(a.RN) AS ATTDOCDAT, dm_query.get_attdocnom(a.RN) AS ATTDOCNOM, DECODE(dm_query.get_attrlz(a.RN), 1, 'Да', 0, 'нет') AS ATTRLZ, dm_query.get_attdoctype_rlz(a.RN) AS ATTRLZTYPE, dm_query.get_attdocwho_rlz(a.RN) AS ATTRLZWHO, dm_query.get_attdocdat_rlz(a.RN) AS ATTRLZDAT, dm_query.get_attdocnom_rlz(a.RN) AS ATTRLZNOM, dm_query.get_agnrwrd_dat(a.RN) AS WRDDAT, dm_query.get_agnrwrd_who(a.RN) AS WRDWHO, dm_query.get_agnrwrd_numb(a.RN) AS WRDNUMB, dm_query.get_agnrwrd_type(a.RN) AS WRDTYPE, dm_query.get_agnrwrd_typ(a.RN) AS WRDTYP, dm_query.get_agnrwrd_vid(a.RN) AS WRDVID, dm_query.get_agnrwrd_reas(a.RN) AS WRDREAS, dm_query.get_agnpnlts_knd(a.RN) AS KND, dm_query.get_agnpnlts_fault(a.RN) AS FOULT, dm_query.get_agnpnlts_nlk(a.RN) AS NLK, dm_query.get_agnpnlts_css(a.RN) AS CSS, dm_query.get_agnpnlts_ini(a.RN) AS INI, dm_query.get_agnpnlts_pnl(a.RN) AS PNL, dm_query.get_agnpnlts_typeon(a.RN) AS TYPEON, dm_query.get_agnpnlts_whonal(a.RN) AS WHONAL, dm_query.get_agnpnlts_docon(a.RN) AS DOCON, dm_query.get_agnpnlts_docon2(a.RN) AS DOCON2, dm_query.get_agnpnlts_cntrldat(a.RN) AS CNTR, dm_query.get_agnpnlts_appl(a.RN) AS APPL, dm_query.get_agnpnlts_remov(a.RN) AS REMOV, dm_query.get_agnpnlts_osn(a.RN) AS OSN1, dm_query.get_agnpnlts_typeoff(a.RN) AS TYPEOFF, dm_query.get_agnpnlts_whosnal(a.RN) AS WHOSNAL, dm_query.get_agnpnlts_datoff(a.RN) AS DATOFF, dm_query.get_agnpnlts_docoff(a.RN) AS DOCOFF, dm_query.get_agnmedcer_beg(a.RN) AS MEDBEG, dm_query.get_agnmedcer_end(a.RN) AS MEDEND, dm_query.get_agnmedcer_pay(a.RN) AS MEDPAY, DECODE(dm_query.get_agnmedcer_hosp(a.RN), 1, 'Да', 0, 'нет') AS HOSP, dm_query.get_agnmedcer_ser(a.RN) AS MEDSER, dm_query.get_agnmedcer_nom(a.RN) AS MEDNOM, dm_query.get_agnmedcer_dat(a.RN) AS MEDDAT, dm_query.get_agnmedcer_med(a.RN) AS MED, dm_query.get_clnpersvac_vid(a.RN) AS OTPVID, DECODE(dm_query.get_clnpersvac_type(a.RN), 1, 'Дополнительный', 0, 'Основной', 2, 'За свой счет') AS OTPTYPE, dm_query.get_clnpersvac_beg(a.RN) AS OTPBEG, dm_query.get_clnpersvac_end(a.RN) AS OTPEND, dm_query.get_clnpersvac_periodon(a.RN) AS OTPPERON, dm_query.get_clnpersvac_periodoff(a.RN) AS OTPPEROFF, dm_query.get_clnpersvac_planon(a.RN) AS OTPPLANON, dm_query.get_clnpersvac_planoff(a.RN) AS OTPPLANOFF, dm_query.get_clnpersvac_major(a.RN) AS OTPMAJOR, dm_query.get_clnpersvac_minor(a.RN) AS OTPMINOR, dm_query.get_clnpersvac_hol(a.RN) AS OTPHOL, dm_query.get_clnpersvac_trav(a.RN) AS OTPTRAV, dm_query.get_clnpersvac_doctype(a.RN) AS OTPDOCTYPE, dm_query.get_clnpersvac_docwho(a.RN) AS OTPDOCWHO, dm_query.get_clnpersvac_docdat(a.RN) AS OTPDOCDAT, dm_query.get_clnpersvac_docnumb(a.RN) AS OTPDOCNUMB, dm_query.get_clnpersvac_prim(a.RN) AS OTPPRIM, DECODE(dm_query.get_clnpersvac_censel(a.RN), 1, 'Да', 0, 'Нет') AS OTPCENSEL, DECODE(dm_query.get_clnpersvac_hold(a.RN), 1, 'Да', 0, 'Нет') AS OTPHOLD, dm_query.get_clnpersvac_ret(a.RN) AS OTPRET, dm_query.get_clnpersvac_hnote(a.RN) AS OTPHNOTE, dm_query.get_clnpersvac_rettype(a.RN) AS OTPRETTYPE, dm_query.get_clnpersvac_retwho(a.RN) AS OTPRETWHO, dm_query.get_clnpersvac_retdat(a.RN) AS OTPRETDAT, dm_query.get_clnpersvac_retnumb(a.RN) AS OTPRETNUMB, dm_query.get_clnperstrip_beg(a.RN) AS TRIPBEG, dm_query.get_clnperstrip_end(a.RN) AS TRIPEND, dm_query.get_clnperstrip_cel(a.RN) AS TRIPCEL, DECODE(dm_query.get_clnperstrip_cansel(a.RN), 1, 'Да', 0, 'Нет') AS TRIPCANSEL, dm_query.get_clnperstrip_goal(a.RN) AS TRIPGOAL, dm_query.get_clnperstrip_type(a.RN) AS TRIPTYPE, dm_query.get_clnperstrip_who(a.RN) AS TRIPWHO, dm_query.get_clnperstrip_dat(a.RN) AS TRIPDAT, dm_query.get_clnperstrip_numb(a.RN) AS TRIPNUMB, dm_query.get_clnperstrip_adres(a.RN) AS TRIPADRES, dm_query.get_clnperstrip_dopadres(a.RN) AS TRIPDOPADRES, dm_query.get_agnfrntrp_geog(a.RN) AS TRPGEOG, dm_query.get_agnfrntrp_trp(a.RN) AS TRPTRP, dm_query.get_agnfrntrp_beg(a.RN) AS TRPBEG, dm_query.get_agnfrntrp_end2(a.RN) AS TRPEND, dm_query.get_agnaddinf_type(a.RN) AS DDTYPE, dm_query.get_agnaddinf_datbeg(a.RN) AS DDDATBEG, dm_query.get_agnaddinf_datend(a.RN) AS DDDATEND, dm_query.get_agnaddinf_sod(a.RN) AS DDSOD, dm_query.get_sppvd_vid(a.RN) AS SPPVID, dm_query.get_sppvd_numb(a.RN) AS SPPNUMB, dm_query.get_sppvd_docdat(a.RN) AS SPPDOCDAT, dm_query.get_sppvd_docwho(a.RN) AS SPPDOCWHO, dm_query.get_sppvd_vidpro(a.RN) AS SPPVIDPRO, dm_query.get_sppvd_datpro(a.RN) AS SPPDATPRO, dm_query.get_sppvd_opis(a.RN) AS SPPOPIS, dm_query.get_sppvd_stat(a.RN) AS SPPSTAT, dm_query.get_sppvd_mesto(a.RN) AS SPPMESTO, DECODE(dm_query.get_sppvd_vina(a.RN), 1, 'Да', 0, 'Нет') AS SPPVINA, dm_query.get_sppvd_oznak(a.RN) AS SPPOZNAK, DECODE(dm_query.get_sppvd_noznak(a.RN), 1, 'Да', 0, 'Нет') AS SPPNOZNAK, dm_query.get_sppvd_prim(a.RN) AS SPPPRIM, dm_query.get_agnacclev_form(a.RN) AS CCFORM, dm_query.get_agnacclev_s(a.RN) AS CCS, dm_query.get_agnacclev_po(a.RN) AS CCPO, dm_query.get_agnacclev_prim(a.RN) AS CCPRIM, dm_query.get_agnacclev_type(a.RN) AS CCTYPE, dm_query.get_agnacclev_who(a.RN) AS CCWHO, dm_query.get_agnacclev_dat(a.RN) AS CCDAT, dm_query.get_agnacclev_numb(a.RN) AS CCNUMB, dm_query.get_profpsih_dat(a.RN) AS SIHDAT, dm_query.get_profpsih_kat(a.RN) AS SIHKAT, dm_query.get_profpsih_zak(a.RN) AS SIHZAK, dm_query.get_profpsih_note(a.RN) AS SIHNOTE, dm_query.get_profvalid_dat(a.RN) AS LIDDAT, dm_query.get_profvalid_kat(a.RN) AS LIDKAT, dm_query.get_profvalid_who(a.RN) AS LIDWHO, dm_query.get_profvalid_nom(a.RN) AS LIDNOM, dm_query.get_profvalid_note(a.RN) AS LIDNOTE, dm_query.get_agnranks_zvan(a.RN) AS NKSZVAN, dm_query.get_agnranks_vid(a.RN) AS NKSVID, dm_query.get_agnranks_dat(a.RN) AS NKSDAT, dm_query.get_agnranks_ndat(a.RN) AS NKSNDAT, dm_query.get_agnranks_type(a.RN) AS NKSTYPE, dm_query.get_agnranks_who(a.RN) AS NKSWHO, dm_query.get_agnranks_prdat(a.RN) AS NKSPRDAT, dm_query.get_agnranks_nom(a.RN) AS NKSNOM, dm_query.get_agneduc_vidob(a.RN) AS EDVID, dm_query.get_agneduc_prof(a.RN) AS EDPROF, dm_query.get_agneduc_begdat(a.RN) AS EDBEGDAT, dm_query.get_agneduc_enddat(a.RN) AS EDENDDAT, dm_query.get_agneduc_educ(a.RN) AS EDEDUC, dm_query.get_agneduc_form(a.RN) AS EDFORM, DECODE(dm_query.get_agneduc_otm(a.RN), 1, 'Учится', 0, 'Не задана', 2, 'Учится на последнем курсе') AS EDOTM, dm_query.get_agneduc_dtype(a.RN) AS EDDTYPE, dm_query.get_agneduc_ser(a.RN) AS EDSER, dm_query.get_agneduc_nom(a.RN) AS EDNOM, dm_query.get_agneduc_spec(a.RN) AS EDSPEC, dm_query.get_agneduc_kval(a.RN) AS EDKVAL, dm_query.get_agneduc_type(a.RN) AS EDTYPE, dm_query.get_agneduc_who(a.RN) AS EDWHO, dm_query.get_agneduc_ddoc(a.RN) AS EDDDOC, dm_query.get_agneduc_dnom(a.RN) AS EDDNOM, dm_query.get_agneducperv_mesto(a.RN) AS EDPERVMESTO, dm_query.get_agneducperv_tipdoc(a.RN) AS EDPERVTIPDOC, dm_query.get_agneducperv_ser(a.RN) AS EDPERVSER, dm_query.get_agneducperv_nom(a.RN) AS EDPERVNOM, dm_query.get_agneducperv_bdate(a.RN) AS EDPERVBDATE, dm_query.get_agneducperv_edate(a.RN) AS EDPERVEDATE, dm_query.get_agneducperv_type(a.RN) AS EDPERVTYPE, dm_query.get_agneducperv_who(a.RN) AS EDPERVWHO, dm_query.get_agneducperv_dat(a.RN) AS EDPERVDAT, dm_query.get_agneducperv_numb(a.RN) AS EDPERVNUMB, dm_query.get_agneducpov_vidobr(a.RN) AS POVVIDOBR, dm_query.get_agneducpov_uch(a.RN) AS POVUCH, dm_query.get_agneducpov_kurs(a.RN) AS POVKURS, dm_query.get_agneducpov_form(a.RN) AS POVFORM, dm_query.get_agneducpov_fin(a.RN) AS POVFIN, dm_query.get_agneducpov_prof(a.RN) AS POVPROF, dm_query.get_agneducpov_bdate(a.RN) AS POVBDATE, dm_query.get_agneducpov_edate(a.RN) AS POVEDATE, dm_query.get_agneducpov_chas(a.RN) AS POVCHAS, dm_query.get_agneducpov_doctype(a.RN) AS POVDOCTYPE, dm_query.get_agneducpov_docser(a.RN) AS POVDOCSER, dm_query.get_agneducpov_docnom(a.RN) AS POVDOCNOM, dm_query.get_agneducpov_type(a.RN) AS POVTYPE, dm_query.get_agneducpov_who(a.RN) AS POVWHO, dm_query.get_agneducpov_dat(a.RN) AS POVDAT, dm_query.get_agneducpov_nom(a.RN) AS POVNOM, dm_query.get_agneducper_vid(a.RN) AS PERVID, dm_query.get_agneducper_uch(a.RN) AS PERUCH, dm_query.get_agneducper_kurs(a.RN) AS PERKURS, dm_query.get_agneducper_form(a.RN) AS PERFORM, dm_query.get_agneducper_fin(a.RN) AS PERFIN, dm_query.get_agneducper_prof(a.RN) AS PERPROF, dm_query.get_agneducper_bdate(a.RN) AS PERBDATE, dm_query.get_agneducper_edate(a.RN) AS PEREDATE, dm_query.get_agneducper_chas(a.RN) AS PERCHAS, dm_query.get_agneducper_doctype(a.RN) AS PERDOCTYPE, dm_query.get_agneducper_docser(a.RN) AS PERDOCSER, dm_query.get_agneducper_docnom(a.RN) AS PERDOCNOM, dm_query.get_agneducper_type(a.RN) AS PERTYPE, dm_query.get_agneducper_who(a.RN) AS PERWHO, dm_query.get_agneducper_dat(a.RN) AS PERDAT, dm_query.get_agneducper_nom(a.RN) AS PERNOM, dm_query.get_agneducstaj_mesto(a.RN) AS STAJMESTO, dm_query.get_agneducstaj_fin(a.RN) AS STAJFIN, dm_query.get_agneducstaj_bdate(a.RN) AS STAJBDATE, dm_query.get_agneducstaj_edate(a.RN) AS STAJEDATE, dm_query.get_agneducstaj_type(a.RN) AS STAJTYPE, dm_query.get_agneducstaj_who(a.RN) AS STAJWHO, dm_query.get_agneducstaj_dat(a.RN) AS STAJDAT, dm_query.get_agneducstaj_nom(a.RN) AS STAJNOM, dm_query.get_agneducstep_uch(a.RN) AS STEPUCH, dm_query.get_agneducstep_spec(a.RN) AS STEPSPEC, dm_query.get_agneducstep_dat(a.RN) AS STEPDAT, dm_query.get_agneducstep_type(a.RN) AS STEPTYPE, dm_query.get_agneducstep_ser(a.RN) AS STEPSER, dm_query.get_agneducstep_nom(a.RN) AS STEPNOM, dm_query.get_agneducstep_note(a.RN) AS STEPNOTE, dm_query.get_agneduczvan_uch(a.RN) AS ZVANUCH, dm_query.get_agneduczvan_dat(a.RN) AS ZVANDAT, dm_query.get_agneduczvan_type(a.RN) AS ZVANTYPE, dm_query.get_agneduczvan_ser(a.RN) AS ZVANSER, dm_query.get_agneduczvan_nom(a.RN) AS ZVANNOM, dm_query.get_agneduczvan_note(a.RN) AS ZVANNOTE, dm_query.get_agnlangs_in(a.RN) AS LANIN, dm_query.get_agnlangs_step(a.RN) AS LANSTEP, dm_query.get_agnlangs_note(a.RN) AS LANNOTE, dm_query.get_agnatnms_vid(a.RN) AS NMSVID, dm_query.get_agnatnms_step(a.RN) AS NMSSTEP, dm_query.get_docsprops_agnatnms(a.RN) AS DPNM, dm_query.get_agnaddresses_full(a.RN) AS ADRESFULL, dm_query.get_agnaddresses_index(a.RN) AS ADRESINDEX, dm_query.get_agnaddresses_dom(a.RN) AS ADRESDOM, dm_query.get_agnaddresses_blok(a.RN) AS ADRESBLOK, dm_query.get_agnaddresses_flat(a.RN) AS ADRESFLAT, dm_query.get_agnaddresses_note(a.RN) AS ADRESNOTE, dm_query.get_agndocums_type(a.RN) AS CUMSTYPE, dm_query.get_agndocums_ser(a.RN) AS CUMSSER, dm_query.get_agndocums_numb(a.RN) AS CUMSNUMB, dm_query.get_agndocums_vida(a.RN) AS CUMSVIDA, dm_query.get_agndocums_po(a.RN) AS CUMSPO, dm_query.get_agndocums_who(a.RN) AS CUMSWHO, dm_query.get_agndocums_osn(a.RN) AS CUMSOSN, dm_query.get_agndocums_doctype(a.RN) AS CUMSDOCTYPE, dm_query.get_agndocums_docwho(a.RN) AS CUMSDOCWHO, dm_query.get_agndocums_docdat(a.RN) AS CUMSDOCDAT, dm_query.get_agndocums_docnumb(a.RN) AS CUMSDOCNUMB, dm_query.get_agnsprtsk_kval(a.RN) AS TSKKVAL, dm_query.get_agnsprtsk_vid(a.RN) AS TSKVID, dm_query.get_agnsprtsk_god(a.RN) AS TSKGOD, dm_query.get_agnlbractivityprs_bdate(a.RN) AS TRUDBDATE, dm_query.get_agnlbractivityprs_edate(a.RN) AS TRUDEDATE, dm_query.get_agnlbractivityprs_gos(a.RN) AS TRUDGOS, dm_query.get_agnlbractivityprs_work(a.RN) AS TRUDWORK, dm_query.get_agnlbractivityprs_dolj(a.RN) AS TRUDDOLJ, dm_query.get_agnlbractivityprs_prich(a.RN) AS TRUDPRICH, dm_query.get_agnlbractivityprs_type(a.RN) AS TRUDTYPE, dm_query.get_agnlbractivityprs_who(a.RN) AS TRUDWHO, dm_query.get_agnlbractivityprs_date(a.RN) AS TRUDDATE, dm_query.get_agnlbractivityprs_numb(a.RN) AS TRUDNUMB, dm_query.get_agnlbractivityprs_note(a.RN) AS TRUDNOTE, dm_query.get_agnsvcprd_sluj(a.RN) AS SVCSLUJ, dm_query.get_agnsvcprd_post(a.RN) AS SVCPOST, dm_query.get_agnsvcprd_post2(a.RN) AS SVCPOST2, dm_query.get_agnsvcprd_vidpost(a.RN) AS SVCVIDPOST, dm_query.get_agnsvcprd_kat(a.RN) AS SVCKAT, dm_query.get_agnsvcprd_primtype(a.RN) AS SVCPRIMTYPE, dm_query.get_agnsvcprd_primwho(a.RN) AS SVCPRIMWHO, dm_query.get_agnsvcprd_primdate(a.RN) AS SVCPRIMDATE, dm_query.get_agnsvcprd_primnumb(a.RN) AS SVCPRIMNUMB, dm_query.get_agnsvcprd_ispdat(a.RN) AS SVCISPDAT, dm_query.get_agnsvcprd_ispedat(a.RN) AS SVCISPEDAT, dm_query.get_agnsvcprd_ispzak(a.RN) AS SVCISPZAK, dm_query.get_agnsvcprd_isptype(a.RN) AS SVCISPTYPE, dm_query.get_agnsvcprd_ispwho(a.RN) AS SVCISPWHO, dm_query.get_agnsvcprd_ispdate(a.RN) AS SVCISPDATE, dm_query.get_agnsvcprd_ispnumb(a.RN) AS SVCISPNUMB, dm_query.get_agnsvcmov_bdate(a.RN) AS MOVBDATE, dm_query.get_agnsvcmov_edate(a.RN) AS MOVEDATE, dm_query.get_agnsvcmov_pervid(a.RN) AS MOVPERVID, dm_query.get_agnsvcmov_perprich(a.RN) AS MOVPERPRICH, dm_query.get_agnsvcmov_permot(a.RN) AS MOVPERMOT, dm_query.get_agnsvcmov_mesto(a.RN) AS MOVMESTO, dm_query.get_agnsvcmov_org(a.RN) AS MOVORG, dm_query.get_agnsvcmov_block(a.RN) AS MOVBLOCK, dm_query.get_agnsvcmov_group(a.RN) AS MOVGROUP, dm_query.get_agnsvcmov_kod(a.RN) AS MOVKOD, dm_query.get_agnsvcmov_dolj(a.RN) AS MOVDOLJ, dm_query.get_agnsvcmov_fdolj(a.RN) AS MOVFDOLJ, dm_query.get_agnsvcmov_type(a.RN) AS MOVTYPE, dm_query.get_agnsvcmov_who(a.RN) AS MOVWHO, dm_query.get_agnsvcmov_date(a.RN) AS MOVDATE, dm_query.get_agnsvcmov_numb(a.RN) AS MOVNUMB, dm_query.get_agnsvcmov_prekprich(a.RN) AS MOVPRIKPRICH, dm_query.get_agnsvcmov_prektype(a.RN) AS MOVPRIKTYPE, dm_query.get_agnsvcmov_prekwho(a.RN) AS MOVPRIKWHO, dm_query.get_agnsvcmov_prekdate(a.RN) AS MOVPRIKDATE, dm_query.get_agnsvcmov_preknumb(a.RN) AS MOVPRIKNUMB, dm_query.get_agnsvcmov_note(a.RN) AS MOVNOTE, dm_query.get_max_execution(a.RN) AS EXEC, dm_query.get_max_shortexecution(a.RN) AS SHORTEXEC, dm_query.get_max_dateexecution(a.RN) AS DATEEXECUTION, dm_query.get_clnpsdep_code(a.RN) AS CLNCODE, dm_query.get_clnpsdep_zvan2(a.RN) AS CLNZVAN, dm_query.get_clnpsdep_doldat(a.RN) AS CLNDOLDAT, dm_query.get_clnpsdep_type(a.RN) AS CLNTYPE, dm_query.get_clnpsdep_who(a.RN) AS CLNWHO, dm_query.get_clnpsdep_date(a.RN) AS CLNDATE, dm_query.get_clnpsdep_numb(a.RN) AS CLNNUMB, dm_query.get_clnpsdep_kod(a.RN) AS CLNKOD, dm_query.get_clnpsdep_kad(a.RN) AS CLNKAD, dm_query.get_clnpsdep_group(a.RN) AS CLNGROUP FROM UK_PARUS.AGNLIST a WHERE a.AGNTYPE = 1 AND a.RN IN(SELECT pp.AGNLIST FROM UK_PARUS.PREMPLFLS pp)) MAIN_AGNRN ON MAIN_AGNRN.RN = MAIN.AGN_RN WHERE  " + Peroid_Parametr + " AND (MAIN_AGNRN.SVCKAT != 'Внутри т/о' and MAIN_AGNRN.SVCKAT != 'из др.там.органа') AND EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP WHERE UP.CATALOG = MAIN.CRN)";      // Создать объект Command.
            }
            if (arrTM_depart != 0)
            {
                Sotrudnik_Prinyat = "SELECT MAIN_crn.NAME AS Name_TO,   MAIN.MOVORG AS Mnemo_TO,          MAIN_crn.PARENT AS Verhnii_Level,          MAIN.INICIALI AS FIO,           MAIN.SEX AS Pol,          MAIN.AGE AS Vozrast,           MAIN.EXEC AS Doljnost,       MAIN.CLNGROUP AS Gruppa_Doljnost,           MAIN.PRDTKN AS Status_Slujb,           MAIN.SVCSLUJ AS Vid_Sluj,           TO_CHAR(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'dd.MM.yyyy') AS Data_Naznachenia, MAIN.SHORTEXEC AS Mnemo_Dolj, MAIN.CLNKAD AS Kat_Dolj,           MAIN.EDVID AS Obrazov, MAIN.EDPROF AS Vid_Obrazov,  MAIN.EDSPEC AS Specialnost,TO_CHAR(MAIN.LAST, 'dd.MM.yyyy') AS Date_Priema, MAIN.SVCKAT AS Otkyda FROM(SELECT a.AGNFAMILYNAME || ' ' || a.AGNFIRSTNAME || ' ' || a.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(a.RN) AS INICIALI, TRUNC(MONTHS_BETWEEN(SYSDATE, a.AGNBURN) / 12) AS AGE, a.CRN AS crn,               dm_query.get_premplsts(a.RN) AS PRDTKN, DECODE(a.SEX, 1, 'муж', 2, 'жен') AS SEX, a.PRFMLSTS AS SEM,                dm_query.get_latter_date(a.RN) AS LATTER, dm_query.get_min_datebg(a.RN) AS LAST, dm_query.get_agneduc_vidob(a.RN) AS EDVID, dm_query.get_agneduc_prof(a.RN) AS EDPROF, dm_query.get_agneduc_spec(a.RN) AS EDSPEC, dm_query.get_agnsvcprd_sluj(a.RN) AS SVCSLUJ, dm_query.get_agnsvcprd_kat(a.RN) AS SVCKAT, dm_query.get_agnsvcmov_org(a.RN) AS MOVORG, dm_query.get_max_execution(a.RN) AS EXEC, dm_query.get_max_shortexecution(a.RN) AS SHORTEXEC, dm_query.get_max_dateexecution(a.RN) AS DATEEXECUTION, dm_query.get_clnpsdep_kad(a.RN) AS CLNKAD, dm_query.get_clnpsdep_group(a.RN) AS CLNGROUP       FROM UK_PARUS.AGNLIST a      WHERE a.AGNTYPE = 1        AND a.RN IN(SELECT pp.AGNLIST FROM UK_PARUS.PREMPLFLS pp)) MAIN LEFT OUTER JOIN(SELECT T.RN AS RN, T.NAME AS NAME, DECODE(T.IS_ROOT, 0, 'нет', 1, 'да') AS IS_ROOT,                            (SELECT A.NAME                                FROM UK_PARUS.ACATALOG A                                WHERE LEVEL = 2                              CONNECT BY PRIOR a.crn = a.rn                              START WITH a.rn = t.rn) AS PARENT        FROM UK_PARUS.ACATALOG T) MAIN_crn ON MAIN_crn.RN = MAIN.crn  WHERE MAIN.MOVORG = '" + arrTM_mnemo + "' and   " + Peroid_Parametr + "AND MAIN.SVCKAT <> 'Внутри т/о'    AND MAIN.SVCKAT <> 'из др.там.органа'    AND MAIN.SVCSLUJ <> 'Работник'    AND EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP        WHERE UP.CATALOG = MAIN.crn)  ORDER BY MAIN.LATTER";
                YvolSQL = "SELECT SubQuery.DATA_YVOLN AS Date_Yvol, SubQuery.FIO AS FIO, SubQuery.KATALOG AS Katalog, SUBQUERY.TORG AS TORG, SUBQUERY.AGENT_RN AS Agen_RN, SUBQUERY.VOZRAST AS Vozrast, SUBQUERY.SEX AS POL, SUBQUERY.DATA_NACH AS Date_Nach, SUBQUERY.DATA_FIN AS Date_Fin, SUBQUERY.VID_SLUJB AS Vid_Slujb, SUBQUERY.VID_OBRAZ AS Vid_Obraz, SubQuery_1.NAME as Motiv, SUBQUERY.MNEMO_DOLJ AS Mnemo_Dolj, SUBQUERY.GRUPPA_DOLJNOST AS Gruppa_Dolj, SUBQUERY.KAT_DOLJ AS Kat_Dolj FROM (SELECT MAIN.MOVEDATE AS Data_Yvoln, MAIN.INICIALI AS FIO, MAIN_crn.NAME AS Katalog, MAIN_crn.PARENT AS TORG, MAIN.RN AS AGENT_RN, MAIN.AGE AS Vozrast, MAIN.SVCPOST AS Data_Nach, MAIN.SVCPOST2 AS Data_Fin, MAIN.SVCSLUJ AS Vid_Slujb, MAIN.SEX AS SEX , MAIN.EDVID AS Vid_Obraz, MAIN.SHORTEXEC AS Mnemo_Dolj, MAIN.CLNGROUP AS Gruppa_Doljnost, MAIN.CLNKAD AS Kat_Dolj FROM (SELECT a.AGNFAMILYNAME AS agnfamilyname, a.AGNFIRSTNAME AS agnfirstname, a.AGNLASTNAME AS agnlastname, a.AGNFAMILYNAME || ' ' || a.AGNFIRSTNAME || ' ' || a.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(a.RN) AS INICIALI, a.RN AS RN, a.AGNBURN AS agnburn, TRUNC(MONTHS_BETWEEN(SYSDATE, a.AGNBURN) / 12) AS AGE, dm_query.get_prem_ln(a.RN) AS licnumb, dm_query.get_depart_uprav_ca(a.RN) AS UPRAV, dm_query.get_depart_upravshort_ca(a.RN) AS UPRAV_SHORT, dm_query.get_discharge1(a.RN) AS discharge, dm_query.get_depart(a.RN) AS DEPART, a.AGNABBR AS AGNCODE, a.CRN AS crn, dm_query.get_premplsts(a.RN) AS PRDTKN, DECODE(a.SEX, 1, 'муж', 2, 'жен') AS SEX, a.PRFMLSTS AS SEM, a.PENSION_NBR AS NBR, a.AGNIDNUMB AS ID, a.PHONE AS PHONE, a.PHONE2 AS PHONE2, a.TELEX AS TELEX, DM_QUERY.GET_STAJ(a.RN) AS ssta, dm_query.GET_STAJ_KAD2(a.RN) AS dddss, dm_query.get_agnsvcprd(a.RN) AS TORG, dm_query.get_first_date(a.RN) AS FIRST, dm_query.get_latter_date(a.RN) AS LATTER, dm_query.get_min_datebg(a.RN) AS LAST, dm_query.get_isp_contr(a.RN) AS CONTR, dm_query.get_isp_contrnom(a.RN) AS CONNOM, dm_query.get_isp_contrvid(a.RN) AS CONVID, dm_query.get_isp_contrend(a.RN) AS DATEND, dm_query.get_isp_controsn(a.RN) AS OSN, dm_query.get_isp_contrdopnom(a.RN) AS DOPNOM, dm_query.get_isp_contrdopdat(a.RN) AS DOPDAT, dm_query.get_isp_contrdopopis(a.RN) AS DOPOPIS, dm_query.get_isp_contrdopnom40(a.RN) AS DOPNOM40, dm_query.get_isp_contrdopdat40(a.RN) AS DOPDAT40, dm_query.get_isp_contrdopopis40(a.RN) AS DOPOPIS40, dm_query.get_pris(a.RN) AS PRIS, dm_query.get_provdat(a.RN) AS PROV, dm_query.get_provnom(a.RN) AS PROVNOM, dm_query.get_dak(a.RN) AS DAK, dm_query.get_attdat(a.RN) AS ATTDAT, dm_query.get_attns_date(a.RN) AS ATTNSDATE, dm_query.get_attdatnext(a.RN) AS ATTDATNE, dm_query.get_attosn(a.RN) AS ATTOSN, dm_query.get_attviv(a.RN) AS ATTVIV, dm_query.get_attdoctype(a.RN) AS ATTDOCTYPE, dm_query.get_attwho(a.RN) AS ATTWHO, dm_query.get_attdocdat(a.RN) AS ATTDOCDAT, dm_query.get_attdocnom(a.RN) AS ATTDOCNOM, DECODE(dm_query.get_attrlz(a.RN), 1, 'Да', 0, 'нет') AS ATTRLZ, dm_query.get_attdoctype_rlz(a.RN) AS ATTRLZTYPE, dm_query.get_attdocwho_rlz(a.RN) AS ATTRLZWHO, dm_query.get_attdocdat_rlz(a.RN) AS ATTRLZDAT, dm_query.get_attdocnom_rlz(a.RN) AS ATTRLZNOM, dm_query.get_agnrwrd_dat(a.RN) AS WRDDAT, dm_query.get_agnrwrd_who(a.RN) AS WRDWHO, dm_query.get_agnrwrd_numb(a.RN) AS WRDNUMB, dm_query.get_agnrwrd_type(a.RN) AS WRDTYPE, dm_query.get_agnrwrd_typ(a.RN) AS WRDTYP, dm_query.get_agnrwrd_vid(a.RN) AS WRDVID, dm_query.get_agnrwrd_reas(a.RN) AS WRDREAS, dm_query.get_agnpnlts_knd(a.RN) AS KND, dm_query.get_calc_length(a.RN) AS CALCLEN, dm_query.get_agnpnlts_fault(a.RN) AS FOULT, dm_query.get_agnpnlts_nlk(a.RN) AS NLK, dm_query.get_agnpnlts_css(a.RN) AS CSS, dm_query.get_agnpnlts_ini(a.RN) AS INI, dm_query.get_agnpnlts_pnl(a.RN) AS PNL, dm_query.get_agnpnlts_typeon(a.RN) AS TYPEON, dm_query.get_agnpnlts_whonal(a.RN) AS WHONAL, dm_query.get_agnpnlts_docon(a.RN) AS DOCON, dm_query.get_agnpnlts_docon2(a.RN) AS DOCON2, dm_query.get_agnpnlts_cntrldat(a.RN) AS CNTR, dm_query.get_agnpnlts_appl(a.RN) AS APPL, dm_query.get_agnpnlts_remov(a.RN) AS REMOV, dm_query.get_agnpnlts_osn(a.RN) AS OSN1, dm_query.get_agnpnlts_typeoff(a.RN) AS TYPEOFF, dm_query.get_agnpnlts_whosnal(a.RN) AS WHOSNAL, dm_query.get_agnpnlts_datoff(a.RN) AS DATOFF, dm_query.get_agnpnlts_docoff(a.RN) AS DOCOFF, dm_query.get_agnmedcer_beg(a.RN) AS MEDBEG, dm_query.get_agnmedcer_end(a.RN) AS MEDEND, dm_query.get_agnmedcer_pay(a.RN) AS MEDPAY, DECODE(dm_query.get_agnmedcer_hosp(a.RN), 1, 'Да', 0, 'нет') AS HOSP, dm_query.get_agnmedcer_ser(a.RN) AS MEDSER, dm_query.get_agnmedcer_nom(a.RN) AS MEDNOM, dm_query.get_agnmedcer_dat(a.RN) AS MEDDAT, dm_query.get_agnmedcer_med(a.RN) AS MED, dm_query.get_clnpersvac_vid(a.RN) AS OTPVID, DECODE(dm_query.get_clnpersvac_type(a.RN), 1, 'Дополнительный', 0, 'Основной', 2, 'За свой счет') AS OTPTYPE, dm_query.get_clnpersvac_beg(a.RN) AS OTPBEG, dm_query.get_clnpersvac_end(a.RN) AS OTPEND, dm_query.get_clnpersvac_periodon(a.RN) AS OTPPERON, dm_query.get_clnpersvac_periodoff(a.RN) AS OTPPEROFF, dm_query.get_clnpersvac_planon(a.RN) AS OTPPLANON, dm_query.get_clnpersvac_planoff(a.RN) AS OTPPLANOFF, dm_query.get_clnpersvac_major(a.RN) AS OTPMAJOR, dm_query.get_clnpersvac_minor(a.RN) AS OTPMINOR, dm_query.get_clnpersvac_hol(a.RN) AS OTPHOL, dm_query.get_clnpersvac_trav(a.RN) AS OTPTRAV, dm_query.get_clnpersvac_doctype(a.RN) AS OTPDOCTYPE, dm_query.get_clnpersvac_docwho(a.RN) AS OTPDOCWHO, dm_query.get_clnpersvac_docdat(a.RN) AS OTPDOCDAT, dm_query.get_clnpersvac_docnumb(a.RN) AS OTPDOCNUMB, dm_query.get_clnpersvac_prim(a.RN) AS OTPPRIM, DECODE(dm_query.get_clnpersvac_censel(a.RN), 1, 'Да', 0, 'Нет') AS OTPCENSEL, DECODE(dm_query.get_clnpersvac_hold(a.RN), 1, 'Да', 0, 'Нет') AS OTPHOLD, dm_query.get_clnpersvac_ret(a.RN) AS OTPRET, dm_query.get_clnpersvac_hnote(a.RN) AS OTPHNOTE, dm_query.get_clnpersvac_rettype(a.RN) AS OTPRETTYPE, dm_query.get_clnpersvac_retwho(a.RN) AS OTPRETWHO, dm_query.get_clnpersvac_retdat(a.RN) AS OTPRETDAT, dm_query.get_clnpersvac_retnumb(a.RN) AS OTPRETNUMB, dm_query.get_clnperstrip_beg(a.RN) AS TRIPBEG, dm_query.get_clnperstrip_end(a.RN) AS TRIPEND, dm_query.get_clnperstrip_cel(a.RN) AS TRIPCEL, DECODE(dm_query.get_clnperstrip_cansel(a.RN), 1, 'Да', 0, 'Нет') AS TRIPCANSEL, dm_query.get_clnperstrip_goal(a.RN) AS TRIPGOAL, dm_query.get_clnperstrip_type(a.RN) AS TRIPTYPE, dm_query.get_clnperstrip_who(a.RN) AS TRIPWHO, dm_query.get_clnperstrip_dat(a.RN) AS TRIPDAT, dm_query.get_clnperstrip_numb(a.RN) AS TRIPNUMB, dm_query.get_clnperstrip_adres(a.RN) AS TRIPADRES, dm_query.get_clnperstrip_dopadres(a.RN) AS TRIPDOPADRES, dm_query.get_agnfrntrp_geog(a.RN) AS TRPGEOG, dm_query.get_agnfrntrp_trp(a.RN) AS TRPTRP, dm_query.get_agnfrntrp_beg(a.RN) AS TRPBEG, dm_query.get_agnfrntrp_end2(a.RN) AS TRPEND, dm_query.get_agnaddinf_type(a.RN) AS DDTYPE, dm_query.get_agnaddinf_datbeg(a.RN) AS DDDATBEG, dm_query.get_agnaddinf_datend(a.RN) AS DDDATEND, dm_query.get_agnaddinf_sod(a.RN) AS DDSOD, dm_query.get_sppvd_vid(a.RN) AS SPPVID, dm_query.get_sppvd_numb(a.RN) AS SPPNUMB, dm_query.get_sppvd_docdat(a.RN) AS SPPDOCDAT, dm_query.get_sppvd_docwho(a.RN) AS SPPDOCWHO, dm_query.get_sppvd_vidpro(a.RN) AS SPPVIDPRO, dm_query.get_sppvd_datpro(a.RN) AS SPPDATPRO, dm_query.get_sppvd_opis(a.RN) AS SPPOPIS, dm_query.get_sppvd_stat(a.RN) AS SPPSTAT, dm_query.get_sppvd_mesto(a.RN) AS SPPMESTO, DECODE(dm_query.get_sppvd_vina(a.RN), 1, 'Да', 0, 'Нет') AS SPPVINA, dm_query.get_sppvd_oznak(a.RN) AS SPPOZNAK, DECODE(dm_query.get_sppvd_noznak(a.RN), 1, 'Да', 0, 'Нет') AS SPPNOZNAK, dm_query.get_sppvd_prim(a.RN) AS SPPPRIM, dm_query.get_agnacclev_form(a.RN) AS CCFORM, dm_query.get_agnacclev_s(a.RN) AS CCS, dm_query.get_agnacclev_po(a.RN) AS CCPO, dm_query.get_agnacclev_prim(a.RN) AS CCPRIM, dm_query.get_agnacclev_type(a.RN) AS CCTYPE, dm_query.get_agnacclev_who(a.RN) AS CCWHO, dm_query.get_agnacclev_dat(a.RN) AS CCDAT, dm_query.get_agnacclev_numb(a.RN) AS CCNUMB, dm_query.get_profpsih_dat(a.RN) AS SIHDAT, dm_query.get_profpsih_kat(a.RN) AS SIHKAT, dm_query.get_profpsih_zak(a.RN) AS SIHZAK, dm_query.get_profpsih_note(a.RN) AS SIHNOTE, dm_query.get_profvalid_dat(a.RN) AS LIDDAT, dm_query.get_profvalid_kat(a.RN) AS LIDKAT, dm_query.get_profvalid_who(a.RN) AS LIDWHO, dm_query.get_profvalid_nom(a.RN) AS LIDNOM, dm_query.get_profvalid_note(a.RN) AS LIDNOTE, dm_query.get_agnranks_zvan(a.RN) AS NKSZVAN, dm_query.get_agnranks_vid(a.RN) AS NKSVID, dm_query.get_agnranks_dat(a.RN) AS NKSDAT, dm_query.get_agnranks_ndat(a.RN) AS NKSNDAT, dm_query.get_agnranks_type(a.RN) AS NKSTYPE, dm_query.get_agnranks_who(a.RN) AS NKSWHO, dm_query.get_agnranks_prdat(a.RN) AS NKSPRDAT, dm_query.get_agnranks_nom(a.RN) AS NKSNOM, dm_query.get_agneduc_vidob(a.RN) AS EDVID, dm_query.get_agneduc_prof(a.RN) AS EDPROF, dm_query.get_agneduc_begdat(a.RN) AS EDBEGDAT, dm_query.get_agneduc_enddat(a.RN) AS EDENDDAT, dm_query.get_agneduc_educ(a.RN) AS EDEDUC, dm_query.get_agneduc_form(a.RN) AS EDFORM, DECODE(dm_query.get_agneduc_otm(a.RN), 1, 'Учится', 0, 'Не задана', 2, 'Учится на последнем курсе') AS EDOTM, dm_query.get_agneduc_dtype(a.RN) AS EDDTYPE, dm_query.get_agneduc_ser(a.RN) AS EDSER, dm_query.get_agneduc_nom(a.RN) AS EDNOM, dm_query.get_agneduc_spec(a.RN) AS EDSPEC, dm_query.get_agneduc_kval(a.RN) AS EDKVAL, dm_query.get_agneduc_type(a.RN) AS EDTYPE, dm_query.get_agneduc_who(a.RN) AS EDWHO, dm_query.get_agneduc_ddoc(a.RN) AS EDDDOC, dm_query.get_agneduc_dnom(a.RN) AS EDDNOM, dm_query.get_agneducperv_mesto(a.RN) AS EDPERVMESTO, dm_query.get_agneducperv_tipdoc(a.RN) AS EDPERVTIPDOC, dm_query.get_agneducperv_ser(a.RN) AS EDPERVSER, dm_query.get_agneducperv_nom(a.RN) AS EDPERVNOM, dm_query.get_agneducperv_bdate(a.RN) AS EDPERVBDATE, dm_query.get_agneducperv_edate(a.RN) AS EDPERVEDATE, dm_query.get_agneducperv_type(a.RN) AS EDPERVTYPE, dm_query.get_agneducperv_who(a.RN) AS EDPERVWHO, dm_query.get_agneducperv_dat(a.RN) AS EDPERVDAT, dm_query.get_agneducperv_numb(a.RN) AS EDPERVNUMB, dm_query.get_agneducpov_vidobr(a.RN) AS POVVIDOBR, dm_query.get_agneducpov_uch(a.RN) AS POVUCH, dm_query.get_agneducpov_kurs(a.RN) AS POVKURS, dm_query.get_agneducpov_form(a.RN) AS POVFORM, dm_query.get_agneducpov_fin(a.RN) AS POVFIN, dm_query.get_agneducpov_prof(a.RN) AS POVPROF, dm_query.get_agneducpov_bdate(a.RN) AS POVBDATE, dm_query.get_agneducpov_edate(a.RN) AS POVEDATE, dm_query.get_agneducpov_chas(a.RN) AS POVCHAS, dm_query.get_agneducpov_doctype(a.RN) AS POVDOCTYPE, dm_query.get_agneducpov_docser(a.RN) AS POVDOCSER, dm_query.get_agneducpov_docnom(a.RN) AS POVDOCNOM, dm_query.get_agneducpov_type(a.RN) AS POVTYPE, dm_query.get_agneducpov_who(a.RN) AS POVWHO, dm_query.get_agneducpov_dat(a.RN) AS POVDAT, dm_query.get_agneducpov_nom(a.RN) AS POVNOM, dm_query.get_agneducper_vid(a.RN) AS PERVID, dm_query.get_agneducper_uch(a.RN) AS PERUCH, dm_query.get_agneducper_kurs(a.RN) AS PERKURS, dm_query.get_agneducper_form(a.RN) AS PERFORM, dm_query.get_agneducper_fin(a.RN) AS PERFIN, dm_query.get_agneducper_prof(a.RN) AS PERPROF, dm_query.get_agneducper_bdate(a.RN) AS PERBDATE, dm_query.get_agneducper_edate(a.RN) AS PEREDATE, dm_query.get_agneducper_chas(a.RN) AS PERCHAS, dm_query.get_agneducper_doctype(a.RN) AS PERDOCTYPE, dm_query.get_agneducper_docser(a.RN) AS PERDOCSER, dm_query.get_agneducper_docnom(a.RN) AS PERDOCNOM, dm_query.get_agneducper_type(a.RN) AS PERTYPE, dm_query.get_agneducper_who(a.RN) AS PERWHO, dm_query.get_agneducper_dat(a.RN) AS PERDAT, dm_query.get_agneducper_nom(a.RN) AS PERNOM, dm_query.get_agneducstaj_mesto(a.RN) AS STAJMESTO, dm_query.get_agneducstaj_fin(a.RN) AS STAJFIN, dm_query.get_agneducstaj_bdate(a.RN) AS STAJBDATE, dm_query.get_agneducstaj_edate(a.RN) AS STAJEDATE, dm_query.get_agneducstaj_type(a.RN) AS STAJTYPE, dm_query.get_agneducstaj_who(a.RN) AS STAJWHO, dm_query.get_agneducstaj_dat(a.RN) AS STAJDAT, dm_query.get_agneducstaj_nom(a.RN) AS STAJNOM, dm_query.get_agneducstep_uch(a.RN) AS STEPUCH, dm_query.get_agneducstep_spec(a.RN) AS STEPSPEC, dm_query.get_agneducstep_dat(a.RN) AS STEPDAT, dm_query.get_agneducstep_type(a.RN) AS STEPTYPE, dm_query.get_agneducstep_ser(a.RN) AS STEPSER, dm_query.get_agneducstep_nom(a.RN) AS STEPNOM, dm_query.get_agneducstep_note(a.RN) AS STEPNOTE, dm_query.get_agneduczvan_uch(a.RN) AS ZVANUCH, dm_query.get_agneduczvan_dat(a.RN) AS ZVANDAT, dm_query.get_agneduczvan_type(a.RN) AS ZVANTYPE, dm_query.get_agneduczvan_ser(a.RN) AS ZVANSER, dm_query.get_agneduczvan_nom(a.RN) AS ZVANNOM, dm_query.get_agneduczvan_note(a.RN) AS ZVANNOTE, dm_query.get_agnlangs_in(a.RN) AS LANIN, dm_query.get_agnlangs_step(a.RN) AS LANSTEP, dm_query.get_agnlangs_note(a.RN) AS LANNOTE, dm_query.get_agnatnms_vid(a.RN) AS NMSVID, dm_query.get_agnatnms_step(a.RN) AS NMSSTEP, dm_query.get_docsprops_agnatnms(a.RN) AS DPNM, dm_query.get_agnaddresses_full(a.RN) AS ADRESFULL, dm_query.get_agnaddresses_index(a.RN) AS ADRESINDEX, dm_query.get_agnaddresses_dom(a.RN) AS ADRESDOM, dm_query.get_agnaddresses_blok(a.RN) AS ADRESBLOK, dm_query.get_agnaddresses_flat(a.RN) AS ADRESFLAT, dm_query.get_agnaddresses_note(a.RN) AS ADRESNOTE, dm_query.get_agndocums_type(a.RN) AS CUMSTYPE, dm_query.get_agndocums_ser(a.RN) AS CUMSSER, dm_query.get_agndocums_numb(a.RN) AS CUMSNUMB, dm_query.get_agndocums_vida(a.RN) AS CUMSVIDA, dm_query.get_agndocums_po(a.RN) AS CUMSPO, dm_query.get_agndocums_who(a.RN) AS CUMSWHO, dm_query.get_agndocums_osn(a.RN) AS CUMSOSN, dm_query.get_agndocums_doctype(a.RN) AS CUMSDOCTYPE, dm_query.get_agndocums_docwho(a.RN) AS CUMSDOCWHO, dm_query.get_agndocums_docdat(a.RN) AS CUMSDOCDAT, dm_query.get_agndocums_docnumb(a.RN) AS CUMSDOCNUMB, dm_query.get_agnsprtsk_kval(a.RN) AS TSKKVAL, dm_query.get_agnsprtsk_vid(a.RN) AS TSKVID, dm_query.get_agnsprtsk_god(a.RN) AS TSKGOD, dm_query.get_agnlbractivityprs_bdate(a.RN) AS TRUDBDATE, dm_query.get_agnlbractivityprs_edate(a.RN) AS TRUDEDATE, dm_query.get_agnlbractivityprs_gos(a.RN) AS TRUDGOS, dm_query.get_agnlbractivityprs_work(a.RN) AS TRUDWORK, dm_query.get_agnlbractivityprs_dolj(a.RN) AS TRUDDOLJ, dm_query.get_agnlbractivityprs_prich(a.RN) AS TRUDPRICH, dm_query.get_agnlbractivityprs_type(a.RN) AS TRUDTYPE, dm_query.get_agnlbractivityprs_who(a.RN) AS TRUDWHO, dm_query.get_agnlbractivityprs_date(a.RN) AS TRUDDATE, dm_query.get_agnlbractivityprs_numb(a.RN) AS TRUDNUMB, dm_query.get_agnlbractivityprs_note(a.RN) AS TRUDNOTE, dm_query.get_agnsvcprd_sluj(a.RN) AS SVCSLUJ, dm_query.get_agnsvcprd_post(a.RN) AS SVCPOST, dm_query.get_agnsvcprd_post2(a.RN) AS SVCPOST2, dm_query.get_agnsvcprd_vidpost(a.RN) AS SVCVIDPOST, dm_query.get_agnsvcprd_kat(a.RN) AS SVCKAT, dm_query.get_agnsvcprd_primtype(a.RN) AS SVCPRIMTYPE, dm_query.get_agnsvcprd_primwho(a.RN) AS SVCPRIMWHO, dm_query.get_agnsvcprd_primdate(a.RN) AS SVCPRIMDATE, dm_query.get_agnsvcprd_primnumb(a.RN) AS SVCPRIMNUMB, dm_query.get_agnsvcprd_ispdat(a.RN) AS SVCISPDAT, dm_query.get_agnsvcprd_ispedat(a.RN) AS SVCISPEDAT, dm_query.get_agnsvcprd_ispzak(a.RN) AS SVCISPZAK, dm_query.get_agnsvcprd_isptype(a.RN) AS SVCISPTYPE, dm_query.get_agnsvcprd_ispwho(a.RN) AS SVCISPWHO, dm_query.get_agnsvcprd_ispdate(a.RN) AS SVCISPDATE, dm_query.get_agnsvcprd_ispnumb(a.RN) AS SVCISPNUMB, dm_query.get_agnsvcmov_bdate(a.RN) AS MOVBDATE, dm_query.get_agnsvcmov_edate(a.RN) AS MOVEDATE, dm_query.get_agnsvcmov_pervid(a.RN) AS MOVPERVID, dm_query.get_agnsvcmov_perprich(a.RN) AS MOVPERPRICH, dm_query.get_agnsvcmov_permot(a.RN) AS MOVPERMOT, dm_query.get_agnsvcmov_mesto(a.RN) AS MOVMESTO, dm_query.get_agnsvcmov_org(a.RN) AS MOVORG, dm_query.get_agnsvcmov_block(a.RN) AS MOVBLOCK, dm_query.get_agnsvcmov_group(a.RN) AS MOVGROUP, dm_query.get_agnsvcmov_kod(a.RN) AS MOVKOD, dm_query.get_agnsvcmov_dolj(a.RN) AS MOVDOLJ, dm_query.get_agnsvcmov_fdolj(a.RN) AS MOVFDOLJ, dm_query.get_agnsvcmov_type(a.RN) AS MOVTYPE, dm_query.get_agnsvcmov_who(a.RN) AS MOVWHO, dm_query.get_agnsvcmov_date(a.RN) AS MOVDATE, dm_query.get_agnsvcmov_numb(a.RN) AS MOVNUMB, dm_query.get_agnsvcmov_prekprich(a.RN) AS MOVPRIKPRICH, dm_query.get_agnsvcmov_prektype(a.RN) AS MOVPRIKTYPE, dm_query.get_agnsvcmov_prekwho(a.RN) AS MOVPRIKWHO, dm_query.get_agnsvcmov_prekdate(a.RN) AS MOVPRIKDATE, dm_query.get_agnsvcmov_preknumb(a.RN) AS MOVPRIKNUMB, dm_query.get_agnsvcmov_note(a.RN) AS MOVNOTE, dm_query.get_max_execution(a.RN) AS EXEC, dm_query.get_max_shortexecution(a.RN) AS SHORTEXEC, dm_query.get_max_dateexecution(a.RN) AS DATEEXECUTION, dm_query.get_clnpsdep_code(a.RN) AS CLNCODE, dm_query.get_clnpsdep_zvan2(a.RN) AS CLNZVAN, dm_query.get_clnpsdep_doldat(a.RN) AS CLNDOLDAT, dm_query.get_clnpsdep_type(a.RN) AS CLNTYPE, dm_query.get_clnpsdep_who(a.RN) AS CLNWHO, dm_query.get_clnpsdep_date(a.RN) AS CLNDATE, dm_query.get_clnpsdep_numb(a.RN) AS CLNNUMB, dm_query.get_clnpsdep_kod(a.RN) AS CLNKOD, dm_query.get_clnpsdep_kad(a.RN) AS CLNKAD, dm_query.get_clnpsdep_group(a.RN) AS CLNGROUP FROM UK_PARUS.AGNLIST a WHERE a.AGNTYPE = 1 AND a.RN IN (SELECT pp.AGNLIST FROM UK_PARUS.PREMPLFLS pp)) MAIN LEFT OUTER JOIN (SELECT T.RN AS RN, T.NAME AS NAME, DECODE(T.IS_ROOT, 0, 'нет', 1, 'да') AS IS_ROOT, (SELECT A.NAME FROM UK_PARUS.ACATALOG A WHERE LEVEL = 2 CONNECT BY PRIOR a.crn = a.rn START WITH a.rn = t.rn) AS PARENT FROM UK_PARUS.ACATALOG T) MAIN_crn ON MAIN_crn.RN = MAIN.crn WHERE MAIN_crn.NAME LIKE 'Увол%' AND ((MAIN.MOVEDATE >= TO_DATE('" + Period_S + "', 'yyyy-mm-dd')) AND (MAIN.MOVEDATE <= TO_DATE('" + Period_Po + "', 'yyyy-mm-dd'))) AND MAIN_CRN.PARENT = '" + arrTM_TO + "' AND EXISTS (SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP WHERE UP.CATALOG = MAIN.crn)) SubQuery INNER JOIN (SELECT SubQuery.PRN, SubQuery.NAME, SubQuery.STOP_DATE FROM (SELECT AGNSVCPRD.PRDISMTV, AGNSVCPRD.PRN, SubQuery.NAME, AGNSVCPRD.STOP_DATE FROM (SELECT AGNSVCPRD.PRN, AGNSVCPRD.PRDISMTV, AGNSVCPRD.STOP_DATE FROM UK_PARUS.AGNSVCPRD) AGNSVCPRD INNER JOIN (SELECT PRDISMTV.RN, PRDISMTV.NAME, PRDISMTV.RN AS EXPR1 FROM (SELECT PRDISMTV.RN, PRDISMTV.NAME FROM UK_PARUS.PRDISMTV) PRDISMTV) SubQuery ON AGNSVCPRD.PRDISMTV = SubQuery.RN) SubQuery) SubQuery_1 ON SubQuery.AGENT_RN = SubQuery_1.PRN AND SubQuery.Data_Yvoln = SubQuery_1.STOP_DATE";

                // Sotrudnik_Prinyat = "SELECT MAIN_CRN.NAME AS Name_TO, MAIN.EXEC AS Doljnost, MAIN_AGNRN.CLNGROUP AS Gruppa_Doljnost , MAIN_AGNRN.SEX AS Pol, MAIN.CLNKAD AS MAIN_CLNKAD, MAIN.SLVID AS Vid_Sluj, MAIN_AGNRN.SHORTEXEC AS Mnemo_Dolj, MAIN_AGNRN.CLNKAD AS Kat_Dolj, MAIN_AGNRN.EDVID AS Obrazov, MAIN_AGNRN.EDPROF AS Vid_Obrazov, MAIN_AGNRN.EDSPEC AS Specialnost, MAIN_AGNRN.INICIALI AS FIO, MAIN_AGNRN.AGE AS Vozrast, TO_CHAR(MAIN_AGNRN.LAST, 'dd.MM.yyyy') AS Date_Priema, MAIN_AGNRN.SVCKAT AS Otkyda, MAIN.AGN_RN AS Nomer_Agenta, MAIN.CRN FROM (SELECT DEP.RN AS RN, FR_ADD_COL_PERSONS(DEP.RN) AS DEP_CODE, dm_query.get_svyz(DEP.RN) AS AGN_RN, STS.NAME_NOM AS EXEC, DEP.PSDEP_NAME AS CLNCODE, dm_query.get_clnpsdep_zvan(DEP.RN) AS CLNZVAN, DEP.DO_ACT_FROM AS CLNDOLDAT, DEP.DEPART_DISP AS CLNKOD, dm_query.get_clnpsdep_kad(DEP.RN) AS CLNKAD, dm_query.get_clnpsdep_group(DEP.RN) AS CLNGROUP, dm_query.get_departdisp_post(DEP.RN) AS POST, dm_query.get_depart2(DEP.RN) AS UPRAV, DEP.RATEACC AS KOLSHTAT, DECODE(DEP.ON_STAFF, 0, 'внештатная', 1, 'штатная') AS DOLVID, dm_query.get_block(DEP.RN) AS BLOCK, dm_query.get_pipec(DEP.RN) AS PIPEC, dm_query.get_pipec2(DEP.RN) AS PIPEC2, dm_query.get_nomnaz(DEP.RN) AS NOMNAZ, dm_query.get_slvid(DEP.RN) AS SLVID, dm_query.get_specdol(DEP.RN) AS SPECDOL, dm_query.get_sovdol(DEP.RN) AS SOVDOL, dm_query.get_depart_uprav_ctu(DEP.RN) AS UPRAVL, DEP.CRN AS CRN FROM UK_PARUS.CLNPSDEP DEP INNER JOIN UK_PARUS.CLNPOSTS STS ON DEP.POSTRN = STS.RN INNER JOIN UK_PARUS.CLNPSDEPPRS P ON DEP.RN = P.PRN WHERE DEP.DO_ACT_FROM <= SYSDATE AND (DEP.DO_ACT_TO >= SYSDATE OR DEP.DO_ACT_TO IS NULL)) MAIN LEFT OUTER JOIN (SELECT T.RN AS RN, T.NAME AS NAME, DECODE(T.IS_ROOT, 0, 'нет', 1, 'да') AS IS_ROOT, (SELECT A.NAME FROM UK_PARUS.ACATALOG A WHERE LEVEL = 2 CONNECT BY PRIOR A.crn = A.RN START WITH A.RN = T.RN) AS PARENT FROM UK_PARUS.ACATALOG T) MAIN_CRN ON MAIN_CRN.RN = MAIN.CRN LEFT OUTER JOIN (SELECT A.AGNFAMILYNAME AS AGNFAMILYNAME, A.AGNFIRSTNAME AS AGNFIRSTNAME, A.AGNLASTNAME AS AGNLASTNAME, A.AGNFAMILYNAME || ' ' || A.AGNFIRSTNAME || ' ' || A.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(A.RN) AS INICIALI, A.RN AS RN, A.AGNBURN AS AGNBURN, TRUNC(MONTHS_BETWEEN(SYSDATE, A.AGNBURN) / 12) AS AGE, dm_query.get_prem_ln(A.RN) AS LICNUMB, dm_query.get_depart_uprav_ca(A.RN) AS UPRAV, dm_query.get_depart_upravshort_ca(A.RN) AS UPRAV_SHORT, dm_query.get_discharge1(A.RN) AS DISCHARGE, dm_query.get_depart(A.RN) AS DEPART, A.AGNABBR AS AGNCODE, A.CRN AS CRN, dm_query.get_premplsts(A.RN) AS PRDTKN, DECODE(A.SEX, 1, 'муж', 2, 'жен') AS SEX, A.PRFMLSTS AS SEM, A.PENSION_NBR AS NBR, A.AGNIDNUMB AS ID, A.PHONE AS PHONE, A.PHONE2 AS PHONE2, A.TELEX AS TELEX, dm_query.get_agnsvcprd(A.RN) AS NAME_TO, dm_query.get_first_date(A.RN) AS FIRST, dm_query.get_latter_date(A.RN) AS LATTER, dm_query.get_min_datebg(A.RN) AS LAST, dm_query.get_last_dateend(A.RN) AS TO_END, dm_query.get_isp_contr(A.RN) AS CONTR, dm_query.get_isp_contrnom(A.RN) AS CONNOM, dm_query.get_isp_contrvid(A.RN) AS CONVID, dm_query.get_isp_contrend(A.RN) AS DATEND, dm_query.get_isp_controsn(A.RN) AS OSN, dm_query.get_isp_contrdopnom(A.RN) AS DOPNOM, dm_query.get_isp_contrdopdat(A.RN) AS DOPDAT, dm_query.get_isp_contrdopopis(A.RN) AS DOPOPIS, dm_query.get_isp_contrdopnom40(A.RN) AS DOPNOM40, dm_query.get_isp_contrdopdat40(A.RN) AS DOPDAT40, dm_query.get_isp_contrdopopis40(A.RN) AS DOPOPIS40, dm_query.get_pris(A.RN) AS PRIS, dm_query.get_provdat(A.RN) AS PROV, dm_query.get_provnom(A.RN) AS PROVNOM, dm_query.get_dak(A.RN) AS DAK, dm_query.get_attdat(A.RN) AS ATTDAT, dm_query.get_attns_date(A.RN) AS ATTNSDATE, dm_query.get_attdatnext(A.RN) AS ATTDATNE, dm_query.get_attosn(A.RN) AS ATTOSN, dm_query.get_attviv(A.RN) AS ATTVIV, dm_query.get_attdoctype(A.RN) AS ATTDOCTYPE, dm_query.get_attwho(A.RN) AS ATTWHO, dm_query.get_attdocdat(A.RN) AS ATTDOCDAT, dm_query.get_attdocnom(A.RN) AS ATTDOCNOM, DECODE(dm_query.get_attrlz(A.RN), 1, 'Да', 0, 'нет') AS ATTRLZ, dm_query.get_attdoctype_rlz(A.RN) AS ATTRLZTYPE, dm_query.get_attdocwho_rlz(A.RN) AS ATTRLZWHO, dm_query.get_attdocdat_rlz(A.RN) AS ATTRLZDAT, dm_query.get_attdocnom_rlz(A.RN) AS ATTRLZNOM, dm_query.get_agnrwrd_dat(A.RN) AS WRDDAT, dm_query.get_agnrwrd_who(A.RN) AS WRDWHO, dm_query.get_agnrwrd_numb(A.RN) AS WRDNUMB, dm_query.get_agnrwrd_type(A.RN) AS WRDTYPE, dm_query.get_agnrwrd_typ(A.RN) AS WRDTYP, dm_query.get_agnrwrd_vid(A.RN) AS WRDVID, dm_query.get_agnrwrd_reas(A.RN) AS WRDREAS, dm_query.get_agnpnlts_knd(A.RN) AS KND, dm_query.get_agnpnlts_fault(A.RN) AS FOULT, dm_query.get_agnpnlts_nlk(A.RN) AS NLK, dm_query.get_agnpnlts_css(A.RN) AS CSS, dm_query.get_agnpnlts_ini(A.RN) AS INI, dm_query.get_agnpnlts_pnl(A.RN) AS PNL, dm_query.get_agnpnlts_typeon(A.RN) AS TYPEON, dm_query.get_agnpnlts_whonal(A.RN) AS WHONAL, dm_query.get_agnpnlts_docon(A.RN) AS DOCON, dm_query.get_agnpnlts_docon2(A.RN) AS DOCON2, dm_query.get_agnpnlts_cntrldat(A.RN) AS CNTR, dm_query.get_agnpnlts_appl(A.RN) AS APPL, dm_query.get_agnpnlts_remov(A.RN) AS REMOV, dm_query.get_agnpnlts_osn(A.RN) AS OSN1, dm_query.get_agnpnlts_typeoff(A.RN) AS TYPEOFF, dm_query.get_agnpnlts_whosnal(A.RN) AS WHOSNAL, dm_query.get_agnpnlts_datoff(A.RN) AS DATOFF, dm_query.get_agnpnlts_docoff(A.RN) AS DOCOFF, dm_query.get_agnmedcer_beg(A.RN) AS MEDBEG, dm_query.get_agnmedcer_end(A.RN) AS MEDEND, dm_query.get_agnmedcer_pay(A.RN) AS MEDPAY, DECODE(dm_query.get_agnmedcer_hosp(A.RN), 1, 'Да', 0, 'нет') AS HOSP, dm_query.get_agnmedcer_ser(A.RN) AS MEDSER, dm_query.get_agnmedcer_nom(A.RN) AS MEDNOM, dm_query.get_agnmedcer_dat(A.RN) AS MEDDAT, dm_query.get_agnmedcer_med(A.RN) AS MED, dm_query.get_clnpersvac_vid(A.RN) AS OTPVID, DECODE(dm_query.get_clnpersvac_type(A.RN), 1, 'Дополнительный', 0, 'Основной', 2, 'За свой счет') AS OTPTYPE, dm_query.get_clnpersvac_beg(A.RN) AS OTPBEG, dm_query.get_clnpersvac_end(A.RN) AS OTPEND, dm_query.get_clnpersvac_periodon(A.RN) AS OTPPERON, dm_query.get_clnpersvac_periodoff(A.RN) AS OTPPEROFF, dm_query.get_clnpersvac_planon(A.RN) AS OTPPLANON, dm_query.get_clnpersvac_planoff(A.RN) AS OTPPLANOFF, dm_query.get_clnpersvac_major(A.RN) AS OTPMAJOR, dm_query.get_clnpersvac_minor(A.RN) AS OTPMINOR, dm_query.get_clnpersvac_hol(A.RN) AS OTPHOL, dm_query.get_clnpersvac_trav(A.RN) AS OTPTRAV, dm_query.get_clnpersvac_doctype(A.RN) AS OTPDOCTYPE, dm_query.get_clnpersvac_docwho(A.RN) AS OTPDOCWHO, dm_query.get_clnpersvac_docdat(A.RN) AS OTPDOCDAT, dm_query.get_clnpersvac_docnumb(A.RN) AS OTPDOCNUMB, dm_query.get_clnpersvac_prim(A.RN) AS OTPPRIM, DECODE(dm_query.get_clnpersvac_censel(A.RN), 1, 'Да', 0, 'Нет') AS OTPCENSEL, DECODE(dm_query.get_clnpersvac_hold(A.RN), 1, 'Да', 0, 'Нет') AS OTPHOLD, dm_query.get_clnpersvac_ret(A.RN) AS OTPRET, dm_query.get_clnpersvac_hnote(A.RN) AS OTPHNOTE, dm_query.get_clnpersvac_rettype(A.RN) AS OTPRETTYPE, dm_query.get_clnpersvac_retwho(A.RN) AS OTPRETWHO, dm_query.get_clnpersvac_retdat(A.RN) AS OTPRETDAT, dm_query.get_clnpersvac_retnumb(A.RN) AS OTPRETNUMB, dm_query.get_clnperstrip_beg(A.RN) AS TRIPBEG, dm_query.get_clnperstrip_end(A.RN) AS TRIPEND, dm_query.get_clnperstrip_cel(A.RN) AS TRIPCEL, DECODE(dm_query.get_clnperstrip_cansel(A.RN), 1, 'Да', 0, 'Нет') AS TRIPCANSEL, dm_query.get_clnperstrip_goal(A.RN) AS TRIPGOAL, dm_query.get_clnperstrip_type(A.RN) AS TRIPTYPE, dm_query.get_clnperstrip_who(A.RN) AS TRIPWHO, dm_query.get_clnperstrip_dat(A.RN) AS TRIPDAT, dm_query.get_clnperstrip_numb(A.RN) AS TRIPNUMB, dm_query.get_clnperstrip_adres(A.RN) AS TRIPADRES, dm_query.get_clnperstrip_dopadres(A.RN) AS TRIPDOPADRES, dm_query.get_agnfrntrp_geog(A.RN) AS TRPGEOG, dm_query.get_agnfrntrp_trp(A.RN) AS TRPTRP, dm_query.get_agnfrntrp_beg(A.RN) AS TRPBEG, dm_query.get_agnfrntrp_end2(A.RN) AS TRPEND, dm_query.get_agnaddinf_type(A.RN) AS DDTYPE, dm_query.get_agnaddinf_datbeg(A.RN) AS DDDATBEG, dm_query.get_agnaddinf_datend(A.RN) AS DDDATEND, dm_query.get_agnaddinf_sod(A.RN) AS DDSOD, dm_query.get_sppvd_vid(A.RN) AS SPPVID, dm_query.get_sppvd_numb(A.RN) AS SPPNUMB, dm_query.get_sppvd_docdat(A.RN) AS SPPDOCDAT, dm_query.get_sppvd_docwho(A.RN) AS SPPDOCWHO, dm_query.get_sppvd_vidpro(A.RN) AS SPPVIDPRO, dm_query.get_sppvd_datpro(A.RN) AS SPPDATPRO, dm_query.get_sppvd_opis(A.RN) AS SPPOPIS, dm_query.get_sppvd_stat(A.RN) AS SPPSTAT, dm_query.get_sppvd_mesto(A.RN) AS SPPMESTO, DECODE(dm_query.get_sppvd_vina(A.RN), 1, 'Да', 0, 'Нет') AS SPPVINA, dm_query.get_sppvd_oznak(A.RN) AS SPPOZNAK, DECODE(dm_query.get_sppvd_noznak(A.RN), 1, 'Да', 0, 'Нет') AS SPPNOZNAK, dm_query.get_sppvd_prim(A.RN) AS SPPPRIM, dm_query.get_agnacclev_form(A.RN) AS CCFORM, dm_query.get_agnacclev_s(A.RN) AS CCS, dm_query.get_agnacclev_po(A.RN) AS CCPO, dm_query.get_agnacclev_prim(A.RN) AS CCPRIM, dm_query.get_agnacclev_type(A.RN) AS CCTYPE, dm_query.get_agnacclev_who(A.RN) AS CCWHO, dm_query.get_agnacclev_dat(A.RN) AS CCDAT, dm_query.get_agnacclev_numb(A.RN) AS CCNUMB, dm_query.get_profpsih_dat(A.RN) AS SIHDAT, dm_query.get_profpsih_kat(A.RN) AS SIHKAT, dm_query.get_profpsih_zak(A.RN) AS SIHZAK, dm_query.get_profpsih_note(A.RN) AS SIHNOTE, dm_query.get_profvalid_dat(A.RN) AS LIDDAT, dm_query.get_profvalid_kat(A.RN) AS LIDKAT, dm_query.get_profvalid_who(A.RN) AS LIDWHO, dm_query.get_profvalid_nom(A.RN) AS LIDNOM, dm_query.get_profvalid_note(A.RN) AS LIDNOTE, dm_query.get_agnranks_zvan(A.RN) AS NKSZVAN, dm_query.get_agnranks_vid(A.RN) AS NKSVID, dm_query.get_agnranks_dat(A.RN) AS NKSDAT, dm_query.get_agnranks_ndat(A.RN) AS NKSNDAT, dm_query.get_agnranks_type(A.RN) AS NKSTYPE, dm_query.get_agnranks_who(A.RN) AS NKSWHO, dm_query.get_agnranks_prdat(A.RN) AS NKSPRDAT, dm_query.get_agnranks_nom(A.RN) AS NKSNOM, dm_query.get_agneduc_vidob(A.RN) AS EDVID, dm_query.get_agneduc_prof(A.RN) AS EDPROF, dm_query.get_agneduc_begdat(A.RN) AS EDBEGDAT, dm_query.get_agneduc_enddat(A.RN) AS EDENDDAT, dm_query.get_agneduc_educ(A.RN) AS EDEDUC, dm_query.get_agneduc_form(A.RN) AS EDFORM, DECODE(dm_query.get_agneduc_otm(A.RN), 1, 'Учится', 0, 'Не задана', 2, 'Учится на последнем курсе') AS EDOTM, dm_query.get_agneduc_dtype(A.RN) AS EDDTYPE, dm_query.get_agneduc_ser(A.RN) AS EDSER, dm_query.get_agneduc_nom(A.RN) AS EDNOM, dm_query.get_agneduc_spec(A.RN) AS EDSPEC, dm_query.get_agneduc_kval(A.RN) AS EDKVAL, dm_query.get_agneduc_type(A.RN) AS EDTYPE, dm_query.get_agneduc_who(A.RN) AS EDWHO, dm_query.get_agneduc_ddoc(A.RN) AS EDDDOC, dm_query.get_agneduc_dnom(A.RN) AS EDDNOM, dm_query.get_agneducperv_mesto(A.RN) AS EDPERVMESTO, dm_query.get_agneducperv_tipdoc(A.RN) AS EDPERVTIPDOC, dm_query.get_agneducperv_ser(A.RN) AS EDPERVSER, dm_query.get_agneducperv_nom(A.RN) AS EDPERVNOM, dm_query.get_agneducperv_bdate(A.RN) AS EDPERVBDATE, dm_query.get_agneducperv_edate(A.RN) AS EDPERVEDATE, dm_query.get_agneducperv_type(A.RN) AS EDPERVTYPE, dm_query.get_agneducperv_who(A.RN) AS EDPERVWHO, dm_query.get_agneducperv_dat(A.RN) AS EDPERVDAT, dm_query.get_agneducperv_numb(A.RN) AS EDPERVNUMB, dm_query.get_agneducpov_vidobr(A.RN) AS POVVIDOBR, dm_query.get_agneducpov_uch(A.RN) AS POVUCH, dm_query.get_agneducpov_kurs(A.RN) AS POVKURS, dm_query.get_agneducpov_form(A.RN) AS POVFORM, dm_query.get_agneducpov_fin(A.RN) AS POVFIN, dm_query.get_agneducpov_prof(A.RN) AS POVPROF, dm_query.get_agneducpov_bdate(A.RN) AS POVBDATE, dm_query.get_agneducpov_edate(A.RN) AS POVEDATE, dm_query.get_agneducpov_chas(A.RN) AS POVCHAS, dm_query.get_agneducpov_doctype(A.RN) AS POVDOCTYPE, dm_query.get_agneducpov_docser(A.RN) AS POVDOCSER, dm_query.get_agneducpov_docnom(A.RN) AS POVDOCNOM, dm_query.get_agneducpov_type(A.RN) AS POVTYPE, dm_query.get_agneducpov_who(A.RN) AS POVWHO, dm_query.get_agneducpov_dat(A.RN) AS POVDAT, dm_query.get_agneducpov_nom(A.RN) AS POVNOM, dm_query.get_agneducper_vid(A.RN) AS PERVID, dm_query.get_agneducper_uch(A.RN) AS PERUCH, dm_query.get_agneducper_kurs(A.RN) AS PERKURS, dm_query.get_agneducper_form(A.RN) AS PERFORM, dm_query.get_agneducper_fin(A.RN) AS PERFIN, dm_query.get_agneducper_prof(A.RN) AS PERPROF, dm_query.get_agneducper_bdate(A.RN) AS PERBDATE, dm_query.get_agneducper_edate(A.RN) AS PEREDATE, dm_query.get_agneducper_chas(A.RN) AS PERCHAS, dm_query.get_agneducper_doctype(A.RN) AS PERDOCTYPE, dm_query.get_agneducper_docser(A.RN) AS PERDOCSER, dm_query.get_agneducper_docnom(A.RN) AS PERDOCNOM, dm_query.get_agneducper_type(A.RN) AS PERTYPE, dm_query.get_agneducper_who(A.RN) AS PERWHO, dm_query.get_agneducper_dat(A.RN) AS PERDAT, dm_query.get_agneducper_nom(A.RN) AS PERNOM, dm_query.get_agneducstaj_mesto(A.RN) AS STAJMESTO, dm_query.get_agneducstaj_fin(A.RN) AS STAJFIN, dm_query.get_agneducstaj_bdate(A.RN) AS STAJBDATE, dm_query.get_agneducstaj_edate(A.RN) AS STAJEDATE, dm_query.get_agneducstaj_type(A.RN) AS STAJTYPE, dm_query.get_agneducstaj_who(A.RN) AS STAJWHO, dm_query.get_agneducstaj_dat(A.RN) AS STAJDAT, dm_query.get_agneducstaj_nom(A.RN) AS STAJNOM, dm_query.get_agneducstep_uch(A.RN) AS STEPUCH, dm_query.get_agneducstep_spec(A.RN) AS STEPSPEC, dm_query.get_agneducstep_dat(A.RN) AS STEPDAT, dm_query.get_agneducstep_type(A.RN) AS STEPTYPE, dm_query.get_agneducstep_ser(A.RN) AS STEPSER, dm_query.get_agneducstep_nom(A.RN) AS STEPNOM, dm_query.get_agneducstep_note(A.RN) AS STEPNOTE, dm_query.get_agneduczvan_uch(A.RN) AS ZVANUCH, dm_query.get_agneduczvan_dat(A.RN) AS ZVANDAT, dm_query.get_agneduczvan_type(A.RN) AS ZVANTYPE, dm_query.get_agneduczvan_ser(A.RN) AS ZVANSER, dm_query.get_agneduczvan_nom(A.RN) AS ZVANNOM, dm_query.get_agneduczvan_note(A.RN) AS ZVANNOTE, dm_query.get_agnlangs_in(A.RN) AS LANIN, dm_query.get_agnlangs_step(A.RN) AS LANSTEP, dm_query.get_agnlangs_note(A.RN) AS LANNOTE, dm_query.get_agnatnms_vid(A.RN) AS NMSVID, dm_query.get_agnatnms_step(A.RN) AS NMSSTEP, dm_query.get_docsprops_agnatnms(A.RN) AS DPNM, dm_query.get_agnaddresses_full(A.RN) AS ADRESFULL, dm_query.get_agnaddresses_index(A.RN) AS ADRESINDEX, dm_query.get_agnaddresses_dom(A.RN) AS ADRESDOM, dm_query.get_agnaddresses_blok(A.RN) AS ADRESBLOK, dm_query.get_agnaddresses_flat(A.RN) AS ADRESFLAT, dm_query.get_agnaddresses_note(A.RN) AS ADRESNOTE, dm_query.get_agndocums_type(A.RN) AS CUMSTYPE, dm_query.get_agndocums_ser(A.RN) AS CUMSSER, dm_query.get_agndocums_numb(A.RN) AS CUMSNUMB, dm_query.get_agndocums_vida(A.RN) AS CUMSVIDA, dm_query.get_agndocums_po(A.RN) AS CUMSPO, dm_query.get_agndocums_who(A.RN) AS CUMSWHO, dm_query.get_agndocums_osn(A.RN) AS CUMSOSN, dm_query.get_agndocums_doctype(A.RN) AS CUMSDOCTYPE, dm_query.get_agndocums_docwho(A.RN) AS CUMSDOCWHO, dm_query.get_agndocums_docdat(A.RN) AS CUMSDOCDAT, dm_query.get_agndocums_docnumb(A.RN) AS CUMSDOCNUMB, dm_query.get_agnsprtsk_kval(A.RN) AS TSKKVAL, dm_query.get_agnsprtsk_vid(A.RN) AS TSKVID, dm_query.get_agnsprtsk_god(A.RN) AS TSKGOD, dm_query.get_agnlbractivityprs_bdate(A.RN) AS TRUDBDATE, dm_query.get_agnlbractivityprs_edate(A.RN) AS TRUDEDATE, dm_query.get_agnlbractivityprs_gos(A.RN) AS TRUDGOS, dm_query.get_agnlbractivityprs_work(A.RN) AS TRUDWORK, dm_query.get_agnlbractivityprs_dolj(A.RN) AS TRUDDOLJ, dm_query.get_agnlbractivityprs_prich(A.RN) AS TRUDPRICH, dm_query.get_agnlbractivityprs_type(A.RN) AS TRUDTYPE, dm_query.get_agnlbractivityprs_who(A.RN) AS TRUDWHO, dm_query.get_agnlbractivityprs_date(A.RN) AS TRUDDATE, dm_query.get_agnlbractivityprs_numb(A.RN) AS TRUDNUMB, dm_query.get_agnlbractivityprs_note(A.RN) AS TRUDNOTE, dm_query.get_agnsvcprd_sluj(A.RN) AS SVCSLUJ, dm_query.get_agnsvcprd_post(A.RN) AS SVCPOST, dm_query.get_agnsvcprd_post2(A.RN) AS SVCPOST2, dm_query.get_agnsvcprd_vidpost(A.RN) AS SVCVIDPOST, dm_query.get_agnsvcprd_kat(A.RN) AS SVCKAT, dm_query.get_agnsvcprd_primtype(A.RN) AS SVCPRIMTYPE, dm_query.get_agnsvcprd_primwho(A.RN) AS SVCPRIMWHO, dm_query.get_agnsvcprd_primdate(A.RN) AS SVCPRIMDATE, dm_query.get_agnsvcprd_primnumb(A.RN) AS SVCPRIMNUMB, dm_query.get_agnsvcprd_ispdat(A.RN) AS SVCISPDAT, dm_query.get_agnsvcprd_ispedat(A.RN) AS SVCISPEDAT, dm_query.get_agnsvcprd_ispzak(A.RN) AS SVCISPZAK, dm_query.get_agnsvcprd_isptype(A.RN) AS SVCISPTYPE, dm_query.get_agnsvcprd_ispwho(A.RN) AS SVCISPWHO, dm_query.get_agnsvcprd_ispdate(A.RN) AS SVCISPDATE, dm_query.get_agnsvcprd_ispnumb(A.RN) AS SVCISPNUMB, dm_query.get_agnsvcmov_bdate(A.RN) AS MOVBDATE, dm_query.get_agnsvcmov_edate(A.RN) AS MOVEDATE, dm_query.get_agnsvcmov_pervid(A.RN) AS MOVPERVID, dm_query.get_agnsvcmov_perprich(A.RN) AS MOVPERPRICH, dm_query.get_agnsvcmov_permot(A.RN) AS MOVPERMOT, dm_query.get_agnsvcmov_mesto(A.RN) AS MOVMESTO, dm_query.get_agnsvcmov_org(A.RN) AS MOVORG, dm_query.get_agnsvcmov_block(A.RN) AS MOVBLOCK, dm_query.get_agnsvcmov_group(A.RN) AS MOVGROUP, dm_query.get_agnsvcmov_kod(A.RN) AS MOVKOD, dm_query.get_agnsvcmov_dolj(A.RN) AS MOVDOLJ, dm_query.get_agnsvcmov_fdolj(A.RN) AS MOVFDOLJ, dm_query.get_agnsvcmov_type(A.RN) AS MOVTYPE, dm_query.get_agnsvcmov_who(A.RN) AS MOVWHO, dm_query.get_agnsvcmov_date(A.RN) AS MOVDATE, dm_query.get_agnsvcmov_numb(A.RN) AS MOVNUMB, dm_query.get_agnsvcmov_prekprich(A.RN) AS MOVPRIKPRICH, dm_query.get_agnsvcmov_prektype(A.RN) AS MOVPRIKTYPE, dm_query.get_agnsvcmov_prekwho(A.RN) AS MOVPRIKWHO, dm_query.get_agnsvcmov_prekdate(A.RN) AS MOVPRIKDATE, dm_query.get_agnsvcmov_preknumb(A.RN) AS MOVPRIKNUMB, dm_query.get_agnsvcmov_note(A.RN) AS MOVNOTE, dm_query.get_max_execution(A.RN) AS EXEC, dm_query.get_max_shortexecution(A.RN) AS SHORTEXEC, dm_query.get_max_dateexecution(A.RN) AS DATEEXECUTION, dm_query.get_clnpsdep_code(A.RN) AS CLNCODE, dm_query.get_clnpsdep_zvan2(A.RN) AS CLNZVAN, dm_query.get_clnpsdep_doldat(A.RN) AS CLNDOLDAT, dm_query.get_clnpsdep_type(A.RN) AS CLNTYPE, dm_query.get_clnpsdep_who(A.RN) AS CLNWHO, dm_query.get_clnpsdep_date(A.RN) AS CLNDATE, dm_query.get_clnpsdep_numb(A.RN) AS CLNNUMB, dm_query.get_clnpsdep_kod(A.RN) AS CLNKOD, dm_query.get_clnpsdep_kad(A.RN) AS CLNKAD, dm_query.get_clnpsdep_group(A.RN) AS CLNGROUP FROM UK_PARUS.AGNLIST A WHERE A.AGNTYPE = 1 AND A.RN IN (SELECT PP.AGNLIST FROM UK_PARUS.PREMPLFLS PP)) MAIN_AGNRN ON MAIN_AGNRN.RN = MAIN.AGN_RN WHERE MAIN.CRN =  " + arrTM_depart + "  AND " + Peroid_Parametr + "  AND (MAIN_AGNRN.SVCKAT != 'Внутри т/о' and MAIN_AGNRN.SVCKAT != 'из др.там.органа') AND EXISTS (SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP WHERE UP.CATALOG = MAIN.CRN)";
            }

            //Sotrudnik_Prinyat = "SELECT MAIN_crn.NAME AS Name_TO,   MAIN.MOVORG AS Mnemo_TO,          MAIN_crn.PARENT AS Verhnii_Level,          MAIN.INICIALI AS FIO,           MAIN.SEX AS Pol,          MAIN.AGE AS Vozrast,           MAIN.EXEC AS Doljnost,       MAIN.CLNGROUP AS Gruppa_Doljnost,           MAIN.PRDTKN AS Status_Slujb,           MAIN.SVCSLUJ AS Vid_Sluj,           TO_CHAR(TO_DATE(MAIN.DATEEXECUTION, 'dd.MM.yyyy'), 'dd.MM.yyyy') AS Data_Naznachenia, MAIN.SHORTEXEC AS Mnemo_Dolj, MAIN.CLNKAD AS Kat_Dolj,           MAIN.EDVID AS Obrazov, MAIN.EDPROF AS Vid_Obrazov,  MAIN.EDSPEC AS Specialnost,TO_CHAR(MAIN.LAST, 'dd.MM.yyyy') AS Date_Priema, MAIN.SVCKAT AS Otkyda FROM(SELECT a.AGNFAMILYNAME || ' ' || a.AGNFIRSTNAME || ' ' || a.AGNLASTNAME AS FULLNAME, dm_query.get_iniciali(a.RN) AS INICIALI, TRUNC(MONTHS_BETWEEN(SYSDATE, a.AGNBURN) / 12) AS AGE, a.CRN AS crn,               dm_query.get_premplsts(a.RN) AS PRDTKN, DECODE(a.SEX, 1, 'муж', 2, 'жен') AS SEX, a.PRFMLSTS AS SEM,                dm_query.get_latter_date(a.RN) AS LATTER, dm_query.get_min_datebg(a.RN) AS LAST, dm_query.get_agneduc_vidob(a.RN) AS EDVID, dm_query.get_agneduc_prof(a.RN) AS EDPROF, dm_query.get_agneduc_spec(a.RN) AS EDSPEC, dm_query.get_agnsvcprd_sluj(a.RN) AS SVCSLUJ, dm_query.get_agnsvcprd_kat(a.RN) AS SVCKAT, dm_query.get_agnsvcmov_org(a.RN) AS MOVORG, dm_query.get_max_execution(a.RN) AS EXEC, dm_query.get_max_shortexecution(a.RN) AS SHORTEXEC, dm_query.get_max_dateexecution(a.RN) AS DATEEXECUTION, dm_query.get_clnpsdep_kad(a.RN) AS CLNKAD, dm_query.get_clnpsdep_group(a.RN) AS CLNGROUP       FROM UK_PARUS.AGNLIST a      WHERE a.AGNTYPE = 1        AND a.RN IN(SELECT pp.AGNLIST FROM UK_PARUS.PREMPLFLS pp)) MAIN LEFT OUTER JOIN(SELECT T.RN AS RN, T.NAME AS NAME, DECODE(T.IS_ROOT, 0, 'нет', 1, 'да') AS IS_ROOT,                            (SELECT A.NAME                                FROM UK_PARUS.ACATALOG A                                WHERE LEVEL = 2                              CONNECT BY PRIOR a.crn = a.rn                              START WITH a.rn = t.rn) AS PARENT        FROM UK_PARUS.ACATALOG T) MAIN_crn ON MAIN_crn.RN = MAIN.crn  WHERE MAIN.MOVORG = '"+ 10100000 +"' and   " + Peroid_Parametr +"AND MAIN.SVCKAT <> 'Внутри т/о'    AND MAIN.SVCKAT <> 'из др.там.органа'    AND MAIN.SVCSLUJ <> 'Работник'    AND EXISTS(SELECT UP.RN, UP.AUTHID, UP.ROLEID, UP.COMPANY, UP.VERSION, UP.UNITCODE, UP.CATALOG, UP.JUR_PERS, UP.HIERARCHY FROM UK_PARUS.V_USERPRIV UP        WHERE UP.CATALOG = MAIN.crn)  ORDER BY MAIN.LATTER";

            OracleCommand cmdYvol = new OracleCommand();

            // Сочетать Command с Connection.
            cmdYvol.Connection = conn;
            cmdYvol.CommandText = YvolSQL;
            massiv._Yvolenn6KAD.Clear();
            using (DbDataReader reader = cmdYvol.ExecuteReader())
            {

                while (reader.Read())
                {
                    massiv._Yvolenn6KAD.Add(new Yvolenn6KAD()
                    {



                        TO = (reader["TORG"].ToString()),
                        FIO = (reader["FIO"].ToString()),
                        Vozrast = (reader["Vozrast"].ToString()),
                        Agen_RN = (reader["Agen_RN"].ToString()),
                        Date_Fin = (reader["Date_Fin"].ToString()),
                        Date_Nach = (reader["Date_Nach"].ToString()),
                        Date_Yvol = (reader["Date_Yvol"].ToString()),
                        Katalog = (reader["Katalog"].ToString()),
                        Motiv = (reader["Motiv"].ToString()),
                        POL = (reader["POL"].ToString()),
                        Vid_Obraz = (reader["Vid_Obraz"].ToString()),
                        Vid_Slujb = (reader["Vid_Slujb"].ToString()),
                        Gruppa_Dolj = (reader["Gruppa_Dolj"].ToString()),
                        Kat_Dolj = (reader["Kat_Dolj"].ToString()),
                        Mnemo_Dolj = (reader["Mnemo_Dolj"].ToString()),

                    });
                }
                cmdYvol.Parameters.Clear();
                conn.Close();

            }

            OracleCommand cmd = new OracleCommand();
            conn.Open();
            // Сочетать Command с Connection.
            cmd.Connection = conn;
            cmd.CommandText = Sotrudnik_Prinyat;
            massiv._Sotrudniki6KAD.Clear();
            try
            {
                //int stroka = 0;
                using (DbDataReader reader = cmd.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        massiv._Sotrudniki6KAD.Add(new Sotrudniki6KAD()
                        {



                            TO = (reader["Name_TO"].ToString()),
                            Doljnost = (reader["Doljnost"].ToString()),
                            Data_Naznachenia = (reader["Data_Naznachenia"].ToString()),
                            Vid_Sluj = (reader["Vid_Sluj"].ToString()),
                            Mnemo_Dolj = (reader["Mnemo_Dolj"].ToString()),
                            Kat_Dolj = (reader["Kat_Dolj"].ToString()),
                            Obrazov = (reader["Obrazov"].ToString()),
                            Vid_Obrazov = (reader["Vid_Obrazov"].ToString()),
                            Specialnost = (reader["Specialnost"].ToString()),
                            FIO = (reader["FIO"].ToString()),
                            Vozrast = (reader["Vozrast"].ToString()),
                            Date_Priema = (reader["Date_Priema"].ToString()),
                            Otkyda = (reader["Otkyda"].ToString()),
                            Mnemo_TO = (reader["Mnemo_TO"].ToString()),
                            Pol = (reader["Pol"].ToString()),
                            Gruppa_Doljnost = (reader["Gruppa_Doljnost"].ToString()),
                            Status_Slujb = (reader["Status_Slujb"].ToString()),
                            Verhnii_Level = (reader["Verhnii_Level"].ToString()),

                        });
                    }
                    cmd.Parameters.Clear();
                    conn.Close();

                }

               

                Zapolnenie_Cad_6();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace);
            }

            return 1;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var Select = Int32.Parse(comboBox1.SelectedIndex.ToString());

            if (Select != 15)
            {
                var n = comboBox1.SelectedIndex.ToString();
                arrTM_RN = Convert.ToInt32(massiv.TM[0, Convert.ToInt32(n)]);
                arrTM_TO = massiv.TM[1, Convert.ToInt32(n)];
                arrTM_depart = Convert.ToInt32(massiv.TM[2, Convert.ToInt32(n)]);
                arrTM_mnemo = Convert.ToInt32(massiv.TM[3, Convert.ToInt32(n)]); //в запрос надо это
            }
            else
            {
                arrTM_depart = 0;
            }

            if ((Period_S != "") || (Period_Po !=""))
            {
                CAD_6.Enabled = true;
            }
        }



        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            items_period = Int32.Parse(checkedListBox1.SelectedIndex.ToString());
            items_period_text = checkedListBox1.SelectedItem.ToString();
            label2.Text = Convert.ToString(items_period);
            var list = sender as CheckedListBox;
            if (e.NewValue == CheckState.Checked)
                foreach (int index in list.CheckedIndices)
                    if (index != e.Index)
                        list.SetItemChecked(index, false);

            if (arrTM_depart != -5)
            {
                CAD_6.Enabled = true;
            }
        }

        private int Zapolnenie_Cad_6()
        {

            Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //сам эксель
            string path_Save_CAD_6 = String.Format(@"D:\Справки\KAD-6_" + arrTM_TO);
            string File_Name_Shablon = (AppDomain.CurrentDomain.BaseDirectory + @"Templates\KAD6.xlsx");
            Excel.Worksheet WS;// в этот лист
            Workbook WB = ObjWorkExcel.Workbooks.Open(Convert.ToString(File_Name_Shablon));
            int RowCount = 0;
            WS = (Excel.Worksheet)WB.Sheets[1];
            WS.Activate();
            ObjWorkExcel.DisplayAlerts = false;
            //Считывание с определенных ячеек
            //string NameTamozn = Convert.ToString(WS.Range["A1"].Value);



            try
            {
                var stroka = 32;
                var stolbec = 3;
                var stolbec_bykv = "C";

                #region
                
                foreach (Sotrudniki6KAD sotrudniki6KAD in massiv._Sotrudniki6KAD)
                {
                    //вновь принятые
                    if ((sotrudniki6KAD.Vid_Sluj == "Сотрудник") || (sotrudniki6KAD.Status_Slujb== "Проходит службу_С") )
                    {
                        WS = (Excel.Worksheet)WB.Sheets[1];
                        WS.Activate();

                        if (sotrudniki6KAD.Mnemo_Dolj == "руководитель ФТС_ЦА")//рук.фтс росии
                        {
                            stroka = 32;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "перв.зам.рук.ФТС_ЦА")//первый зам рук.фтс росии
                        {
                            stroka = 33;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.рук.ФТС России_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "статс-секретарь_ЦА"))//зам рук.фтс росии
                        {
                            stroka = 34;
                        }
                         if (sotrudniki6KAD.Mnemo_Dolj == "нач.глав.управления_ЦА")//нач главного управл
                    {
                        stroka = 35;
                    }
                        if (sotrudniki6KAD.Mnemo_Dolj == "перв.зам.нач.ГУ_ЦА")//первый зам нач главного управл
                        {
                            stroka = 36;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.ГУ_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.гл.упр.-нач.секр.рук.ФТС_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.ГУ-глав.бух_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.ГУ-нач.отдела_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.ГУ-нач.служ_ЦА"))//зам нач главного управл
                        {
                            stroka = 37;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "нач.управления_ЦА")//нач управл
                        {
                            stroka = 38;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "перв.зам.нач.упр_ЦА")//первый зам нач управл
                        {
                            stroka = 39;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.управ_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.управ-нач.службы_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.управл.-нач.отдела_ЦА"))//зам нач управл
                        {
                            stroka = 40;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "гл.советник_ЦА")//главный совет рук фтс россии
                        {
                            stroka = 41;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "советник.рук.ФТС_ЦА")//советник рукв фтс росс
                        {
                            stroka = 42;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "советник.нач.гл.упр_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "советник.нач.упр_ЦА"))//советник нач главн управ
                        {
                            stroka = 43;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "пом.перв.зам.рук.ФТС_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "пом.зам.рук.ФТС_ЦА"))//помощ первого зам руковод фтс россии 
                        {
                            stroka = 44;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "нач.службы_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.главбух-нач.отдела_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.главбух.ФТС_ЦА"))//нач службы зам главн бухг
                        {
                            stroka = 45;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы_НО_ЦА"))//зам нач служб
                        {
                            stroka = 46;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "нач. управления_РТУ")//нач рту
                        {
                            stroka = 48;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "перв.зам.нач.упр_РТУ")//первый зам рту
                        {
                            stroka = 49;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.управ_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.управ-нач.службы_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.управ-нач.отдела_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.управ-нач.там_РТУ"))//зам нач рту
                        {
                            stroka = 50;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "пом.нач.управл_РТУ")//помощ нач рту
                        {
                            stroka = 51;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "нач.службы_РТУ")//нач служб рту
                        {
                            stroka = 52;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "перв.зам.нач.службы_РТУ")//первый зам нач служб рту
                        {
                            stroka = 53;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы-нач.отдела_РТУ"))//зам нач служб рту 
                        {
                            stroka = 54;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "нач.таможни_Т")//нач таможн 
                        {
                            stroka = 55;
                        }
                        if (sotrudniki6KAD.Mnemo_Dolj == "перв.зам.нач.там._Т")//первый зам нач таможн
                        {
                            stroka = 56;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.таможни_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.там.ПОД_Т") || (sotrudniki6KAD.Mnemo_Dolj == "Зам.нач.там.по раб.с кадрами_Т") || (sotrudniki6KAD.Mnemo_Dolj == "Зам.нач.там.по кадрам_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.там-нач.службы_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.там-нач.поста_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.там-нач.отдела_Т") || (sotrudniki6KAD.Mnemo_Dolj == "Зам.нач.там. по ЭД_Т") || (sotrudniki6KAD.Mnemo_Dolj == "Зам.нач.там. по тыл.обесп_Т"))//зам нач таможн
                        {
                            stroka = 57;
                                                    }
                        if (sotrudniki6KAD.Mnemo_Dolj == "нач.служ_Т")//нач служб таможн
                        {
                            stroka = 58;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы-нач.отдела_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы-нач.отдела-глав.бух_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.службы-глав.бух._Т"))//зам нач служб таможн
                        {
                            stroka = 59;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "нач.поста_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "нач.поста-глав.бух_ТП"))//нач таможн поста
                        {
                            stroka = 60;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.поста_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.поста-нач.отд._ТП") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.поста-нач.отделения_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.поста-гл.бухг_ТП"))//зам нач таможн поста
                        {
                            stroka = 61;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отд-гл.бухгалтер_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отд-зам.гл.бухгалтер_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.СОБРа_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_Т_ЛС") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отд-гл.бухгалтер_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отд-зам.гл.бухгалтер_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "Нач.СОБРа_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_РТУ_ПС") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_РТУ_БС") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_РТУ_ЛС") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_РТУ_ИТС") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_Т_М") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела-гл.бухг_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдела_РТА") || (sotrudniki6KAD.Mnemo_Dolj == "нач.СОБРа_РТУ"))//нач отдела нач-к собра
                        {
                            stroka = 62;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-гл.метрол_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-нач.отделен_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_Т_М") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_Т_ЛС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-зам.глав.бух_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "Зам.нач.СОБРа_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_ПС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_БС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_ЛС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_ИТС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.СОБРа_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_БТС_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_РТА") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_Т_ИТС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-зам.глав.бух_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_БТС_Т") || (sotrudniki6KAD.Mnemo_Dolj == "Зам.нач.отдела-нач.отделения_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.СОБРа_Т") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_СТС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_МТС") || (sotrudniki6KAD.Mnemo_Dolj == "зам.нач.отдела_ТП"))//зам нач отдела зам нач собра
                        {
                            stroka = 63;
                        }
                        if ((sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_ЦА") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_Т_М") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения-командир_СТС_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения-командир_МТС_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения-командир_БТС_Т") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отдел-командир_СТС") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_ТП") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_РТУ_ПС") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_РТУ") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_РТА") || (sotrudniki6KAD.Mnemo_Dolj == "нач.отделения_РТУ_БС"))//нач отделения
                        {
                            stroka = 64;
                        }
                        if (sotrudniki6KAD.Kat_Dolj == "Главные инспекторы")//главный инспектор
                        {
                            stroka = 65;
                        }
                        if (sotrudniki6KAD.Kat_Dolj == "Ведущие инспекторы")//ведущий инстпектор
                        {
                            stroka = 66;
                        }
                        if (sotrudniki6KAD.Kat_Dolj == "Старшие инспекторы")//старший инспектор
                        {
                            stroka = 67;
                        }
                        if (sotrudniki6KAD.Kat_Dolj == "Инспекторы")//инспектор
                        {
                            stroka = 68;
                        }
                        if (sotrudniki6KAD.Kat_Dolj == "Младшие инспекторы")//младший инспектор
                        {
                            stroka = 69;
                        }

                        string adr = "C" + stroka;
                        Raschet(adr, WS);
                        WS.Cells[stroka, 3] = Kol_vo;
                        string fio_primech = sotrudniki6KAD.FIO.ToString();
                        RaschetFio(adr, WS, fio_primech);
                        WS.Range[adr].ClearComments();
                        WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        if (sotrudniki6KAD.Pol == "жен")
                        {

                            adr = "C71";
                            Raschet(adr, WS);
                            WS.Cells[71, 3] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        }


                        //высшее
                        if ((sotrudniki6KAD.Obrazov == "ВО.бакалавриат") || (sotrudniki6KAD.Obrazov == "ВО.магистратура") || (sotrudniki6KAD.Obrazov == "ВО.специалитет") || (sotrudniki6KAD.Obrazov == "высшее") || (sotrudniki6KAD.Obrazov == "высшее проф."))
                        {

                            adr = "D" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 4] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "D71";
                                Raschet(adr, WS);
                                WS.Cells[71, 4] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //среднее проф
                        if ((sotrudniki6KAD.Obrazov == "СО.проф") || (sotrudniki6KAD.Obrazov == "среднее проф."))
                        {
                            adr = "F" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 6] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "F71";
                                Raschet(adr, WS);
                                WS.Cells[71, 6] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //среднее полное общее
                        if (sotrudniki6KAD.Obrazov == "среднее общее")
                        {
                            adr = "H" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 8] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "H71";
                                Raschet(adr, WS);
                                WS.Cells[71, 8] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //9 столбец
                        if (sotrudniki6KAD.Vid_Obrazov == "физ.-мат.и ест.науки")
                        {
                            adr = "I" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 9] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "I71";
                                Raschet(adr, WS);
                                WS.Cells[71, 9] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //10 столбец
                        if ((sotrudniki6KAD.Vid_Obrazov == "гум.и социал.науки") || (sotrudniki6KAD.Vid_Obrazov == "юриспруденция") || (sotrudniki6KAD.Vid_Obrazov == "правоохр.деят."))
                        {
                            adr = "J" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 10] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "J71";
                                Raschet(adr, WS);
                                WS.Cells[71, 10] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }

                            //11 столбец
                            if ((sotrudniki6KAD.Specialnost == "Юриспруденция") || (sotrudniki6KAD.Specialnost == "Правоох.деят."))
                            {

                                adr = "K" + stroka;
                                Raschet(adr, WS);
                                WS.Cells[stroka, 11] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (sotrudniki6KAD.Pol == "жен")
                                {

                                    adr = "K71";
                                    Raschet(adr, WS);
                                    WS.Cells[71, 11] = Kol_vo;
                                    fio_primech = sotrudniki6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }


                        }
                        //12 столбец
                        if ((sotrudniki6KAD.Vid_Obrazov == "экономика и управ.") || (sotrudniki6KAD.Vid_Obrazov == "таможенное дело") || (sotrudniki6KAD.Vid_Obrazov == "гос.и мун.управление"))
                        {

                            adr = "L" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 12] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "L71";
                                Raschet(adr, WS);
                                WS.Cells[71, 12] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }

                            //13 столбец
                            if (sotrudniki6KAD.Specialnost == "Таможенное дело")
                            {
                                adr = "M" + stroka;
                                Raschet(adr, WS);
                                WS.Cells[stroka, 13] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (sotrudniki6KAD.Pol == "жен")
                                {

                                    adr = "M71";
                                    Raschet(adr, WS);
                                    WS.Cells[71, 13] = Kol_vo;
                                    fio_primech = sotrudniki6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }

                            //14 столбец
                            if (sotrudniki6KAD.Specialnost == "ГосМунУправление")
                            {
                                adr = "N" + stroka;
                                Raschet(adr, WS);
                                WS.Cells[stroka, 14] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (sotrudniki6KAD.Pol == "жен")
                                {

                                    adr = "N71";
                                    Raschet(adr, WS);
                                    WS.Cells[71, 14] = Kol_vo;
                                    fio_primech = sotrudniki6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }
                        }

                        //15 столбец
                        if (sotrudniki6KAD.Vid_Obrazov == "образ.и педагогика")
                        {
                            adr = "O" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 15] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "O71";
                                Raschet(adr, WS);
                                WS.Cells[71, 15] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //16 столбец
                        if (sotrudniki6KAD.Vid_Obrazov == "культура и искусство")
                        {
                            adr = "P" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 16] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "P71";
                                Raschet(adr, WS);
                                WS.Cells[71, 16] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //17 столбец
                        if (sotrudniki6KAD.Vid_Obrazov == "информ.безопасность")
                        {
                            adr = "Q" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 17] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "Q71";
                                Raschet(adr, WS);
                                WS.Cells[71, 17] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //18 столбец
                        if (sotrudniki6KAD.Vid_Obrazov == "сфера обслуживания")
                        {
                            adr = "R" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 18] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "R71";
                                Raschet(adr, WS);
                                WS.Cells[71, 18] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //19 столбец
                        if (sotrudniki6KAD.Vid_Obrazov == "сельск.и рыбное хоз.")
                        {
                            adr = "S" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 19] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "S71";
                                Raschet(adr, WS);
                                WS.Cells[71, 19] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //20 столбец
                        if (sotrudniki6KAD.Vid_Obrazov == "здравоохранение")
                        {
                            adr = "T" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 20] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "T71";
                                Raschet(adr, WS);
                                WS.Cells[71, 20] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //21 столбец
                        if (!((sotrudniki6KAD.Vid_Obrazov == "здравоохранение") || (sotrudniki6KAD.Vid_Obrazov == "физ.-мат.и ест.науки") || (sotrudniki6KAD.Vid_Obrazov == "гум.и социал.науки") || (sotrudniki6KAD.Vid_Obrazov == "юриспруденция") || (sotrudniki6KAD.Vid_Obrazov == "правоохр.деят.") || (sotrudniki6KAD.Vid_Obrazov == "экономика и управ.") || (sotrudniki6KAD.Vid_Obrazov == "таможенное дело") || (sotrudniki6KAD.Vid_Obrazov == "гос.и мун.управление") || (sotrudniki6KAD.Vid_Obrazov == "образ.и педагогика") || (sotrudniki6KAD.Vid_Obrazov == "культура и искусство") || (sotrudniki6KAD.Vid_Obrazov == "сельск.и рыбное хоз.") || (sotrudniki6KAD.Vid_Obrazov == "сфера обслуживания") || (sotrudniki6KAD.Vid_Obrazov == "информ.безопасность")))
                        {
                            adr = "U" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 21] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "U71";
                                Raschet(adr, WS);
                                WS.Cells[71, 21] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //----------------------------------------
                        //ВОЗРАСТ
                        //до 30
                        if (Convert.ToInt32(sotrudniki6KAD.Vozrast) < 30)
                        {
                            adr = "V" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 22] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "V71";
                                Raschet(adr, WS);
                                WS.Cells[71, 22] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //30-39
                        if ((Convert.ToInt32(sotrudniki6KAD.Vozrast) <= 39) & (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 30))
                        {
                            adr = "W" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 23] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "W71";
                                Raschet(adr, WS);
                                WS.Cells[71, 23] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //40-49
                        if ((Convert.ToInt32(sotrudniki6KAD.Vozrast) <= 49) & (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 40))
                        {
                            adr = "X" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 24] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "X71";
                                Raschet(adr, WS);
                                WS.Cells[71, 24] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //50-59
                        if ((Convert.ToInt32(sotrudniki6KAD.Vozrast) <= 59) & (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 50))
                        {
                            adr = "Y" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 25] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "Y71";
                                Raschet(adr, WS);
                                WS.Cells[71, 25] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //более 60
                        if (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 60)
                        {
                            adr = "Z" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 26] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "Z71";
                                Raschet(adr, WS);
                                WS.Cells[71, 26] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //----------------------------

                        //--------ОТКУДА_НАЧАЛО--------
                        //колонка 27
                        if (sotrudniki6KAD.Otkyda == "из Вооруженных Сил")
                        {
                            adr = "AA" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 27] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AA71";
                                Raschet(adr, WS);
                                WS.Cells[71, 27] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //колонка 28
                        if (sotrudniki6KAD.Otkyda == "из ФСБ")
                        {
                            adr = "AB" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 28] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AB71";
                                Raschet(adr, WS);
                                WS.Cells[71, 28] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //колонка 29
                        if (sotrudniki6KAD.Otkyda == "из МВД")
                        {
                            adr = "AC" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 29] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AC71";
                                Raschet(adr, WS);
                                WS.Cells[71, 29] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //колонка 30
                        if (sotrudniki6KAD.Otkyda == "из ФСКН")
                        {
                            adr = "AD" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 30] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AD71";
                                Raschet(adr, WS);
                                WS.Cells[71, 30] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //колонка 31
                        if (sotrudniki6KAD.Otkyda == "из др.правоохр.орган")
                        {
                            adr = "AE" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 31] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AE71";
                                Raschet(adr, WS);
                                WS.Cells[71, 31] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //колонка 32
                        if (sotrudniki6KAD.Otkyda == "мол.спец. (РТА)")
                        {
                            adr = "AF" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 32] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AF71";
                                Raschet(adr, WS);
                                WS.Cells[71, 32] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //колонка 33
                        if (sotrudniki6KAD.Otkyda == "молодые специалисты")
                        {
                            adr = "AG" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 33] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AG71";
                                Raschet(adr, WS);
                                WS.Cells[71, 33] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //колонка 34
                        if ((sotrudniki6KAD.Otkyda == "другое") || (sotrudniki6KAD.Otkyda == "из др. орг-й") || (sotrudniki6KAD.Otkyda == "из гос.органа") || (sotrudniki6KAD.Otkyda == "из др.орг(служба труда)") || (sotrudniki6KAD.Otkyda == "из коммерческих стр."))
                        {
                            adr = "AH" + stroka;
                            Raschet(adr, WS);
                            WS.Cells[stroka, 34] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();

                            if (sotrudniki6KAD.Pol == "жен")
                            {

                                adr = "AH71";
                                Raschet(adr, WS);
                                WS.Cells[71, 34] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //--------ОТКУДА_КОНЕЦ--------
                    



                    WS.Range["G6"].Value = String.Format("за период с: {0} по: {1} | Дата формирования:{2}.{3}.{4}", Period_S, Period_Po, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year);
                        if (arrTM_TO == null)
                        {
                            arrTM_TO = "ЦТУ";
                        }
                    WS.Range["C22"].Value = String.Format(arrTM_TO);
                    for (int i = 0; i < 15; i++)

                    {
                        var ma = massiv.Kod_Organa_TO[1, i];
                        if (massiv.Kod_Organa_TO[1, i] == arrTM_TO)
                        {

                            WS.Range["O22"].Value = String.Format(massiv.Kod_Organa_TO[0, i]);
                        }

                    }
                }
                    if ((sotrudniki6KAD.Vid_Sluj == "Госслужащий") || (sotrudniki6KAD.Status_Slujb == "Проходит службу_ГС"))
                    {
                        WS = (Excel.Worksheet)WB.Sheets[2];
                        WS.Activate();
                        var strGOS = 11;
                        //foreach (string doljGos in massiv.GOS_Doljnost)
                        
                            if (sotrudniki6KAD.Gruppa_Doljnost == "Высшая группа должностей")
                        {
                            strGOS = 11;
                        }

                            if (sotrudniki6KAD.Gruppa_Doljnost == "Главная группа должностей")
                        {
                            strGOS = 12;
                        }
                            if (sotrudniki6KAD.Gruppa_Doljnost == "Ведущая группа должностей")
                        {
                            strGOS = 13;
                        }
                            if (sotrudniki6KAD.Gruppa_Doljnost == "Старшая группа должностей")
                        {
                            strGOS = 14;
                        }
                            if (sotrudniki6KAD.Gruppa_Doljnost == "Младшая группа должностей")
                        {
                            strGOS = 15;
                        }
                        //if (sotrudniki6KAD.Vid_Sluj == "Госслужащий")//первый зам рук.фтс росии
                        
                        string adr = "C" + strGOS;
                        //string Primechanie_Cells = (WS.Range[adr] as Range).Comment.Shape.TextFrame.Characters(Type.Missing, Type.Missing).Text;
                        Raschet(adr, WS);
                                               

                        WS.Cells[strGOS, 3] = Kol_vo;
                        string fio_primech = sotrudniki6KAD.FIO.ToString();
                        RaschetFio(adr, WS, fio_primech);
                        WS.Range[adr].ClearComments();         
                       WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        if (sotrudniki6KAD.Pol == "жен")
                                {

                                    adr = "C17";
                                    Raschet(adr, WS);
                                    WS.Cells[17, 3] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        }


                                //высшее
                                if ((sotrudniki6KAD.Obrazov == "ВО.бакалавриат") || (sotrudniki6KAD.Obrazov == "ВО.магистратура") || (sotrudniki6KAD.Obrazov == "ВО.специалитет") || (sotrudniki6KAD.Obrazov == "высшее") || (sotrudniki6KAD.Obrazov == "высшее проф."))
                                {

                                    adr = "D" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 4] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "D17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 4] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }

                                //среднее проф
                                if ((sotrudniki6KAD.Obrazov == "СО.проф") || (sotrudniki6KAD.Obrazov == "среднее проф."))
                                {
                                    adr = "F" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 6] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "F17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 6] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //среднее полное общее
                                ///*if (sotrudniki6KAD.Obrazov == "среднее общее")
                                //{
                                //    adr = "H" + strGOS;
                                //    Raschet(adr, WS);
                                //    WS.Cells[strGOS, 8] = Kol_vo;
                                //    if (sotrudniki6KAD.Pol == "жен")
                                //    {

                                //        adr = "H17";
                                //        Raschet(adr, WS);
                                //        WS.Cells[17, 8] = Kol_vo;
                                //    }
                                //}

                                //7 столбец
                                if (sotrudniki6KAD.Vid_Obrazov == "физ.-мат.и ест.науки")
                                {
                                    adr = "G" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 7] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "G17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 7] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //8 столбец
                                if ((sotrudniki6KAD.Vid_Obrazov == "гум.и социал.науки") || (sotrudniki6KAD.Vid_Obrazov == "юриспруденция") || (sotrudniki6KAD.Vid_Obrazov == "правоохр.деят."))
                                {
                                    adr = "H" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 8] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "H17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 8] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }

                                    //9 столбец
                                    if ((sotrudniki6KAD.Specialnost == "Юриспруденция") || (sotrudniki6KAD.Specialnost == "Правоох.деят."))
                                    {

                                        adr = "I" + strGOS;
                                        Raschet(adr, WS);
                                        WS.Cells[strGOS, 9] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (sotrudniki6KAD.Pol == "жен")
                                        {

                                            adr = "I17";
                                            Raschet(adr, WS);
                                            WS.Cells[17, 9] = Kol_vo;
                                    fio_primech = sotrudniki6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                                    }


                                }
                                //10 столбец
                                if ((sotrudniki6KAD.Vid_Obrazov == "экономика и управ.") || (sotrudniki6KAD.Vid_Obrazov == "таможенное дело") || (sotrudniki6KAD.Vid_Obrazov == "гос.и мун.управление"))
                                {

                                    adr = "J" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 10] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "J17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 10] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }

                                    //11 столбец
                                    if (sotrudniki6KAD.Specialnost == "Таможенное дело")
                                    {
                                        adr = "K" + strGOS;
                                        Raschet(adr, WS);
                                        WS.Cells[strGOS, 11] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (sotrudniki6KAD.Pol == "жен")
                                        {

                                            adr = "K17";
                                            Raschet(adr, WS);
                                            WS.Cells[17, 11] = Kol_vo;
                                    fio_primech = sotrudniki6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                                    }

                                    //12 столбец
                                    if (sotrudniki6KAD.Specialnost == "ГосМунУправление")
                                    {
                                        adr = "L" + strGOS;
                                        Raschet(adr, WS);
                                        WS.Cells[strGOS, 12] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (sotrudniki6KAD.Pol == "жен")
                                        {

                                            adr = "L17";
                                            Raschet(adr, WS);
                                            WS.Cells[17, 12] = Kol_vo;
                                    fio_primech = sotrudniki6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                                    }
                                }

                                //13 столбец
                                if (sotrudniki6KAD.Vid_Obrazov == "образ.и педагогика")
                                {
                                    adr = "M" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 13] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "M17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 13] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //14 столбец
                                if (sotrudniki6KAD.Vid_Obrazov == "культура и искусство")
                                {
                                    adr = "N" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 14] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "N17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 14] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //15 столбец
                                if (sotrudniki6KAD.Vid_Obrazov == "информ.безопасность")
                                {
                                    adr = "O" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 15] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "O17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 15] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //16 столбец
                                if (sotrudniki6KAD.Vid_Obrazov == "сфера обслуживания")
                                {
                                    adr = "P" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 16] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "P17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 16] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //17 столбец
                                if (sotrudniki6KAD.Vid_Obrazov == "сельск.и рыбное хоз.")
                                {
                                    adr = "Q" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 17] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "Q17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 17] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //18 столбец
                                if (sotrudniki6KAD.Vid_Obrazov == "здравоохранение")
                                {
                                    adr = "R" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 18] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "R17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 18] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }

                                //19 столбец
                                if (!((sotrudniki6KAD.Vid_Obrazov == "здравоохранение") || (sotrudniki6KAD.Vid_Obrazov == "физ.-мат.и ест.науки") || (sotrudniki6KAD.Vid_Obrazov == "гум.и социал.науки") || (sotrudniki6KAD.Vid_Obrazov == "юриспруденция") || (sotrudniki6KAD.Vid_Obrazov == "правоохр.деят.") || (sotrudniki6KAD.Vid_Obrazov == "экономика и управ.") || (sotrudniki6KAD.Vid_Obrazov == "таможенное дело") || (sotrudniki6KAD.Vid_Obrazov == "гос.и мун.управление") || (sotrudniki6KAD.Vid_Obrazov == "образ.и педагогика") || (sotrudniki6KAD.Vid_Obrazov == "культура и искусство") || (sotrudniki6KAD.Vid_Obrazov == "сельск.и рыбное хоз.") || (sotrudniki6KAD.Vid_Obrazov == "сфера обслуживания") || (sotrudniki6KAD.Vid_Obrazov == "информ.безопасность")))
                                {
                                    adr = "S" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 19] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "S17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 19] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //----------------------------------------
                                //ВОЗРАСТ
                                //до 30
                                if (Int32.Parse(sotrudniki6KAD.Vozrast) < 30)
                                {
                                    adr = "T" + strGOS;
                                    Raschet(adr, WS);
                            var fi =Convert.ToString(sotrudniki6KAD.FIO);
                                    WS.Cells[strGOS, 20] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "T17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 20] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }

                                //30-39
                                if ((Convert.ToInt32(sotrudniki6KAD.Vozrast) <= 39) & (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 30))
                                {
                                    adr = "U" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 21] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "U17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 21] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //40-49
                                if ((Convert.ToInt32(sotrudniki6KAD.Vozrast) <= 49) & (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 40))
                                {
                                    adr = "V" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 22] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "V17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 22] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //50-59
                                if ((Convert.ToInt32(sotrudniki6KAD.Vozrast) <= 59) & (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 50))
                                {
                                    adr = "W" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 23] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "W17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 23] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //более 60
                                if (Convert.ToInt32(sotrudniki6KAD.Vozrast) >= 60)
                                {
                                    adr = "X" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 24] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "X17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 24] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }

                                //----------------------------

                                //--------ОТКУДА_НАЧАЛО--------
                                //колонка 25
                                if (sotrudniki6KAD.Otkyda == "из Вооруженных Сил")
                                {
                                    adr = "Y" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 25] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "Y17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 25] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //колонка 26
                                if (sotrudniki6KAD.Otkyda == "из ФСБ")
                                {
                                    adr = "Z" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 26] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "Z17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 26] = Kol_vo;

                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //колонка 27
                                if (sotrudniki6KAD.Otkyda == "из МВД")
                                {
                                    adr = "AA" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 27] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "AA17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 27] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //колонка 28
                                if (sotrudniki6KAD.Otkyda == "из ФСКН")
                                {
                                    adr = "AB" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 28] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "AB17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 28] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //колонка 29
                                if (sotrudniki6KAD.Otkyda == "из др.правоохр.орган")
                                {
                                    adr = "AC" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 29] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "AC17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 29] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //колонка 30
                                if (sotrudniki6KAD.Otkyda == "мол.спец. (РТА)")
                                {
                                    adr = "AD" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 30] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "AD17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 30] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //колонка 31
                                if ((sotrudniki6KAD.Otkyda == "молодые специалисты") ||  (sotrudniki6KAD.Otkyda == "мол.спец. (др.уч.завед)"))
                                {
                                    adr = "AE" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 31] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "AE17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 31] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //колонка 32
                                if ((sotrudniki6KAD.Otkyda == "прочие ист.комплект.") || (sotrudniki6KAD.Otkyda == "другое") || (sotrudniki6KAD.Otkyda == "из др. орг-й") || (sotrudniki6KAD.Otkyda == "из гос.органа") || (sotrudniki6KAD.Otkyda == "из др.орг(служба труда)") || (sotrudniki6KAD.Otkyda == "из коммерческих стр."))
                                {
                                    adr = "AF" + strGOS;
                                    Raschet(adr, WS);
                                    WS.Cells[strGOS, 32] = Kol_vo;
                            fio_primech = sotrudniki6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();

                            if (sotrudniki6KAD.Pol == "жен")
                                    {

                                        adr = "AF17";
                                        Raschet(adr, WS);
                                        WS.Cells[17, 32] = Kol_vo;
                                fio_primech = sotrudniki6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                                }
                                //--------ОТКУДА_КОНЕЦ--------
                            


                        

                    }

                        WB.SaveAs(path_Save_CAD_6, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


                }
               
                
                #endregion


                   
                foreach (Yvolenn6KAD yvolenn6KAD in massiv._Yvolenn6KAD)
                {
                    // уволенные
                    if (yvolenn6KAD.Vid_Slujb == "Сотрудник")
                    {
                        WS = (Excel.Worksheet)WB.Sheets[3];
                        WS.Activate();

                        var strYvol = 9;
                        if (yvolenn6KAD.Mnemo_Dolj == "руководитель ФТС_ЦА") { strYvol = 9; }
                        if (yvolenn6KAD.Mnemo_Dolj == "перв.зам.рук.ФТС_ЦА") { strYvol = 10; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.рук.ФТС России_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "статс-секретарь_ЦА")) { strYvol = 11; }
                        if (yvolenn6KAD.Mnemo_Dolj == "нач.глав.управления_ЦА") { strYvol = 12; }
                        if (yvolenn6KAD.Mnemo_Dolj == "перв.зам.нач.ГУ_ЦА") { strYvol = 13; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.ГУ_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.гл.упр.-нач.секр.рук.ФТС_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.ГУ-глав.бух_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.ГУ-нач.отдела_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.ГУ-нач.служ_ЦА")) { strYvol = 14; }

                        if (yvolenn6KAD.Mnemo_Dolj == "нач.управления_ЦА") { strYvol = 15; }
                        if (yvolenn6KAD.Mnemo_Dolj == "перв.зам.нач.упр_ЦА") { strYvol = 16; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.управ_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.управ-нач.службы_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.управл.-нач.отдела_ЦА")) { strYvol = 17; }
                        if (yvolenn6KAD.Mnemo_Dolj == "гл.советник_ЦА") { strYvol = 18; }
                        if (yvolenn6KAD.Mnemo_Dolj == "советник.рук.ФТС_ЦА") { strYvol = 19; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "советник.нач.гл.упр_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "советник.нач.упр_ЦА")) { strYvol = 20; }

                        if ((yvolenn6KAD.Mnemo_Dolj == "пом.перв.зам.рук.ФТС_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "пом.зам.рук.ФТС_ЦА")) { strYvol = 21; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "нач.службы_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.главбух-нач.отдела_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.главбух.ФТС_ЦА")) { strYvol = 22; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы_НО_ЦА")) { strYvol = 23; }
                        if (yvolenn6KAD.Mnemo_Dolj == "нач. управления_РТУ") { strYvol = 25; }
                        if (yvolenn6KAD.Mnemo_Dolj == "перв.зам.нач.упр_РТУ") { strYvol = 26; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.управ_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.управ-нач.службы_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.управ-нач.отдела_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.управ-нач.там_РТУ")) { strYvol = 27; }

                        if (yvolenn6KAD.Mnemo_Dolj == "пом.нач.управл_РТУ") { strYvol = 28; }
                        if (yvolenn6KAD.Mnemo_Dolj == "нач.службы_РТУ") { strYvol = 29; }
                        if (yvolenn6KAD.Mnemo_Dolj == "перв.зам.нач.службы_РТУ") { strYvol = 30; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы-нач.отдела_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы-гл.бух_РТУ")) { strYvol = 31; }
                        if (yvolenn6KAD.Mnemo_Dolj == "нач.таможни_Т") { strYvol = 32; }
                        if (yvolenn6KAD.Mnemo_Dolj == "перв.зам.нач.там._Т") { strYvol = 33; }

                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.таможни_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.там.ПОД_Т") || (yvolenn6KAD.Mnemo_Dolj == "Зам.нач.там.по раб.с кадрами_Т") || (yvolenn6KAD.Mnemo_Dolj == "Зам.нач.там.по кадрам_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.там-нач.службы_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.там-нач.поста_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.там-нач.отдела_Т") || (yvolenn6KAD.Mnemo_Dolj == "Зам.нач.там. по ЭД_Т") || (yvolenn6KAD.Mnemo_Dolj == "Зам.нач.там. по тыл.обесп_Т")) { strYvol = 34; }
                        if (yvolenn6KAD.Mnemo_Dolj == "нач.служ_Т") { strYvol = 35; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы-нач.отдела_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы-нач.отдела-глав.бух_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.службы-глав.бух._Т")) { strYvol = 36; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "нач.поста_ТП") || (yvolenn6KAD.Mnemo_Dolj == "нач.поста-глав.бух_ТП")) { strYvol = 37; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.поста_ТП") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.поста-нач.отд._ТП") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.поста-нач.отделения_ТП") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.поста-гл.бухг_ТП")) { strYvol = 38; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "нач.отдела_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.отд-гл.бухгалтер_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.отд-зам.гл.бухгалтер_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.СОБРа_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_Т_ЛС") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_ТП") || (yvolenn6KAD.Mnemo_Dolj == "нач.отд-гл.бухгалтер_ТП") || (yvolenn6KAD.Mnemo_Dolj == "нач.отд-зам.гл.бухгалтер_ТП") || (yvolenn6KAD.Mnemo_Dolj == "Нач.СОБРа_ТП") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_РТУ_ПС") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_РТУ_БС") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_РТУ_ЛС") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_РТУ_ИТС") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_Т_М") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела-гл.бухг_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдела_РТА") || (yvolenn6KAD.Mnemo_Dolj == "нач.СОБРа_РТУ")) { strYvol = 39; }

                        if ((yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-гл.метрол_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-нач.отделен_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_Т_М") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_Т_ЛС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-зам.глав.бух_ТП") || (yvolenn6KAD.Mnemo_Dolj == "Зам.нач.СОБРа_ТП") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_ПС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_БС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_ЛС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ_ИТС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.СОБРа_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_БТС_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_РТА") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_Т_ИТС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-зам.глав.бух_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_БТС_Т") || (yvolenn6KAD.Mnemo_Dolj == "Зам.нач.отдела-нач.отделения_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.СОБРа_Т") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_СТС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела-командир_МТС") || (yvolenn6KAD.Mnemo_Dolj == "зам.нач.отдела_ТП")) { strYvol = 40; }
                        if ((yvolenn6KAD.Mnemo_Dolj == "нач.отделения_ЦА") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения_Т_М") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения-командир_СТС_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения-командир_МТС_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения-командир_БТС_Т") || (yvolenn6KAD.Mnemo_Dolj == "нач.отдел-командир_СТС") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения_ТП") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения_РТУ_ПС") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения_РТУ") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения_РТА") || (yvolenn6KAD.Mnemo_Dolj == "нач.отделения_РТУ_БС")) { strYvol = 41; }
                        if (yvolenn6KAD.Kat_Dolj == "Главные инспекторы") { strYvol = 42; }
                        if (yvolenn6KAD.Kat_Dolj == "Ведущие инспекторы") { strYvol = 43; }
                        if (yvolenn6KAD.Kat_Dolj == "Старшие инспекторы") { strYvol = 44; }
                        if (yvolenn6KAD.Kat_Dolj == "Инспекторы") { strYvol = 45; }
                        if (yvolenn6KAD.Kat_Dolj == "Младшие инспекторы") { strYvol = 46; }


                        if (strYvol==9)
                        { strYvol = 0; }

                        string adr = "C" + strYvol;
                        Local_ADR_TEST = adr.ToString();
                        Raschet(adr, WS);
                        WS.Cells[strYvol, 3] = Kol_vo;
                        var fio_primech = yvolenn6KAD.FIO.ToString();
                        RaschetFio(adr, WS, fio_primech);
                        WS.Range[adr].ClearComments();
                        WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        if (yvolenn6KAD.POL == "жен")
                        {

                            adr = "C48";
                            Raschet(adr, WS);
                            WS.Cells[48, 3] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        }


                        //высшее
                        if ((yvolenn6KAD.Vid_Obraz == "ВО.бакалавриат") || (yvolenn6KAD.Vid_Obraz == "ВО.магистратура") || (yvolenn6KAD.Vid_Obraz == "ВО.специалитет") || (yvolenn6KAD.Vid_Obraz == "высшее") || (yvolenn6KAD.Vid_Obraz == "высшее проф."))
                        {

                            adr = "D" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 4] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "D48";
                                Raschet(adr, WS);
                                WS.Cells[48, 4] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (yvolenn6KAD.Vid_Obraz == "незаконч.высшее")
                        {
                            adr = "E" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 5] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "E48";
                                Raschet(adr, WS);
                                WS.Cells[48, 5] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }


                        //среднее проф
                        if ((yvolenn6KAD.Vid_Obraz == "СО.проф") || (yvolenn6KAD.Vid_Obraz == "среднее проф."))
                        {
                            adr = "F" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 6] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "F48";
                                Raschet(adr, WS);
                                WS.Cells[48, 6] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if (yvolenn6KAD.Vid_Obraz == "начальное проф.")
                        {
                            adr = "G" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 7] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "G48";
                                Raschet(adr, WS);
                                WS.Cells[48, 7] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //среднее полное общее
                        if (yvolenn6KAD.Vid_Obraz == "среднее общее")
                        {
                            adr = "H" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 8] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "H48";
                                Raschet(adr, WS);
                                WS.Cells[48, 8] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //ВОЗРАСТ
                        //до 30
                        if (Convert.ToInt32(yvolenn6KAD.Vozrast) < 30)
                        {
                            adr = "I" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 9] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "I48";
                                Raschet(adr, WS);
                                WS.Cells[48, 9] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //30-39
                        if ((Convert.ToInt32(yvolenn6KAD.Vozrast) <= 39) & (Convert.ToInt32(yvolenn6KAD.Vozrast) >= 30))
                        {
                            adr = "J" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 10] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "J48";
                                Raschet(adr, WS);
                                WS.Cells[48, 10] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //40-49
                        if ((Convert.ToInt32(yvolenn6KAD.Vozrast) <= 49) & (Convert.ToInt32(yvolenn6KAD.Vozrast) >= 40))
                        {
                            adr = "K" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 11] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "K48";
                                Raschet(adr, WS);
                                WS.Cells[48, 11] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //50-59
                        if ((Convert.ToInt32(yvolenn6KAD.Vozrast) <= 59) & (Convert.ToInt32(yvolenn6KAD.Vozrast) >= 50))
                        {
                            adr = "L" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 12] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "L48";
                                Raschet(adr, WS);
                                WS.Cells[48, 12] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //более 60
                        if ((Convert.ToInt32(yvolenn6KAD.Vozrast) >= 60) & (Convert.ToInt32(yvolenn6KAD.Vozrast) <= 65))
                        {
                            adr = "M" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 13] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "M48";
                                Raschet(adr, WS);
                                WS.Cells[48, 13] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //----------------------------

                        //-- СТАЖ----------НАЧАЛО
                        var DatNachSluj = DateTime.Parse(yvolenn6KAD.Date_Nach.ToString());
                        var DatFinlSluj = DateTime.Parse(yvolenn6KAD.Date_Fin.ToString());

                        var Raznica = (DatFinlSluj - DatNachSluj).Days;

                        if (Raznica < 365) //до года
                        {
                            adr = "N" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 14] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "N48";
                                Raschet(adr, WS);
                                WS.Cells[48, 14] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((Raznica >= 365) & (Raznica < 1825)) // 1-5 лет
                        {
                            adr = "O" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 15] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "O48";
                                Raschet(adr, WS);
                                WS.Cells[48, 15] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((Raznica >= 1825) & (Raznica < 3650)) // 5-10лет
                        {
                            adr = "P" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 16] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "P48";
                                Raschet(adr, WS);
                                WS.Cells[48, 16] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((Raznica >= 3650) & (Raznica < 5475)) // 10-15 лет
                        {
                            adr = "Q" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 17] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "Q48";
                                Raschet(adr, WS);
                                WS.Cells[48, 17] = Kol_vo; fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (Raznica >= 5475) // больше 15 лет
                        {
                            adr = "R" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 18] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "R48";
                                Raschet(adr, WS);
                                WS.Cells[48, 18] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //-- СТАЖ----------КОНЕЦ

                        //-- Причина Увол----------НАЧАЛО
                        if ((yvolenn6KAD.Motiv == "по собственному желанию до истечения срока контракта") || (yvolenn6KAD.Motiv == "собственное желание")  || (yvolenn6KAD.Motiv == "по собственному желанию"))
                        {
                            adr = "S" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 19] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "S48";
                                Raschet(adr, WS);
                                WS.Cells[48, 19] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if (yvolenn6KAD.Motiv == "окончание срока службы, предусмотренного контрактом")
                        {
                            adr = "T" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 20] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "T48";
                                Raschet(adr, WS);
                                WS.Cells[48, 20] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((yvolenn6KAD.Motiv == "по выслуге срока службы, дающего право на пенсию") || (yvolenn6KAD.Motiv == "по выслуге срока службы") || (yvolenn6KAD.Motiv == "пенсия"))
                        {
                            adr = "U" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 21] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "U48";
                                Raschet(adr, WS);
                                WS.Cells[48, 21] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if ((yvolenn6KAD.Motiv == "по достижению предельного возраста") || (yvolenn6KAD.Motiv == "достижение предельного возраста"))
                        {
                            adr = "V" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 22] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "V48";
                                Raschet(adr, WS);
                                WS.Cells[48, 22] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((yvolenn6KAD.Motiv == "орг-штатные мероприятия") || (yvolenn6KAD.Motiv == "в связи с организационно-штатными мероприятиями") || (yvolenn6KAD.Motiv == "по сокращению штатов") || (yvolenn6KAD.Motiv == "сокращение штатов"))
                        {
                            adr = "W" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 23] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "W48";
                                Raschet(adr, WS);
                                WS.Cells[48, 23] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if (yvolenn6KAD.Motiv == "по результатам аттестации")
                        {
                            adr = "X" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 24] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "X48";
                                Raschet(adr, WS);
                                WS.Cells[48, 24] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((yvolenn6KAD.Motiv == "по состоянию здоровья34BA") || (yvolenn6KAD.Motiv == "по причине болезни"))
                        {
                            adr = "Y" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 25] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "Y48";
                                Raschet(adr, WS);
                                WS.Cells[48, 25] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (yvolenn6KAD.Motiv == "за нарушения")
                        {
                            adr = "Z" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 26] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "Z48";
                                Raschet(adr, WS);
                                WS.Cells[48, 26] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (!((yvolenn6KAD.Motiv == "за нарушения") || (yvolenn6KAD.Motiv == "по состоянию здоровья34BA") || (yvolenn6KAD.Motiv == "по причине болезни") || (yvolenn6KAD.Motiv == "по результатам аттестации") || (yvolenn6KAD.Motiv == "орг-штатные мероприятия") || (yvolenn6KAD.Motiv == "в связи с организационно-штатными мероприятиями") || (yvolenn6KAD.Motiv == "по сокращению штатов") || (yvolenn6KAD.Motiv == "сокращение штатов") || (yvolenn6KAD.Motiv == "по достижению предельного возраста") || (yvolenn6KAD.Motiv == "достижение предельного возраста") || (yvolenn6KAD.Motiv == "по выслуге срока службы, дающего право на пенсию") || (yvolenn6KAD.Motiv == "по выслуге срока службы") || (yvolenn6KAD.Motiv == "пенсия") || (yvolenn6KAD.Motiv == "по собственному желанию до истечения срока контракта") || (yvolenn6KAD.Motiv == "собственное желание") || (yvolenn6KAD.Motiv == "по собственному желанию") || (yvolenn6KAD.Motiv == "окончание срока службы, предусмотренного контрактом")))
                        {
                            adr = "AB" + strYvol;
                            Raschet(adr, WS);
                            WS.Cells[strYvol, 28] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "AB48";
                                Raschet(adr, WS);
                                WS.Cells[48, 28] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //-- Причина Увол----------КОНЕЦ


                    }

                    if (yvolenn6KAD.Vid_Slujb == "Госслужащий")
                    {
                        WS = (Excel.Worksheet)WB.Sheets[4];
                        WS.Activate();
                        var strokaYvolGos = 9;
                        if (yvolenn6KAD.Gruppa_Dolj == "Высшая группа должностей")    {  strokaYvolGos = 9;   }

                        if (yvolenn6KAD.Gruppa_Dolj == "Главная группа должностей")  { strokaYvolGos = 10;   }
                        if (yvolenn6KAD.Gruppa_Dolj == "Ведущая группа должностей")         { strokaYvolGos = 11; }
                        if (yvolenn6KAD.Gruppa_Dolj == "Старшая группа должностей")      { strokaYvolGos = 12; }
                        if (yvolenn6KAD.Gruppa_Dolj == "Младшая группа должностей")    { strokaYvolGos = 13; }

                        if (strokaYvolGos == 9)
                        { strokaYvolGos = 0; }

                        string adr = "C" + strokaYvolGos;
                        Local_ADR_TEST = adr.ToString();
                        Raschet(adr, WS);
                        WS.Cells[strokaYvolGos, 3] = Kol_vo;
                        var fio_primech = yvolenn6KAD.FIO.ToString();
                        RaschetFio(adr, WS, fio_primech);
                        WS.Range[adr].ClearComments();
                        WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        if (yvolenn6KAD.POL == "жен")
                        {

                            adr = "C15";
                            Raschet(adr, WS);
                            WS.Cells[15, 3] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                        }


                        //высшее
                        if ((yvolenn6KAD.Vid_Obraz == "ВО.бакалавриат") || (yvolenn6KAD.Vid_Obraz == "ВО.магистратура") || (yvolenn6KAD.Vid_Obraz == "ВО.специалитет") || (yvolenn6KAD.Vid_Obraz == "высшее") || (yvolenn6KAD.Vid_Obraz == "высшее проф."))
                        {

                            adr = "D" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 4] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "D15";
                                Raschet(adr, WS);
                                WS.Cells[15, 4] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (yvolenn6KAD.Vid_Obraz == "незаконч.высшее")
                        {
                            adr = "E" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 5] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "E15";
                                Raschet(adr, WS);
                                WS.Cells[15, 5] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }


                        //среднее проф
                        if ((yvolenn6KAD.Vid_Obraz == "СО.проф") || (yvolenn6KAD.Vid_Obraz == "среднее проф."))
                        {
                            adr = "F" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 6] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "F15";
                                Raschet(adr, WS);
                                WS.Cells[15, 6] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if (yvolenn6KAD.Vid_Obraz == "начальное проф.")
                        {
                            adr = "G" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 7] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "G15";
                                Raschet(adr, WS);
                                WS.Cells[15, 7] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //среднее полное общее
                        if (yvolenn6KAD.Vid_Obraz == "среднее общее")
                        {
                            adr = "H" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 8] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "H15";
                                Raschet(adr, WS);
                                WS.Cells[15, 8] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //ВОЗРАСТ
                        if (yvolenn6KAD.Vozrast != "")
                        {
                            //до 30
                            if (Convert.ToInt32(yvolenn6KAD.Vozrast) < 30)
                            {
                                adr = "I" + strokaYvolGos;
                                Raschet(adr, WS);
                                WS.Cells[strokaYvolGos, 9] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (yvolenn6KAD.POL == "жен")
                                {

                                    adr = "I15";
                                    Raschet(adr, WS);
                                    WS.Cells[15, 9] = Kol_vo;
                                    fio_primech = yvolenn6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }

                            //30-39
                            if ((Convert.ToInt32(yvolenn6KAD.Vozrast) <= 39) & (Convert.ToInt32(yvolenn6KAD.Vozrast) >= 30))
                            {
                                adr = "J" + strokaYvolGos;
                                Raschet(adr, WS);
                                WS.Cells[strokaYvolGos, 10] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (yvolenn6KAD.POL == "жен")
                                {

                                    adr = "J15";
                                    Raschet(adr, WS);
                                    WS.Cells[15, 10] = Kol_vo;
                                    fio_primech = yvolenn6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }
                            //40-49
                            if ((Convert.ToInt32(yvolenn6KAD.Vozrast) <= 49) & (Convert.ToInt32(yvolenn6KAD.Vozrast) >= 40))
                            {
                                adr = "K" + strokaYvolGos;
                                Raschet(adr, WS);
                                WS.Cells[strokaYvolGos, 11] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (yvolenn6KAD.POL == "жен")
                                {

                                    adr = "K15";
                                    Raschet(adr, WS);
                                    WS.Cells[15, 11] = Kol_vo;
                                    fio_primech = yvolenn6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }
                            //50-59
                            if ((Convert.ToInt32(yvolenn6KAD.Vozrast) <= 59) & (Convert.ToInt32(yvolenn6KAD.Vozrast) >= 50))
                            {
                                adr = "L" + strokaYvolGos;
                                Raschet(adr, WS);
                                WS.Cells[strokaYvolGos, 12] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (yvolenn6KAD.POL == "жен")
                                {

                                    adr = "L15";
                                    Raschet(adr, WS);
                                    WS.Cells[15, 12] = Kol_vo;
                                    fio_primech = yvolenn6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }
                            //более 60
                            if ((Convert.ToInt32(yvolenn6KAD.Vozrast) >= 60) & (Convert.ToInt32(yvolenn6KAD.Vozrast) <= 65))
                            {
                                adr = "M" + strokaYvolGos;
                                Raschet(adr, WS);
                                WS.Cells[strokaYvolGos, 13] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                if (yvolenn6KAD.POL == "жен")
                                {

                                    adr = "M15";
                                    Raschet(adr, WS);
                                    WS.Cells[15, 13] = Kol_vo;
                                    fio_primech = yvolenn6KAD.FIO.ToString();
                                    RaschetFio(adr, WS, fio_primech);
                                    WS.Range[adr].ClearComments();
                                    WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                                }
                            }
                        }
                        //----------------------------

                        //-- СТАЖ----------НАЧАЛО
                        var DatNachSluj = DateTime.Parse(yvolenn6KAD.Date_Nach.ToString());
                        var DatFinlSluj = DateTime.Parse(yvolenn6KAD.Date_Fin.ToString());

                        var Raznica = (DatFinlSluj - DatNachSluj).Days;

                        if (Raznica < 365) //до года
                        {
                            adr = "N" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 14] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "N15";
                                Raschet(adr, WS);
                                WS.Cells[15, 14] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((Raznica >= 365) & (Raznica < 1825)) // 1-5 лет
                        {
                            adr = "O" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 15] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "O15";
                                Raschet(adr, WS);
                                WS.Cells[15, 15] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((Raznica >= 1825) & (Raznica < 3650)) // 5-10лет
                        {
                            adr = "P" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 16] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "P15";
                                Raschet(adr, WS);
                                WS.Cells[15, 16] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((Raznica >= 3650) & (Raznica < 5475)) // 10-15 лет
                        {
                            adr = "Q" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 17] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "Q15";
                                Raschet(adr, WS);
                                WS.Cells[15, 17] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (Raznica >= 5475) // больше 15 лет
                        {
                            adr = "R" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 18] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "R15";
                                Raschet(adr, WS);
                                WS.Cells[15, 18] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        //-- СТАЖ----------КОНЕЦ

                        //-- Причина Увол----------НАЧАЛО
                        if ((yvolenn6KAD.Motiv == "по собственному желанию до истечения срока контракта") || (yvolenn6KAD.Motiv == "собственное желание") || (yvolenn6KAD.Motiv == "по собственному желанию"))
                        {
                            adr = "S" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 19] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "S15";
                                Raschet(adr, WS);
                                WS.Cells[15, 19] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if (yvolenn6KAD.Motiv == "окончание срока службы, предусмотренного контрактом")
                        {
                            adr = "T" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 20] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "T15";
                                Raschet(adr, WS);
                                WS.Cells[15, 20] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((yvolenn6KAD.Motiv == "по выслуге срока службы, дающего право на пенсию") || (yvolenn6KAD.Motiv == "по выслуге срока службы") || (yvolenn6KAD.Motiv == "пенсия"))
                        {
                            adr = "U" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 21] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "U15";
                                Raschet(adr, WS);
                                WS.Cells[15, 21] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if ((yvolenn6KAD.Motiv == "по достижению предельного возраста") || (yvolenn6KAD.Motiv == "достижение предельного возраста"))
                        {
                            adr = "V" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 22] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "V15";
                                Raschet(adr, WS);
                                WS.Cells[15, 22] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((yvolenn6KAD.Motiv == "орг-штатные мероприятия") || (yvolenn6KAD.Motiv == "в связи с организационно-штатными мероприятиями") || (yvolenn6KAD.Motiv == "по сокращению штатов") || (yvolenn6KAD.Motiv == "сокращение штатов"))
                        {
                            adr = "W" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 23] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "W15";
                                Raschet(adr, WS);
                                WS.Cells[15, 23] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        if (yvolenn6KAD.Motiv == "по результатам аттестации")
                        {
                            adr = "X" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 24] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "X15";
                                Raschet(adr, WS);
                                WS.Cells[15, 24] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if ((yvolenn6KAD.Motiv == "по состоянию здоровья34BA") || (yvolenn6KAD.Motiv == "по причине болезни"))
                        {
                            adr = "Y" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 25] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "Y15";
                                Raschet(adr, WS);
                                WS.Cells[15, 25] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (yvolenn6KAD.Motiv == "за нарушения")
                        {
                            adr = "Z" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 26] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "Z15";
                                Raschet(adr, WS);
                                WS.Cells[15, 26] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }
                        if (!((yvolenn6KAD.Motiv == "за нарушения") || (yvolenn6KAD.Motiv == "по состоянию здоровья34BA") || (yvolenn6KAD.Motiv == "по причине болезни") || (yvolenn6KAD.Motiv == "по результатам аттестации") || (yvolenn6KAD.Motiv == "орг-штатные мероприятия") || (yvolenn6KAD.Motiv == "в связи с организационно-штатными мероприятиями") || (yvolenn6KAD.Motiv == "по сокращению штатов") || (yvolenn6KAD.Motiv == "сокращение штатов") || (yvolenn6KAD.Motiv == "по достижению предельного возраста") || (yvolenn6KAD.Motiv == "достижение предельного возраста") || (yvolenn6KAD.Motiv == "по выслуге срока службы, дающего право на пенсию") || (yvolenn6KAD.Motiv == "по выслуге срока службы") || (yvolenn6KAD.Motiv == "пенсия") || (yvolenn6KAD.Motiv == "по собственному желанию до истечения срока контракта") || (yvolenn6KAD.Motiv == "собственное желание") || (yvolenn6KAD.Motiv == "по собственному желанию") || (yvolenn6KAD.Motiv == "окончание срока службы, предусмотренного контрактом")))
                        {
                            adr = "AB" + strokaYvolGos;
                            Raschet(adr, WS);
                            WS.Cells[strokaYvolGos, 28] = Kol_vo;
                            fio_primech = yvolenn6KAD.FIO.ToString();
                            RaschetFio(adr, WS, fio_primech);
                            WS.Range[adr].ClearComments();
                            WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            if (yvolenn6KAD.POL == "жен")
                            {

                                adr = "AB15";
                                Raschet(adr, WS);
                                WS.Cells[15, 28] = Kol_vo;
                                fio_primech = yvolenn6KAD.FIO.ToString();
                                RaschetFio(adr, WS, fio_primech);
                                WS.Range[adr].ClearComments();
                                WS.Range[adr].AddComment(Primechanie_Cells).ToString();
                            }
                        }

                        //-- Причина Увол----------КОНЕЦ




                    }
                    WB.SaveAs(path_Save_CAD_6, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                }


                WB.Close(true);
                ObjWorkExcel.Quit();



            }

            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + Ex.StackTrace + "Перем:" + WS + "Адрес:" + Local_ADR_TEST);
            }



            return 1;
        }

        private int Raschet(string adr, Worksheet WS)
        {
            
            string kol_vo_parse = Convert.ToString(WS.Range[adr].Value);
            
            //Kol_vo=0;
            
            if (kol_vo_parse == "")
            {
                Kol_vo = 0;
            }
            else
            {
                Kol_vo = Convert.ToInt32(kol_vo_parse);
            }
            Kol_vo++;

           

            return Kol_vo;
        }

        private string RaschetFio(string adr, Worksheet WS, string fio_primech)
        {
            Object missing = Type.Missing;
            string primech = "";

            if (WS.Range[adr].Comment == null)
            {
                primech = " ";
            }
            if (WS.Range[adr].Comment != null)
            {
                primech = WS.Range[adr].Comment.Shape.TextFrame.Characters(missing, missing).Text;

            }
            Primechanie_Cells = primech + fio_primech ;

            return Primechanie_Cells;
        }

        private void PeriodS_DateSelected(object sender, DateRangeEventArgs e)
        {
            Period_S =e.End.ToString("yyyy-MM-dd");
            groupBox2.Text = "C: " + Period_S;
        }

        private void PeriodPo_DateSelected(object sender, DateRangeEventArgs e)
        {
            Period_Po = e.End.ToString("yyyy-MM-dd");
            groupBox3.Text = "ПО: " + Period_Po;
        }
    }
}