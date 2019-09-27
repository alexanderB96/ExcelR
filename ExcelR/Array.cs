using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq;

namespace ExcelR
{

    #region  
    //для штат-факт
    public class Otdel
    {
        public string DEPART_DISP;
        public int PRN;
        public string NAME;
        public string NAME_NOM;
        public int RN;
        public string CODE;
        public int urlev;

    }


    public class Slujba
    {
        public string dep;
        public int urlev;
        public string NAME;
        public string NAME_NOM;
        public string SHORTNAME_NOM;
        public string CODE;
        public int RN;
        public string DEPART_DISP;

    }
    public class Dolj
    {
        public string MAIN_POST;
        public string MAIN_UPRAV;
        public int MAIN_KOLSHTAT;
        public string MAIN_PIPEC;
        public string EXEC;
        public int CLNKOD;
        public string SLVID;
        public int RN;
        public string DEP_CODE;
        public string AGN_RN;
        public string CLNZVAN;
        public int CRN;

    }

    public class Svedenia
    {
        public string SOSTAV;
        public string FIO;
        public string Numb;
        public string CRN;
        public string DEP_CODE;
        public string CLNZVAN;
        public string CLNKAD;
        public string SLVID;

    }

    public class SlujDolj
    {
        public string MAIN_POST;
        public string MAIN_UPRAV;
        public double MAIN_KOLSHTAT;
        public string MAIN_PIPEC;
        public string EXEC;
        public int CLNKOD;
        public string SLVID;
        public int RN;
        public string DEP_CODE;
        public string AGN_RN;
        public string CLNZVAN;
        public int CRN;
        public string CLNCODE;
        public string CLNGROUP;

    }

    public class SotrudnikiNizBlok
    {
        public string MAIN_POST;
        public string MAIN_UPRAV;
        public double MAIN_KOLSHTAT;
        public string MAIN_PIPEC;
        public string EXEC;
        public int CLNKOD;
        public string SLVID;
        public int RN;
        public string DEP_CODE;
        public string AGN_RN;
        public string CLNZVAN;
        public int CRN;
        public string CLNCODE;
        public string CLNGROUP;
    }

    public class RabotnikiBlok
    {
        public string Dolj;
        public string VidSljub;
        public string NUMB;
        public string CatDolj;
    }

    public class RabotnikiBlokFact
    {
        public string UPRAV;
        public string KOLSHTAT;
        public string PIPEC;
        public string EXEC;
        public string SLVID;
        public string DEP_CODE;
        public string AGN_RN;
        public string NUMB;
    }

    #endregion

    //для 6 КАД

    public class Sotrudniki6KAD
    {
        public string TO;
        public string Mnemo_TO;
        public string Verhnii_Level;
        public string FIO;
        public string Pol;
        public string Vozrast;
        public string Doljnost;
        public string Gruppa_Doljnost;
        public string Status_Slujb;
        public string Vid_Sluj;
        public string Data_Naznachenia;
        public string Mnemo_Dolj;
        public string Kat_Dolj;
        public string Obrazov;
        public string Vid_Obrazov;
        public string Specialnost;
        public string Date_Priema;
        public string Otkyda;

        //public string MAIN_CLNKAD;
       // public string Nomer_Agenta;
        
        

    }

    public class Yvolenn6KAD
    {
        public string Date_Yvol;
        public string FIO;
        public string Katalog;
        public string TO;
        public string Agen_RN;
        public string Vozrast;
        public string POL;
        public string Date_Nach;
        public string Date_Fin;
        public string Vid_Slujb;
        public string Vid_Obraz;
        public string Motiv;
        public string Mnemo_Dolj;
        public string Gruppa_Dolj;
        public string Kat_Dolj;
    }
    class MassivDannyh
    {
        public int company = 39154;

        public string[,] RN = { 
            { "222318", "221665", "221668", "221671", "221674", "221683", "221689", "221692", "451367509", "221737", "221719", "221725", "221728", "221652", "221734" },
            { "Аппарат Центрального таможенного управления","Белгородская таможня","Брянская таможня","Владимирская таможня","Воронежская таможня","Калужская таможня","Курская таможня","Липецкая таможня","Московская таможня","Приокский тыловой таможенный пост","Смоленская  таможня","Тверская  таможня","Тульская  таможня","Центральная оперативная таможня","Ярославская таможня",}
                            };

        public string[,] TM =  new string[,]{
            { "222318", "221665", "221668", "221671", "221674", "221683", "221689", "221692", "451367509", "221737", "221719", "221725", "221728", "221652", "221734" },
            { "Аппарат Центрального таможенного управления","Белгородская таможня","Брянская таможня","Владимирская таможня","Воронежская таможня","Калужская таможня","Курская таможня","Липецкая таможня","Московская таможня","Приокский тыловой таможенный пост","Смоленская  таможня","Тверская  таможня","Тульская  таможня","Центральная оперативная таможня","Ярославская таможня"},
            { "451832872 ","452074154","451962607","452266993","452220115","452359246","451775958","452422329","451529019","452948795","452869667","452342520","452343330","453189635","452263643"},
           {                                    "10100000",             "10101000",        "10102000",           "10103000",           "10104000",        "10106000",        "10108000",         "10109000",         "10129000",                         "10120010",           "10113000",          "10115000",        "10116000",                       "10119000",           "10117000" },

        };

        public string[,] Name_Kolonki = new string[,]
        {
            {  "0","1","2","3","4","  5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34" },
            {"nol","A","B","C", "D", "E", "F", "G", "H", "I", "J",  "K",  "L",  "M",  "N",  "O",  "P",  "Q",  "R",  "S",  "T",  "U",  "V",  "W",  "X", "Y",  "Z",  "AA", "AB",  "AC", "AD", "AE", "AF", "AG", "AH" }
        };

        public string[,] Kod_Organa_TO = new string[,]
        {
            {                                    "10100000",             "10101000",        "10102000",           "10103000",           "10104000",        "10106000",        "10108000",         "10109000",         "10129000",                         "10120010",           "10113000",          "10115000",        "10116000",                       "10119000",           "10117000" },
            { "Аппарат Центрального таможенного управления","Белгородская таможня","Брянская таможня","Владимирская таможня","Воронежская таможня","Калужская таможня","Курская таможня","Липецкая таможня","Московская таможня","Приокский тыловой таможенный пост","Смоленская  таможня","Тверская  таможня","Тульская  таможня","Центральная оперативная таможня","Ярославская таможня",}
                            };
        public string[] GOS_Doljnost = new string[]
        {
            "Высшая группа должностей", "Главная группа должностей","Ведущая группа должностей","Старшая группа должностей","Младшая группа должностей" 
        };

        public List<Otdel> _Otdel = new List<Otdel>();
        public List<Slujba> _Slujba = new List<Slujba>();
        public List<Dolj> _Dolj = new List<Dolj>();
        public List<SlujDolj> _SlujDolj = new List<SlujDolj>();
        public List<Svedenia> _Svedenia = new List<Svedenia>();
        public List<SotrudnikiNizBlok> _SotrudnikiNizBlok = new List<SotrudnikiNizBlok>();
        public List<RabotnikiBlok> _RabotnikiBlok = new List<RabotnikiBlok>();
        public List<RabotnikiBlokFact> _RabotnikiBlokFact = new List<RabotnikiBlokFact>();


        public List<Sotrudniki6KAD> _Sotrudniki6KAD = new List<Sotrudniki6KAD>();
        public List<Yvolenn6KAD> _Yvolenn6KAD = new List<Yvolenn6KAD>();



    }

    
}
