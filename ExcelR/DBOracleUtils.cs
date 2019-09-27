using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelR
{
    class DBOracleUtils
    {
        public  OracleConnection GetDBConnection()
        {
            string host;
            int port;
            string sid;

            host = "10.21.21.24";
            port = 1526;
            sid = "KADR";


           // string connString = "Data Source = (DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST =" + host + ")(PORT = " + port + "))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = " + sid + "))); Password = " + password + "; User ID = " + user;
            string connString = "Data Source=(DESCRIPTION=" + "(ADDRESS_LIST=" + "(ADDRESS=" + "(PROTOCOL=TCP)" + "(HOST=10.21.21.24)" + "(PORT=1526)" + ")" + ")" + "(CONNECT_DATA=" + "(SERVER=DEDICATED)" + "(SERVICE_NAME=KADR)" + ")" + ");" + "User Id=CTU_NESTERKINA;Password=A845510;";

            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = connString;
            return conn;
        }

    }
}
