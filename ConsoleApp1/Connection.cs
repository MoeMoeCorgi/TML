using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;
using System.Drawing;
using System.Security.Policy;

namespace Vehicle.Helper
{
    
    public class Connection
    {
        public OracleConnection con;
        public OracleConnection con1;
        private static string connString1 = "User Id=hzaf;Password=hzaf123;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.19.33.103)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=vidm)))";
        private static string connString = "User Id=MESPRD;Password=MESPRD;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.19.30.72)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=mhdb)))";
        public void Conn()
        {
            try
            {
                
                using (con)
                {
                    con = new OracleConnection(connString);
                    con.ConnectionString = connString;
                    con.Open();
                }
                Console.WriteLine("打开数据库0");
            }
            catch
            {
                Console.WriteLine("异常数据库0");
            }
 

        }
        public void Conn1() {
            try
            {
                
                using (con1)
                {
                    con1 = new OracleConnection(connString);
                    con1.ConnectionString = connString1;
                    con1.Open();
                }
                Console.WriteLine("打开数据库1");
            }
            catch
            {
                Console.WriteLine("异常数据库1");
            }


        }
        public void Conn_close()
        {
            con.Close();
            Console.WriteLine("关闭数据库0");
        }
        public void Conn1_close() {
            con1.Close();
            Console.WriteLine("关闭数据库1");
        }
        

    }
}
