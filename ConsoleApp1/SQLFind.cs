using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace Vehicle.Helper
{
    
    internal class SQLFind
    {
        private Connection Connection= new Connection();
        public DataTable SQLSearch(string url) 
        {
            try
            {
                Connection.Conn();
                OracleDataAdapter mAdapter = new OracleDataAdapter(url, Connection.con);
                DataTable dt = new DataTable();
                mAdapter.Fill(dt);
                Connection.Conn_close();
                return dt;
            }
            catch(Exception ex)
            {
                Console.WriteLine("存储异常0"+ex);
                return null;
            }
        }
        public DataTable SQLSearch1(string url)
        {
            try
            {
                Connection.Conn1();
                OracleDataAdapter mAdapter = new OracleDataAdapter(url, Connection.con1);
                DataTable dt = new DataTable();
                mAdapter.Fill(dt);
                Connection.Conn1_close();
                return dt;
            }
            catch (Exception ex)
            {
                Console.WriteLine("存储异常1"+ex);
                return null;
            }
        }


    }
}
