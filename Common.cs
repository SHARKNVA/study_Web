using System;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using System.Data.OleDb;

namespace WebApplication1.App_Start
{
    public class Common
    {
        public static string Connect_UserSQL
        {
            get
            {
                var connSettings = ConfigurationManager.ConnectionStrings["ConnectionString_User"];

                try {
                    if (connSettings == null)
                    {
                        throw new InvalidOperationException(" not found in Web.config.");
                    }
                }
                catch (System.Exception)
                {
                    throw;
                }
                return connSettings.ConnectionString;
            }
        }

        public static OleDbConnection GetConnectionUser()
        {
            return new OleDbConnection(Connect_UserSQL);
        }

        //public static DataTable ExcuteDataTable_User(string query)
        //{
        //    using (SqlConnection conn = new SqlConnection(Connect_UserSQL))
        //    {
        //        using (SqlDataAdapter dap = new SqlDataAdapter(query, conn))
        //        {
        //            using (DataSet ds = new DataSet())
        //            {
        //                dap.Fill(ds);
        //                conn.Close();
        //                conn.Dispose();
        //                return ds.Tables[0];
        //            }
        //        }
        //    }
        //}
        public static DataTable ExcuteDataTable_User(string query)
        {
            using (OleDbConnection conn = new OleDbConnection(Connect_UserSQL))
            {
                using (OleDbDataAdapter dap = new OleDbDataAdapter(query, conn))
                {
                    using (DataSet ds = new DataSet())
                    {
                        dap.Fill(ds);
                        return ds.Tables[0];
                    }
                }
            }
        }
        //public static void Excute_SQL(string sql)
        //{
        //    using (OleDbConnection conn = new OleDbConnection(Connect_UserSQL))
        //    {
        //        conn.Open();
        //        using (OleDbCommand cmd = new OleDbCommand(sql, conn))
        //        {
        //            cmd.Connection = GetConnectionUser();
        //            cmd.ExecuteNonQuery();
        //        }
        //    }
        //}

        public static void Excute_SQL(string sql)
        {
            using (OleDbConnection conn = Common.GetConnectionUser())
            {
                conn.Open();
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                
                {
                   cmd.CommandText = sql;
                        cmd.Prepare();
                        cmd.ExecuteNonQuery();
                        
                }
            }
        }

    }

}
