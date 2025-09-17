using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using Microsoft.SqlServer.Server;
using System.Collections;

namespace PrintProgram
{
    class db_engsr
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        public static SqlConnection GetCon()
        {
            string cnstr = "server=192.168.5.25;database=i_Factory_SOP_DB;uid=SuperSOP;pwd=A12345678;Trusted_Connection=False;Connect Timeout = 10";
            //i_Program_Server .5.x網段↓↓；共用DB.4.37對i_Server→.5.28
            //string cnstr = "server=192.168.5.28;database=i_Factory_SOP_DB;uid=SuperSOP;pwd=A12345678;Connect Timeout = 10";

            SqlConnection icn = new SqlConnection();
            icn.ConnectionString = cnstr;
            if (icn.State == ConnectionState.Open) icn.Close();
            icn.Open();

            return icn;
        }

        public static bool Exsql(string cmdtxt)
        {
            SqlConnection con = db_engsr.GetCon();//連接資料庫
            //con.Open();
            SqlCommand cmd = new SqlCommand(cmdtxt, con);
            try
            {
                cmd.ExecuteNonQuery();//執行SQL 語句並返回受影響的行數
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
            finally
            {
                con.Dispose();//釋放連接物件資源
                con.Close();
            }
        }
        public static DataSet reDs(string cmdtxt)
        {
            SqlConnection con = db_engsr.GetCon();
            SqlDataAdapter da = new SqlDataAdapter(cmdtxt, con);
            //建立資料集ds
            DataSet ds = new DataSet();
            da.Fill(ds);

            return ds;
        }

        public static string scalDs(string str_select)
        {
            //執行ExecuteScalar()，傳回單一字串,若遇NULL值,直接當空字串作
            //--------------------------------------------------------------------
            SqlConnection con = db_engsr.GetCon();
            SqlCommand com_select = new SqlCommand(str_select, con);
            try
            {
                con.Open();
                str_select = Convert.ToString(com_select.ExecuteScalar());
            }
            catch (Exception ex)
            {
                con.Close();
                return Convert.ToString(ex);
            }
            finally
            {
                con.Close();
            }
            return str_select;
        }
    }
}
