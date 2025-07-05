using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Drawing;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Data.SqlClient;
using System.Text;
using System.IO;
using System.Security.Cryptography;
using System.Collections.Generic;
using System.Net;
using System.Data.OleDb;
using System.Xml;

namespace Webserver
{
    public class basec
    {
        #region  建立数据库连接
        /// <summary>
        /// 建立数据库连接.
        /// </summary>
        /// <returns>返回SqlConnection对象</returns>
        /// 
        // 摘要:
        public SqlConnection getcon()
        {
            
            string M_str_sqlcon = ConfigurationManager.AppSettings["ConnectionDB"].ToString();
            SqlConnection myCon = new SqlConnection(M_str_sqlcon);
          
            return myCon;
        }
        #endregion
        public  string abc()
        {
            return   ConfigurationManager.AppSettings["ConnectionDB"].ToString();
        }
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
        private static  bool _IFExecutionSUCCESS;
        public static bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        StringBuilder sqb = new StringBuilder();
        #region  执行SqlCommand命令
        /// <summary>
        /// 执行SqlCommand
        /// </summary>
        /// <param name="M_str_sqlstr">SQL语句</param>
        /// 
        public void getcom(string M_str_sqlstr)
        {
            SqlConnection sqlcon = this.getcon();
            sqlcon.Open();
            SqlCommand sqlcom = new SqlCommand(M_str_sqlstr, sqlcon);
            sqlcom.ExecuteNonQuery();
            sqlcom.Dispose();
            sqlcon.Close();
            sqlcon.Dispose();
        }
        #endregion
        #region getcoms
        public static void getcoms(string M_str_sqlstr)
        {
            basec bc = new basec();
            SqlConnection sqlcon = bc.getcon();
            sqlcon.Open();
            SqlCommand sqlcom = new SqlCommand(M_str_sqlstr, sqlcon);
            sqlcom.ExecuteNonQuery();
            sqlcom.Dispose();
            sqlcon.Close();
            sqlcon.Dispose();
        }
        #endregion

        #region  创建DataSet对象
        /// <summary>
        /// 创建一个DataSet对象
        /// </summary>
        /// <param name="M_str_sqlstr">SQL语句</param>
        /// <param name="M_str_table">表名</param>
        /// <returns>返回DataSet对象</returns>
        public DataSet getds(string M_str_sqlstr, string M_str_table)
        {
            SqlConnection sqlcon = this.getcon();
            SqlDataAdapter sqlda = new SqlDataAdapter(M_str_sqlstr, sqlcon);
            DataSet myds = new DataSet();
            sqlda.Fill(myds, M_str_table);
            return myds;
        }
        #endregion

        #region  创建SqlDataReader对象
        /// <summary>
        /// 创建一个SqlDataReader对象
        /// </summary>
        /// <param name="M_str_sqlstr">SQL语句</param>
        /// <returns>返回SqlDataReader对象</returns>
        public SqlDataReader getread(string M_str_sqlstr)
        {
            SqlConnection sqlcon = this.getcon();
            SqlCommand sqlcom = new SqlCommand(M_str_sqlstr, sqlcon);
            sqlcon.Open();
            SqlDataReader sqlread = sqlcom.ExecuteReader(CommandBehavior.CloseConnection);
            return sqlread;
        }
        #endregion

        public DataTable table(string M_str_sql)
        {
            SqlConnection sqlcon = this.getcon();
            SqlCommand sqlcmd = new SqlCommand(M_str_sql, sqlcon);
            sqlcon.Open();
            SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            sqlcon.Close();
            GC.Collect();
            return dt;
        }
  
        public DataTable getdt(string M_str_sql)
        {
            SqlConnection sqlcon = this.getcon();
            SqlCommand sqlcom = new SqlCommand(M_str_sql, sqlcon);
            SqlDataAdapter da = new SqlDataAdapter(sqlcom);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;

        }
        public static DataTable getdts(string M_str_sql)
        {
            basec bc = new basec();
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(M_str_sql, sqlcon);
            SqlDataAdapter da = new SqlDataAdapter(sqlcom);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;

        }
        public SqlDataAdapter getda(string M_str_sql)
        {
            SqlConnection sqlcon = this.getcon();
            SqlCommand sqlcom = new SqlCommand(M_str_sql, sqlcon);
            SqlDataAdapter da = new SqlDataAdapter(sqlcom);
            return da;


        }
        #region getOnlyString
        public string getOnlyString(string sql)
        {
            string s2 = "";
            DataTable dtu2 = this.getdt(sql);
            if (dtu2.Rows.Count > 0)
            {
                s2 = dtu2.Rows[0][0].ToString();

            }

            return s2;
        }
        #endregion
        
 
       
    }

}
