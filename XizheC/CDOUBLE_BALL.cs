using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using XizheC;

namespace XizheC
{
    public class CDOUBLE_BALL
    {

        private string _sql;
        public string sql
        {
            set { _sql = value; }
            get { return _sql; ; }

        }
        private string _sqlo;
        public string sqlo
        {
            set { _sqlo = value; }
            get { return _sqlo; ; }

        }
        private string _sqlt;
        public string sqlt
        {
            set { _sqlt = value; }
            get { return _sqlt; ; }

        }
        private string _sqlth;
        public string sqlth
        {
            set { _sqlth = value; }
            get { return _sqlth; ; }

        }
        private string _sqlf;
        public string sqlf
        {
            set { _sqlf = value; }
            get { return _sqlf; ; }

        }
        private string _sqlfi;
        public string sqlfi
        {
            set { _sqlfi = value; }
            get { return _sqlfi; ; }

        }

        private string _sqlsi;
        public string sqlsi
        {
            set { _sqlsi = value; }
            get { return _sqlsi; ; }

        }
        private string _sqlse;
        public string sqlse
        {
            set { _sqlse = value; }
            get { return _sqlse; ; }

        }
        private string _PWD;
        public string PWD
        {
            set { _PWD = value; }
            get { return _PWD; }

        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        
        string setsql = @"

";
        string setsqlo = @"

SELECT
RED_BALL_ONE,
COUNT(*) AS COUNT 
FROM 
DOUBLE_BALL
GROUP BY RED_BALL_ONE 


";
        string setsqlt = @"

SELECT
RED_BALL_TWO,
COUNT(*) AS COUNT
FROM 
DOUBLE_BALL 
GROUP BY RED_BALL_TWO 


";
        string setsqlth = @"
SELECT
RED_BALL_THREE,
COUNT(*) AS COUNT
FROM DOUBLE_BALL 
GROUP BY 
RED_BALL_THREE
";
        string setsqlf = @"
SELECT
RED_BALL_FOUR ,
COUNT(*) AS COUNT 
FROM
DOUBLE_BALL  
GROUP BY RED_BALL_FOUR 
";
        string setsqlfi = @"
SELECT
RED_BALL_FIVE ,
COUNT(*) AS COUNT
FROM DOUBLE_BALL 
GROUP BY RED_BALL_FIVE
";
        string setsqlsi = @"
SELECT
RED_BALL_SIX ,
COUNT(*) AS COUNT
FROM DOUBLE_BALL 
GROUP BY RED_BALL_SIX 

";
        string setsqlse = @"
SELECT
BLUE_BALL ,
COUNT(*) AS COUNT 
FROM DOUBLE_BALL 
GROUP BY BLUE_BALL 

";
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        public CDOUBLE_BALL()
        {
            IFExecution_SUCCESS = true;
            sql = setsql;
            sqlo = setsqlo;
            sqlt = setsqlt;
            sqlth = setsqlth;
            sqlf = setsqlf;
            sqlfi = setsqlfi;
            sqlsi = setsqlsi;
            sqlse = setsqlse;
        }
        public string GETID()
        {
            string v1 = bc.numYM(7, 3, "001", "SELECT * FROM EMPLOYEEINFO", "EMID", "");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #region GET_COLUMNS_INFO()
        public DataTable GET_COLUMNS_INFO()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("红球一", typeof(string));
            dt.Columns.Add("红球一出现次数", typeof(decimal));
            dt.Columns.Add("红球二", typeof(string));
            dt.Columns.Add("红球二出现次数", typeof(decimal));
            dt.Columns.Add("红球三", typeof(string));
            dt.Columns.Add("红球三出现次数", typeof(decimal));
            dt.Columns.Add("红球四", typeof(string));
            dt.Columns.Add("红球四出现次数", typeof(decimal));
            dt.Columns.Add("红球五", typeof(string));
            dt.Columns.Add("红球五出现次数", typeof(decimal));
            dt.Columns.Add("红球六", typeof(string));
            dt.Columns.Add("红球六出现次数", typeof(decimal));
            dt.Columns.Add("蓝球", typeof(string));
            dt.Columns.Add("蓝球出现次数", typeof(decimal));
            return dt;
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = GET_COLUMNS_INFO();
            string j = "";
            for (int i = 1; i <= 33; i++)
            {
                DataRow dr = dt.NewRow();
                if (i.ToString().Length < 2)
                {
                    j = "0" + i.ToString();
                }
                else
                {
                    j = i.ToString();
                }
                dr["红球一"] = j;
                dr["红球二"] = j;
                dr["红球三"] = j;
                dr["红球四"] = j;
                dr["红球五"] = j;
                dr["红球六"] = j;

                dr["红球一出现次数"] = "0";
                dr["红球二出现次数"] = "0";
                dr["红球三出现次数"] = "0";
                dr["红球四出现次数"] = "0";
                dr["红球五出现次数"] = "0";
                dr["红球六出现次数"] = "0";
                if (i < 17)
                {
                    dr["蓝球"] = j;
                    dr["蓝球出现次数"] = "0";
                }
                dt.Rows.Add(dr);
            }
       
            DataTable dtt = new DataTable();
            dtt = bc.getdt(sqlo);
            foreach (DataRow dr in dtt.Rows)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    if (dr["RED_BALL_ONE"].ToString() == dr1["红球一"].ToString())
                    {

                        dr1["红球一出现次数"] = dr["COUNT"].ToString();
                        break;
                    }
                }
            }
            dtt = bc.getdt(sqlt);
            foreach (DataRow dr in dtt.Rows)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    if (dr["RED_BALL_TWO"].ToString() == dr1["红球二"].ToString())
                    {

                        dr1["红球二出现次数"] = dr["COUNT"].ToString();
                        break;
                    }
                }
            }
            dtt = bc.getdt(sqlth);
            foreach (DataRow dr in dtt.Rows)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    if (dr["RED_BALL_THREE"].ToString() == dr1["红球三"].ToString())
                    {

                        dr1["红球三出现次数"] = dr["COUNT"].ToString();
                        break;
                    }
                }
            }
            dtt = bc.getdt(sqlf);
            foreach (DataRow dr in dtt.Rows)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    if (dr["RED_BALL_FOUR"].ToString() == dr1["红球四"].ToString())
                    {

                        dr1["红球四出现次数"] = dr["COUNT"].ToString();
                        break;
                    }
                }
            }
            dtt = bc.getdt(sqlfi);
            foreach (DataRow dr in dtt.Rows)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    if (dr["RED_BALL_FIVE"].ToString() == dr1["红球五"].ToString())
                    {

                        dr1["红球五出现次数"] = dr["COUNT"].ToString();
                        break;
                    }
                }
            }
            dtt = bc.getdt(sqlsi);
            foreach (DataRow dr in dtt.Rows)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    if (dr["RED_BALL_SIX"].ToString() == dr1["红球六"].ToString())
                    {

                        dr1["红球六出现次数"] = dr["COUNT"].ToString();
                        break;
                    }
                }
            }
            dtt = bc.getdt(sqlse);
            foreach (DataRow dr in dtt.Rows)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    if (dr["BLUE_BALL"].ToString() == dr1["蓝球"].ToString())
                    {

                        dr1["蓝球出现次数"] = dr["COUNT"].ToString();
                        break;
                    }
                }
            }
            return dt;
        }
        #endregion

        #region GetTableInfo_t
        public DataTable GetTableInfo_t()
        {
            DataTable  dt = GET_COLUMNS_INFO();
            DataTable dtt = GetTableInfo();
            for (int i = 0; i < 2; i++)
            {
                DataRow dr = dt.NewRow();
                dr["序号"] = i + 1;
                dt.Rows.Add(dr);

            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "红球一出现次数 DESC", "");
            for (int j = 0; j < 2; j++)
            {

                dt.Rows[j]["红球一"] = dtt.Rows[j]["红球一"].ToString();
                dt.Rows[j]["红球一出现次数"] = dtt.Rows[j]["红球一出现次数"].ToString();

            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "红球二出现次数 DESC", "");
            for (int j = 0; j < 2; j++)
            {

              dt.Rows[j]["红球二"] = dtt.Rows[j]["红球二"].ToString();
              dt.Rows[j]["红球二出现次数"] = dtt.Rows[j]["红球二出现次数"].ToString();

            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "红球三出现次数 DESC", "");
            for (int j = 0; j < 2; j++)
            {

                dt.Rows[j]["红球三"] = dtt.Rows[j]["红球三"].ToString();
                dt.Rows[j]["红球三出现次数"] = dtt.Rows[j]["红球三出现次数"].ToString();

            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "红球四出现次数 DESC", "");
            for (int j = 0; j < 2; j++)
            {

                dt.Rows[j]["红球四"] = dtt.Rows[j]["红球四"].ToString();
                dt.Rows[j]["红球四出现次数"] = dtt.Rows[j]["红球四出现次数"].ToString();

            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "红球五出现次数 DESC", "");
            for (int j = 0; j < 2; j++)
            {

                dt.Rows[j]["红球五"] = dtt.Rows[j]["红球五"].ToString();
                dt.Rows[j]["红球五出现次数"] = dtt.Rows[j]["红球五出现次数"].ToString();

            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "红球六出现次数 DESC", "");
            for (int j = 0; j < 2; j++)
            {

                dt.Rows[j]["红球六"] = dtt.Rows[j]["红球六"].ToString();
                dt.Rows[j]["红球六出现次数"] = dtt.Rows[j]["红球六出现次数"].ToString();

            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "蓝球出现次数 ASC", "");
            dt.Rows[0]["蓝球"] = dtt.Rows[32]["蓝球"].ToString();
            dt.Rows[0]["蓝球出现次数"] = dtt.Rows[32]["蓝球出现次数"].ToString();
            dt.Rows[1]["蓝球"] = dtt.Rows[31]["蓝球"].ToString();
            dt.Rows[1]["蓝球出现次数"] = dtt.Rows[31]["蓝球出现次数"].ToString();
            return dt;
        }
        #endregion
      
    }
}
