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
    public class CVOUCHER
    {
        #region nature
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
        private string _getsql;
        public string getsql
        {
            set { _getsql = value; }
            get { return _getsql; }

        }
        private string _getsqlX;
        public string getsqlX
        {
            set { _getsqlX = value; }
            get { return _getsqlX; }

        }
        private string _COURSE_TYPE;
        public string COURSE_TYPE
        {
            set { _COURSE_TYPE = value; }
            get { return _COURSE_TYPE; }

        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private string _FINANCIAL_YEAR_INITIAL_DATE;
        public string FINANCIAL_YEAR_INITIAL_DATE
        {

            set { _FINANCIAL_YEAR_INITIAL_DATE = value; }
            get { return _FINANCIAL_YEAR_INITIAL_DATE; }

        }
        private string _ACCOUNTING_PERIOD_EXPIRATION_DATE;
        public string ACCOUNTING_PERIOD_EXPIRATION_DATE
        {

            set { _ACCOUNTING_PERIOD_EXPIRATION_DATE = value; }
            get { return _ACCOUNTING_PERIOD_EXPIRATION_DATE; }

        }
        private string _VOUCHER_DATE;
        public string VOUCHER_DATE
        {

            set { _VOUCHER_DATE = value; }
            get { return _VOUCHER_DATE; }

        }
        private string _NUM_ID;
        public string NUM_ID
        {

            set { _NUM_ID = value; }
            get { return _NUM_ID; }

        }

        private decimal _EXCHANGE_RATE;
        public decimal EXCHANGE_RATE
        {

            set { _EXCHANGE_RATE = value; }
            get { return _EXCHANGE_RATE; }

        }
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _STATUS;
        public string STATUS
        {
            set { _STATUS = value; }
            get { return _STATUS; }

        }
        private string _MANAGE_AUDIT_STATUS;
        public string MANAGE_AUDIT_STATUS
        {
            set { _MANAGE_AUDIT_STATUS = value; }
            get { return _MANAGE_AUDIT_STATUS; }

        }
        private string _FINANCIAL_AUDIT_STATUS;
        public string FINANCIAL_AUDIT_STATUS
        {
            set { _FINANCIAL_AUDIT_STATUS = value; }
            get { return _FINANCIAL_AUDIT_STATUS; }

        }
        private string _GENERAL_MANAGE_AUDIT_STATUS;
        public string GENERAL_MANAGE_AUDIT_STATUS
        {
            set { _GENERAL_MANAGE_AUDIT_STATUS = value; }
            get { return _GENERAL_MANAGE_AUDIT_STATUS; }

        }
        private string _IF_PAYFOR;
        public string IF_PAYFOR
        {

            set { _IF_PAYFOR = value; }
            get { return _IF_PAYFOR; }

        }
        #endregion
        #region sql
        string VOKEY;
        string sql = @"
SELECT 
A.VOID AS 凭证号,
B.VOUCHER_DATE AS 凭证日期,
CASE WHEN B.IF_PAYFOR='Y' THEN '已打款'
ELSE '未打款'
END  AS 是否打款,
A.SN AS 项次,
A.Abstract AS 摘要,
C.ACCODE AS 科目代码,
C.ACNAME AS 科目名称,
A.UNITPRICE AS 单价,
A.COUNT AS 数量,
D.CYCODE  AS 币别,
A.EXCHANGE_RATE AS 汇率,
A.DEBIT_ORIGINALAMOUNT AS 支出金额,
A.DEBIT_AMOUNT AS 支出本币,
A.CREDITED_ORIGINALAMOUNT AS 收入金额,
A.CREDITED_AMOUNT AS 收入本币,
A.INITIAL_DEBIT_ORIGINALAMOUNT AS 期初借方,
A.INITIAL_CREDITED_ORIGINALAMOUNT AS 期初贷方
 FROM VOUCHER_DET A
 LEFT JOIN VOUCHER_MST B ON A.VOID=B.VOID 
 LEFT JOIN ACCOUNTANT_COURSE C ON A.ACID =C.ACID 
 LEFT JOIN CURRENCY_MST D ON A.CYID =D.CYID 

";
        string sqlX = @"
SELECT 
A.VOID AS 凭证号,
B.VOUCHER_DATE AS 凭证日期,
A.SN AS 项次,
CASE WHEN B.MANAGE_AUDIT_STATUS='N'  THEN '经理未审核'
WHEN B.MANAGE_AUDIT_STATUS='Y' AND  B.FINANCIAL_AUDIT_STATUS='N' AND B.GENERAL_MANAGE_AUDIT_STATUS='N' THEN '经理已审核'
WHEN B.FINANCIAL_AUDIT_STATUS='N'  THEN '财务未审核'
WHEN B.FINANCIAL_AUDIT_STATUS='Y'AND B.GENERAL_MANAGE_AUDIT_STATUS='N' THEN '财务已审核'
WHEN B.GENERAL_MANAGE_AUDIT_STATUS='N' THEN '总经理未审核'
WHEN B.GENERAL_MANAGE_AUDIT_STATUS='Y' THEN '总经理已审核'
END AS 状态,
C.COURSE_NATURE AS 科目性质,
A.Abstract AS 摘要,
C.ACCODE AS 科目代码,
C.ACNAME AS 科目名称,
A.UNITPRICE AS 单价,
A.COUNT AS 数量,
RTRIM(CONVERT(DECIMAL(18,2),A.DEBIT_ORIGINALAMOUNT )) AS 支出金额,
RTRIM(CONVERT(DECIMAL(18,2),A.CREDITED_ORIGINALAMOUNT )) AS 收入金额,
CASE WHEN B.IF_PAYFOR='Y' THEN '已打款'
ELSE '未打款'
END  AS 是否打款,
E.ENAME AS 制单人,
B.DATE AS 制单日期
 FROM VOUCHER_DET A
 LEFT JOIN VOUCHER_MST B ON A.VOID=B.VOID 
 LEFT JOIN ACCOUNTANT_COURSE C ON A.ACID =C.ACID 
 LEFT JOIN CURRENCY_MST D ON A.CYID =D.CYID 
 LEFT JOIN EMPLOYEEINFO E ON B.MAKERID=E.EMID

";


        string sql2 = @"INSERT INTO VOUCHER_MST(

VOID,
VOUCHER_DATE,
STATUS,
MANAGE_AUDIT_STATUS,
FINANCIAL_AUDIT_STATUS,
GENERAL_MANAGE_AUDIT_STATUS,
ORIGINAL_MAKERID,
ORIGINAL_DATE,
FINANCIAL_YEAR,
PERIOD,
LAST_YEAR_CARYY_OVER_DATE,
ACCOUNTING_PERIOD_EXPIRATION_DATE,
IF_PAYFOR,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
) VALUES 

(
@VOID,
@VOUCHER_DATE,
@STATUS,
@MANAGE_AUDIT_STATUS,
@FINANCIAL_AUDIT_STATUS,
@GENERAL_MANAGE_AUDIT_STATUS,
@ORIGINAL_MAKERID,
@ORIGINAL_DATE,
@FINANCIAL_YEAR,
@PERIOD,
@LAST_YEAR_CARYY_OVER_DATE,
@ACCOUNTING_PERIOD_EXPIRATION_DATE,
@IF_PAYFOR,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY

)

";
        string sql3 = @"UPDATE VOUCHER_MST SET 
VOID=@VOID,
VOUCHER_DATE=@VOUCHER_DATE,
ORIGINAL_MAKERID=@ORIGINAL_MAKERID,
ORIGINAL_DATE=@ORIGINAL_DATE,
AUDIT_MAKERID=@AUDIT_MAKERID,
AUDIT_DATE=@AUDIT_DATE,
POSTING_MAKERID=@POSTING_MAKERID,
POSTING_DATE=@POSTING_DATE,
FINANCIAL_YEAR=@FINANCIAL_YEAR,
PERIOD=@PERIOD,
LAST_YEAR_CARYY_OVER_DATE=@LAST_YEAR_CARYY_OVER_DATE,
ACCOUNTING_PERIOD_EXPIRATION_DATE=@ACCOUNTING_PERIOD_EXPIRATION_DATE,
IF_PAYFOR=@IF_PAYFOR,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH,
DAY=@DAY

";
        string sql4 = @"INSERT INTO VOUCHER_DET(
VOKEY,
VOID,
SN,
Abstract,
ACID,
UNITPRICE,
COUNT,
CYID,
EXCHANGE_RATE,
DEBIT_ORIGINALAMOUNT,
DEBIT_AMOUNT,
CREDITED_ORIGINALAMOUNT,
CREDITED_AMOUNT,
INITIAL_DEBIT_ORIGINALAMOUNT,
INITIAL_DEBIT_AMOUNT,
INITIAL_CREDITED_ORIGINALAMOUNT,
INITIAL_CREDITED_AMOUNT,
YEAR,
MONTH,
DAY
)
VALUES (
@VOKEY,
@VOID,
@SN,
@Abstract,
@ACID,
@UNITPRICE,
@COUNT,
@CYID,
@EXCHANGE_RATE,
@DEBIT_ORIGINALAMOUNT,
@DEBIT_AMOUNT,
@CREDITED_ORIGINALAMOUNT,
@CREDITED_AMOUNT,
@INITIAL_DEBIT_ORIGINALAMOUNT,
@INITIAL_DEBIT_AMOUNT,
@INITIAL_CREDITED_ORIGINALAMOUNT,
@INITIAL_CREDITED_AMOUNT,
@YEAR,
@MONTH,
@DAY
)

";
        string sql5 = @" 
 SELECT 
 A.INITIAL_RATE,
 A.PERIOD,
 B.CYCODE,
 B.FINANCIAL_YEAR
 FROM CURRENCY_DET A 
 LEFT JOIN CURRENCY_MST B ON A.CYID=B.CYID 
";

        #endregion
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        public CVOUCHER()
        {
            IFExecution_SUCCESS = true;
            getsql = sql;
            getsqlX = sqlX;

        
        }
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYMD(12, 4, "0001", "select * from VOUCHER_GETID", "VOID", "VO");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO VOUCHER_GETID(VOID,DATE,YEAR,MONTH,DAY) VALUES ('" + v1 + "','"+varDate +"','"+year +"','"+month +"','"+day +"')");
            }
            return GETID;
        }
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            //dt.Columns.Add("索引", typeof(string));
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("科目", typeof(string));
            //dt.Columns.Add("币别", typeof(string));
            //dt.Columns.Add("汇率", typeof(decimal));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("支出金额", typeof(decimal));
            //dt.Columns.Add("支出本币", typeof(decimal));
            dt.Columns.Add("收入金额", typeof(decimal));
            //dt.Columns.Add("收入本币", typeof(decimal));
            return dt;
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo_O()
        {
            dt = this.GetTableInfo();

            //dt.Columns.Add("索引", typeof(string));
            dt.Columns.Add("凭证号", typeof(string));

            return dt;
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo_INITIAL()
        {
            dt = new DataTable();
            //dt.Columns.Add("索引", typeof(string));
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("科目代码", typeof(string));
            dt.Columns.Add("科目名称", typeof(string));
            dt.Columns.Add("年初借方", typeof(decimal));
            dt.Columns.Add("年初贷方", typeof(decimal));
            dt.Columns.Add("累计借方", typeof(decimal));
            dt.Columns.Add("累计贷方", typeof(decimal));
            dt.Columns.Add("方向", typeof(string));
            dt.Columns.Add("期初借方", typeof(decimal));
            dt.Columns.Add("期初贷方", typeof(decimal));

            return dt;
        }
        #endregion
        #region GET_TABLEINFO
        public DataTable GET_TABLEINFO_INITIAL_O()
        {

            DataTable dtX = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE");
            DataTable dt4 = this.GetTableInfo_INITIAL();
            if (dtX.Rows.Count > 0)
            {
                int n = 1;
                foreach (DataRow dr in dtX.Rows)
                {

                    DataRow dr1 = dt4.NewRow();
                    dr1["项次"] = Convert.ToInt32(n);
                    dr1["科目代码"] = dr["ACCODE"].ToString();
                    dr1["科目名称"] = dr["ACNAME"].ToString();
                    dr1["方向"] = dr["BALANCE_DIRECTION"].ToString();
                    n = n + 1;
                    dt4.Rows.Add(dr1);

                }
            }
            return dt4;
        }
        #endregion
        #region basedate
        private void basedata_INITIAL(string sql, DataTable dt, int i, string COLUMNID, string IDVALUE)
        {

            string CYCODE = bc.getOnlyString(@"
SELECT B.CYCODE FROM ACCOUNTANT_COURSE A LEFT JOIN CURRENCY_MST B ON A.CYID=B.CYID WHERE A.ACCODE='" + dt.Rows[i]["科目代码"].ToString() + "'");
            string v1 = bc.getOnlyString("SELECT FINANCIAL_YEAR FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD ='Y'");
            string v2 = bc.getOnlyString("SELECT PERIOD FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD ='Y'");
            DataTable dtx1 = bc.getdt(sql5 + " WHERE B.FINANCIAL_YEAR='" + v1 +
                "' AND A.PERIOD='" + v2 + "' AND B.CYCODE='" + CYCODE + "'");
            if (dtx1.Rows.Count > 0)
            {
                EXCHANGE_RATE = decimal.Parse(dtx1.Rows[0]["INITIAL_RATE"].ToString());
            }
            else
            {
                EXCHANGE_RATE = 1;
            }

            VOKEY = bc.numYMD(20, 12, "000000000001", "select * from VOUCHER_DET", "VOKEY", "VO");
            SQlcommandE(sql,
                VOKEY,
                IDVALUE,
                dt.Rows[i]["项次"].ToString(),
                "",
                dt.Rows[i]["科目代码"].ToString(),
                "",
                "",
                CYCODE,
                EXCHANGE_RATE.ToString(),/*9*/
                dt.Rows[i]["累计借方"].ToString(),
                (EXCHANGE_RATE * decimal.Parse(dt.Rows[i]["累计借方"].ToString())).ToString(),
                dt.Rows[i]["累计贷方"].ToString(),
                (EXCHANGE_RATE * decimal.Parse(dt.Rows[i]["累计贷方"].ToString())).ToString(),
                dt.Rows[i]["期初借方"].ToString(),
                (EXCHANGE_RATE * decimal.Parse(dt.Rows[i]["期初借方"].ToString())).ToString(),
                dt.Rows[i]["期初贷方"].ToString(),
                (EXCHANGE_RATE * decimal.Parse(dt.Rows[i]["期初贷方"].ToString())).ToString()
            );



        }
        #endregion
        #region GET_TABLEINFO
        public DataTable GET_TABLEINFO_INITIAL()/*INITIAL COURSE DATA*/
        {

            dt = this.GET_TABLEINFO_INITIAL_O();
            DataTable dt5 = bc.getdt(sql);
            if (dt5.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dt5.Rows)
                {

                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0, d5 = 0, d6 = 0;
                            if (dr["科目代码"].ToString() == dr1["科目代码"].ToString())
                            {

                                if (!string.IsNullOrEmpty(dr1["支出金额"].ToString()))
                                {
                                    d1 = decimal.Parse(dr1["支出金额"].ToString());
                                    dr["累计借方"] = d1.ToString();

                                }
                                if (!string.IsNullOrEmpty(dr1["收入金额"].ToString()))
                                {
                                    d2 = decimal.Parse(dr1["收入金额"].ToString());
                                    dr["累计贷方"] = d2.ToString();
                                }
                                if (!string.IsNullOrEmpty(dr1["期初借方"].ToString()))
                                {
                                    d3 = decimal.Parse(dr1["期初借方"].ToString());
                                    dr["期初借方"] = d3.ToString();
                                }
                                if (!string.IsNullOrEmpty(dr1["期初贷方"].ToString()))
                                {
                                    d4 = decimal.Parse(dr1["期初贷方"].ToString());
                                    dr["期初贷方"] = d4.ToString();
                                }
                                d5 = d2 - d1 + d3 - d4;
                                d6 = d1 - d2 - d3 + d4;

                                if (d5 > 0)
                                {
                                    dr["年初借方"] = d5.ToString();

                                }
                                if (d6 > 0)
                                {
                                    dr["年初贷方"] = d6.ToString();

                                }
                                break;
                            }

                        }

                    }


                }
            }
            return dt;
        }
        #endregion

        
        #region save
        public void save(string TABLENAME_MST, string TABLENAME_DET, string COLUMNID,
           string IDVALUE, DataTable dt)
        {

            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string STATUS = "OPEN";
            //string varMakerID;
            if (!bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME_DET + " WHERE " + COLUMNID + "='" + IDVALUE + "'"))
            {
                if (STATUS == "INITIAL")
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        basedata_INITIAL(sql4, dt, i, COLUMNID, IDVALUE);
                    }

                }
                else
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        basedata(sql4, dt, i, COLUMNID, IDVALUE);
                    }

                }

            }
            else
            {
                if (dt.Rows.Count > 0)
                {
                    bc.getcom("DELETE " + TABLENAME_DET + " WHERE " + COLUMNID + "='" + IDVALUE + "'");
                    if (STATUS == "INITIAL")
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            basedata_INITIAL(sql4, dt, i, COLUMNID, IDVALUE);
                        }

                    }
                    else
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            basedata(sql4, dt, i, COLUMNID, IDVALUE);
                        }

                    }
                }
                else
                {
                    bc.getcom("DELETE " + TABLENAME_MST + " WHERE " + COLUMNID + "='" + IDVALUE + "'");
                    bc.getcom("DELETE " + TABLENAME_DET + " WHERE " + COLUMNID + "='" + IDVALUE + "'");

                }
            }
            if (!bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME_DET + " WHERE " + COLUMNID + "='" + IDVALUE + "'"))
            {
                return;
            }
            else if (!bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME_MST + " WHERE " + COLUMNID + "='" + IDVALUE + "'"))
            {

                SQlcommandE(
                    sql2,
                    IDVALUE,
                    STATUS
                    );
            }
            else
            {


                SQlcommandE(sql3 + " WHERE " + COLUMNID + "='" + IDVALUE + "'", IDVALUE, STATUS);
            }
           

        }
        #endregion
        #region basedate
        private void basedata(string sql, DataTable dt, int i, string COLUMNID, string IDVALUE)
        {
            VOKEY = bc.numYMD(20, 12, "000000000001", "select * from VOUCHER_DET", "VOKEY", "VO");
            SQlcommandE(
            sql,
            VOKEY,
            IDVALUE,
            dt.Rows[i]["项次"].ToString(),
            dt.Rows[i]["摘要"].ToString(),
            bc.REMOVE_NAME(dt.Rows[i]["科目"].ToString()),
            dt.Rows[i]["单价"].ToString(),
            dt.Rows[i]["数量"].ToString(),
            "",
            "",
            dt.Rows[i]["支出金额"].ToString(),
           "",
            dt.Rows[i]["收入金额"].ToString(),
          "",
            "",
            "",
            "",
            ""
            );
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql,
            string KEY, /*1*/
            string IDVALUE, /*2*/
            string SN,  /*3*/
            string ABSTRACT, /*4*/
            string ACCODE,
            string UNITPRICE,
            string COUNT,
            string CYCODE,
            string EXCHANGE_RATE, /*9*/
            string DEBIT_ORIGINALAMOUNT,
            string DEBIT_AMOUNT,
            string CREDITED_ORIGINALAMOUNT,
            string CREDITED_AMOUNT,
            string INITIAL_DEBIT_ORIGINALAMOUNT,
            string INITIAL_DEBIT_AMOUNT,
            string INITIAL_CREDITED_ORIGINALAMOUNT,
            string INITIAL_CREDITED_AMOUNT)
        {

            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            //string varMakerID = "";
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@VOKEY", SqlDbType.VarChar, 20).Value = KEY;
            sqlcom.Parameters.Add("@VOID", SqlDbType.VarChar, 20).Value = IDVALUE;
            sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = SN;
            sqlcom.Parameters.Add("@ABSTRACT", SqlDbType.VarChar, 100).Value = ABSTRACT;
            sqlcom.Parameters.Add("@ACID", SqlDbType.VarChar, 20).Value = bc.getOnlyString(@"SELECT ACID FROM ACCOUNTANT_COURSE WHERE ACCODE='" + ACCODE
                + "' ");
            sqlcom.Parameters.Add("@UNITPRICE", SqlDbType.VarChar, 20).Value = UNITPRICE;
            sqlcom.Parameters.Add("@COUNT", SqlDbType.VarChar, 20).Value = COUNT;
            sqlcom.Parameters.Add("@CYID", SqlDbType.VarChar, 20).Value = "";
            if (!string.IsNullOrEmpty(EXCHANGE_RATE))
            {
                sqlcom.Parameters.Add("@EXCHANGE_RATE", SqlDbType.VarChar, 20).Value = EXCHANGE_RATE;
            }
            else
            {
                sqlcom.Parameters.Add("@EXCHANGE_RATE", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }


            if (!string.IsNullOrEmpty(DEBIT_ORIGINALAMOUNT))
            {
                sqlcom.Parameters.Add("@DEBIT_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = DEBIT_ORIGINALAMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@DEBIT_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            if (!string.IsNullOrEmpty(DEBIT_AMOUNT))
            {
                sqlcom.Parameters.Add("@DEBIT_AMOUNT", SqlDbType.VarChar, 20).Value = DEBIT_AMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@DEBIT_AMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            if (!string.IsNullOrEmpty(CREDITED_ORIGINALAMOUNT))
            {
                sqlcom.Parameters.Add("@CREDITED_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = CREDITED_ORIGINALAMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@CREDITED_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;

            }
            if (!string.IsNullOrEmpty(CREDITED_AMOUNT))
            {
                sqlcom.Parameters.Add("@CREDITED_AMOUNT", SqlDbType.VarChar, 20).Value = CREDITED_AMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@CREDITED_AMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }


            if (!string.IsNullOrEmpty(INITIAL_DEBIT_ORIGINALAMOUNT))
            {
                sqlcom.Parameters.Add("@INITIAL_DEBIT_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = INITIAL_DEBIT_ORIGINALAMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@INITIAL_DEBIT_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;

            }
            if (!string.IsNullOrEmpty(INITIAL_DEBIT_AMOUNT))
            {
                sqlcom.Parameters.Add("@INITIAL_DEBIT_AMOUNT", SqlDbType.VarChar, 20).Value = INITIAL_DEBIT_AMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@INITIAL_DEBIT_AMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;

            }
            if (!string.IsNullOrEmpty(INITIAL_CREDITED_ORIGINALAMOUNT))
            {
                sqlcom.Parameters.Add("@INITIAL_CREDITED_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = INITIAL_CREDITED_ORIGINALAMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@INITIAL_CREDITED_ORIGINALAMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
   
            if (!string.IsNullOrEmpty(INITIAL_CREDITED_AMOUNT))
            {
                sqlcom.Parameters.Add("@INITIAL_CREDITED_AMOUNT", SqlDbType.VarChar, 20).Value = INITIAL_CREDITED_AMOUNT;
            }
            else
            {
                sqlcom.Parameters.Add("@INITIAL_CREDITED_AMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql, string v1, string v3)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            string varMakerID = EMID;
            PERIOD period = new PERIOD();
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@VOID", SqlDbType.VarChar, 20).Value = v1;
            sqlcom.Parameters.Add("@VOUCHER_DATE", SqlDbType.VarChar, 20).Value = VOUCHER_DATE;
            sqlcom.Parameters.Add("@MANAGE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = MANAGE_AUDIT_STATUS;
            sqlcom.Parameters.Add("@FINANCIAL_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = FINANCIAL_AUDIT_STATUS;
            sqlcom.Parameters.Add("@GENERAL_MANAGE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = GENERAL_MANAGE_AUDIT_STATUS;
            sqlcom.Parameters.Add("@STATUS", SqlDbType.VarChar, 20).Value = "OPEN";
            sqlcom.Parameters.Add("@ORIGINAL_MAKERID", SqlDbType.VarChar, 20).Value = "";
            sqlcom.Parameters.Add("@ORIGINAL_DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@AUDIT_DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@POSTING_MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@POSTING_DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@FINANCIAL_YEAR", SqlDbType.VarChar, 20).Value = "";
            sqlcom.Parameters.Add("@PERIOD", SqlDbType.VarChar, 20).Value = "";
            sqlcom.Parameters.Add("@IF_PAYFOR", SqlDbType.VarChar, 20).Value = IF_PAYFOR;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            if (FINANCIAL_YEAR_INITIAL_DATE == null)
            {
                sqlcom.Parameters.Add("@LAST_YEAR_CARYY_OVER_DATE", SqlDbType.VarChar, 20).Value = "";
            }
            else
            {
                sqlcom.Parameters.Add("@LAST_YEAR_CARYY_OVER_DATE", SqlDbType.VarChar, 20).Value = FINANCIAL_YEAR_INITIAL_DATE;
            }

            sqlcom.Parameters.Add("@ACCOUNTING_PERIOD_EXPIRATION_DATE", SqlDbType.VarChar, 20).Value = ACCOUNTING_PERIOD_EXPIRATION_DATE;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region Search()
        public DataTable Search(string ACCODE, string ACNAME)
        {

            string sql1 = @" where A.ACCODE like '%" + ACCODE + "%' AND A.ACNAME LIKE '%" + ACNAME + "%' ORDER BY ACCODE ASC";
            dt = basec.getdts(sql + sql1);
            return dt;
        }
        #endregion

   
        public bool CheckIfALLOW_SAVEOR_DELETE(string VOID,string USID)
        {
            bool b = false;
            EMID= bc.getOnlyString ("SELECT EMID FROM USERINFO WHERE USID='"+USID +"'");
            string MAKERID = bc.getOnlyString("SELECT MAKERID FROM VOUCHER_MST WHERE VOID='"+VOID +"'");
            if (bc.exists("SELECT STATUS FROM VOUCHER_MST WHERE VOID='" + VOID + "'"))
            {
                string s2 = bc.getOnlyString("SELECT STATUS FROM VOUCHER_MST WHERE VOID='" + VOID + "'");
                string v1 = bc.getOnlyString("SELECT GENERAL_MANAGE FROM RIGHTLIST WHERE USID='"+USID +"' AND NODE_NAME='录入凭证作业'");
                if ((this.RETURN_FINANCIAL_AUDIT_STATUS (VOID ) == "Y" || this .RETURN_GENERAL_AUDIT_STATUS (VOID ) == "Y") && v1!="Y")
                {

                    b = true;
                    ErrowInfo = "财务已审核或总经理已审核凭证只允许总经理删除与修改";

                }
                else if (EMID !=MAKERID && v1!="Y")
                {
                    b = true;
                    ErrowInfo = "非凭证的制单人只允许总经理删除与修改";

                }

             
            }

            return b;
        }
        #region GetTableInfoX
        public DataTable GetTableInfo_X()
        {
            dt = new DataTable();
            dt.Columns.Add("凭证号", typeof(string));
            dt.Columns.Add("凭证日期", typeof(string));
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("状态", typeof(string));
            dt.Columns.Add("科目性质", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("科目代码", typeof(string));
            dt.Columns.Add("科目名称", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("支出金额", typeof(string));
            dt.Columns.Add("收入金额", typeof(string ));
            dt.Columns.Add("公司余额", typeof(string ));
            dt.Columns.Add("是否打款", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
            dt.Columns.Add("制单日期", typeof(string));
            return dt;
        }
        #endregion
        #region  GET_CALCULATE
        public DataTable GET_CALCULATE(DataTable dt) /*流水账余额TABLE*/
        {
            DataTable dtt = this.GetTableInfo_X();
            if (dt.Rows.Count > 0)
            {
                decimal SUM = 0;
                string v1 = dt.Rows[0]["科目代码"].ToString();
                foreach (DataRow dr1 in dt.Rows)
                {
                    decimal d1 = 0, d2 = 0;
                    if (v1 != dr1["科目代码"].ToString())
                    {
                        SUM = 0;
                        v1 = dr1["科目代码"].ToString();

                    }
                    if (!string.IsNullOrEmpty(dr1["支出金额"].ToString()))
                    {
                        d1 = decimal.Parse(dr1["支出金额"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr1["收入金额"].ToString()))
                    {
                        d2 = decimal.Parse(dr1["收入金额"].ToString());
                    }
                    SUM = SUM + d2 - d1;
                    DataRow dr = dtt.NewRow();
                    dr["凭证号"] = dr1["凭证号"].ToString();
                    dr["凭证日期"] = dr1["凭证日期"].ToString();
                    dr["项次"] = dr1["项次"].ToString();
                    dr["状态"] = dr1["状态"].ToString();
                    dr["科目性质"] = dr1["科目性质"].ToString();
                    dr["摘要"] = dr1["摘要"].ToString();
                    dr["科目代码"] = dr1["科目代码"].ToString();
                    dr["科目名称"] = dr1["科目名称"].ToString();
                    dr["单价"] = dr1["单价"].ToString();
                    dr["数量"] = dr1["数量"].ToString();
                    dr["支出金额"] = dr1["支出金额"].ToString();
                    dr["收入金额"] = dr1["收入金额"].ToString();
                    dr["公司余额"] = SUM;
                    dr["是否打款"] = dr1["是否打款"].ToString();
                    dr["制单人"] = dr1["制单人"].ToString();
                    dr["制单日期"] = dr1["制单日期"].ToString();
                    dtt.Rows.Add(dr);
                }

            }
            return dtt;
        }
        #endregion

        public bool JUAGE_IFCOURSE_NATURE_NOTSAME(DataTable dt)
        {
            bool b = false;
           for(int i=0;i<dt.Rows .Count-1 ;i++)
            { 
                string v1 = bc.getOnlyString("SELECT COURSE_NATURE FROM ACCOUNTANT_COURSE WHERE ACCODE='" +bc.REMOVE_NAME (dt.Rows [i]["科目"].ToString()) + "'");
                string v2 = bc.getOnlyString("SELECT COURSE_NATURE FROM ACCOUNTANT_COURSE WHERE ACCODE='" +bc.REMOVE_NAME (dt.Rows [i+1]["科目"].ToString()) + "'");
                if (v1 != v2)
                {
                    b = true;
                    ErrowInfo = "同一个凭证科目性质需要一至";
                    break;
                }

            }
            return b;
        }
        public string RETURN_GENERAL_AUDIT_STATUS(string VOID)
        {
         
            STATUS  = bc.getOnlyString("SELECT GENERAL_MANAGE_AUDIT_STATUS FROM VOUCHER_MST WHERE VOID='" + VOID + "'");
            return STATUS;
        }
        public string RETURN_FINANCIAL_AUDIT_STATUS(string VOID)
        {

            STATUS = bc.getOnlyString("SELECT FINANCIAL_AUDIT_STATUS FROM VOUCHER_MST WHERE VOID='" + VOID + "'");
            return STATUS;
        }
        public string RETURN_MANAGE_AUDIT_STATUS(string VOID)
        {

            STATUS = bc.getOnlyString("SELECT MANAGE_AUDIT_STATUS FROM VOUCHER_MST WHERE VOID='" + VOID + "'");
            return STATUS;
        }
    }
}
