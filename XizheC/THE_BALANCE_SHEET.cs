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
    public class THE_BALANCE_SHEET
    {
        
        private string _getsql;
        public string getsql
        {
            set { _getsql = value; }
            get { return _getsql; ; }

        }
        private string _getsqlo;
        public string getsqlo
        {
            set { _getsqlo = value; }
            get { return _getsqlo; ; }

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
        private List<string> _PROPERTY;
        public List<string> PROPERTY
        {
            set { _PROPERTY = value; }
            get { return _PROPERTY; ; }

        }
        private List<string> _LIABILITIES_AND_CAPITIAL;
        public List<string> LIABILITIES_AND_CAPITIAL
        {
            set { _LIABILITIES_AND_CAPITIAL = value; }
            get { return _LIABILITIES_AND_CAPITIAL; ; }

        }
        string sql = @"
SELECT 
SUBSTRING(C.ACCODE,1,4) AS 分类科目,
(SELECT ACNAME FROM ACCOUNTANT_COURSE WHERE ACCODE=SUBSTRING(C.ACCODE,1,4)) AS 分类科目名称,
(SELECT ACNAME FROM ACCOUNTANT_COURSE WHERE ACCODE=C.ACCODE) AS 科目名称,
C.ACCODE AS 科目代码,
B.FINANCIAL_YEAR AS 会计年度,
B.PERIOD AS 帐期,
A.CYID AS 币别编号,
D.CYCODE  AS 币别,
D.CYNAME AS 名称,
A.EXCHANGE_RATE AS 汇率,
B.STATUS AS 状态,
C.BALANCE_DIRECTION AS 方向,
A.INITIAL_DEBIT_ORIGINALAMOUNT AS 期初借方原币,
A.INITIAL_DEBIT_AMOUNT AS 期初借方本币,
A.INITIAL_CREDITED_ORIGINALAMOUNT  AS 期初贷方原币,
A.INITIAL_CREDITED_AMOUNT AS 期初贷方本币,
A.DEBIT_ORIGINALAMOUNT AS 借方原币金额,
A.DEBIT_AMOUNT AS 借方本币金额,
A.CREDITED_ORIGINALAMOUNT AS 贷方原币金额,
A.CREDITED_AMOUNT AS 贷方本币金额
FROM VOUCHER_DET A 
LEFT JOIN VOUCHER_MST B ON A.VOID=B.VOID 
LEFT JOIN ACCOUNTANT_COURSE C ON A.ACID =C.ACID 
LEFT JOIN CURRENCY_MST D ON A.CYID =D.CYID
";



        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        List<string> list = new List<string>();
        List<string> listo = new List<string>();
        string[] xw=new string[]{
"流动资产",
"货币资金",
"交易性金融资产"
        };
        
        string[] xwo = new string[]{
            "",
            ""};
        public THE_BALANCE_SHEET()
        {
            IFExecution_SUCCESS = true;
            getsql = sql;
            
            for (int i = 0; i < xw.Length; i++)
            {
               list .Add (xw[i]);
            }
            for (int i = 0; i < xwo.Length; i++)
            {
                listo .Add(xwo[i]);
            }
            PROPERTY = list;
            LIABILITIES_AND_CAPITIAL = listo;

        }
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("科目代码", typeof(string));
            dt.Columns.Add("科目名称", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("凭证号", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("借方", typeof(decimal));
            dt.Columns.Add("贷方", typeof(decimal));
            dt.Columns.Add("方向", typeof(string));
            dt.Columns.Add("余额", typeof(decimal));
            return dt;
        }
        #endregion
        #region GetTableInfo_INITIAL
        public DataTable GetTableInfo_INITAIL()
        {
            dt = new DataTable();
            dt.Columns.Add("科目代码", typeof(string));
            dt.Columns.Add("科目名称", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("凭证号", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("借方", typeof(decimal));
            dt.Columns.Add("贷方", typeof(decimal));
            dt.Columns.Add("方向", typeof(string));
            dt.Columns.Add("余额", typeof(decimal));
            DataTable dtx = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE WHERE PARENT_NODEID IS NULL ORDER BY ACCODE ASC");
            foreach (DataRow dr in dtx.Rows)
            {

                DataRow dr1 = dt.NewRow();
                dr1["科目代码"] = dr["ACCODE"].ToString();
                dr1["科目名称"] = dr["ACNAME"].ToString();
                dr1["摘要"] = "上年结转";
                dt.Rows.Add(dr1);
                DataRow dr2 = dt.NewRow();
                dr2["摘要"] = "本期合计";
                dt.Rows.Add(dr2);
                DataRow dr3 = dt.NewRow();
                dr3["摘要"] = "本年累计";
                dt.Rows.Add(dr3);
            }
            return dt;
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo_O()
        {
            dt = this.GetTableInfo();

            dt.Columns.Add("结转", typeof(decimal));
            dt.Columns.Add("本期合计", typeof(decimal));
            dt.Columns.Add("本年累计", typeof(decimal));
            return dt;
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

        #region GET_TABLEINFO
        public DataTable GET_TABLEINFO(DataTable dt)
        {
            decimal sum = 0, sum1 = 0, sum2 = 0;
            string accode;

            DataTable dt4 = this.GetTableInfo();
            if (dt.Rows.Count > 0)
            {
                accode = dt.Rows[0]["科目代码"].ToString();

                for (int i = 0; i < dt.Rows.Count; i++)
                {

                    decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0, d5 = 0, d6 = 0, d7 = 0, d8 = 0;
                    if (!string.IsNullOrEmpty(dt.Rows[i]["借方原币金额"].ToString()))
                    {
                        d1 = decimal.Parse(dt.Rows[i]["借方原币金额"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["贷方原币金额"].ToString()))
                    {
                        d2 = decimal.Parse(dt.Rows[i]["贷方原币金额"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["期初借方"].ToString()))
                    {
                        d3 = decimal.Parse(dt.Rows[i]["期初借方"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["期初贷方"].ToString()))
                    {
                        d4 = decimal.Parse(dt.Rows[i]["期初贷方"].ToString());
                    }

                    d5 = d2 - d1 + d3 - d4;
                    d6 = d1 - d2 - d3 + d4;
                    if (accode != dt.Rows[i]["科目代码"].ToString())
                    {
                        sum = 0;
                        sum1 = 0;
                        sum2 = 0;
                        accode = dt.Rows[i]["科目代码"].ToString();
                    }
                    DataRow dr1 = dt4.NewRow();
                    if (dt.Rows[i]["期初借方"].ToString() != "" || dt.Rows[i]["期初贷方"].ToString() != "")
                    {
                        dr1["摘要"] = "上年结转";
                        dr1["科目代码"] = dt.Rows[i]["科目代码"].ToString();
                        dr1["科目名称"] = dt.Rows[i]["科目名称"].ToString();
                        if (dt.Rows[i]["状态"].ToString() != "INITIAL")
                        {
                            MessageBox.Show(dt.Rows[i]["状态"].ToString());

                            dr1["凭证号"] = dt.Rows[i]["凭证号"].ToString();
                        }


                        dr1["日期"] = dt.Rows[i]["结转日期"].ToString();


                        if (d5 > 0)
                        {
                            dr1["余额"] = d5;
                            dr1["方向"] = dt.Rows[i]["方向"].ToString();
                            sum2 = d5;
                        }

                        if (d6 > 0)
                        {
                            dr1["余额"] = d6;
                            dr1["方向"] = "贷";
                            sum2 = -d6;
                        }
                        dt4.Rows.Add(dr1);
                    }

                    DataRow dr2 = dt4.NewRow();
                    dr2["摘要"] = "本期合计";
                    if (dt.Rows[i]["期初借方"].ToString() != "" || dt.Rows[i]["期初贷方"].ToString() != "")
                    {
                    }
                    else
                    {
                        dr2["科目代码"] = dt.Rows[i]["科目代码"].ToString();
                        dr2["科目名称"] = dt.Rows[i]["科目名称"].ToString();
                        if (dt.Rows[i]["状态"].ToString() != "INITIAL")
                        {

                            dr2["凭证号"] = dt.Rows[i]["凭证号"].ToString();
                        }


                    }
                    dr2["日期"] = dt.Rows[i]["本期结帐日期"].ToString();
                    if (!string.IsNullOrEmpty(dt.Rows[i]["借方原币金额"].ToString()))
                    {
                        dr2["借方"] = dt.Rows[i]["借方原币金额"].ToString();
                        sum = sum + decimal.Parse(dt.Rows[i]["借方原币金额"].ToString());


                    }
                    else
                    {
                        dr2["借方"] = DBNull.Value;
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["贷方原币金额"].ToString()))
                    {
                        dr2["贷方"] = dt.Rows[i]["贷方原币金额"].ToString();
                        sum1 = sum1 + decimal.Parse(dt.Rows[i]["贷方原币金额"].ToString());
                    }
                    else
                    {
                        dr2["贷方"] = DBNull.Value;
                    }

                    d7 = d1 - d2 + sum2;
                    d8 = d2 - d1 - sum2;

                    if (d7 > 0)
                    {
                        dr2["余额"] = d7;
                        dr2["方向"] = dt.Rows[i]["方向"].ToString();
                        sum2 = d7;
                    }

                    if (d8 > 0)
                    {
                        dr2["余额"] = d8;
                        dr2["方向"] = "贷";
                        sum2 = -d8;
                    }
                    if (d7 == 0)
                    {

                        dr2["余额"] = 0;
                        dr2["方向"] = "平";
                        sum2 = 0;
                    }
                    dt4.Rows.Add(dr2);


                    DataRow dr3 = dt4.NewRow();
                    dr3["摘要"] = "本年累计";
                    /*dr3["科目代码"] = dt.Rows[i]["科目代码"].ToString();
                    dr3["科目名称"] = dt.Rows[i]["科目名称"].ToString();
                    dr3["凭证号"] = dt.Rows[i]["凭证号"].ToString();*/
                    dr3["日期"] = dt.Rows[i]["本期结帐日期"].ToString();
                    if (sum > 0)
                    {
                        dr3["借方"] = sum;
                    }
                    else
                    {
                        dr3["借方"] = DBNull.Value;

                    }
                    if (sum1 > 0)
                    {
                        dr3["贷方"] = sum1;
                    }
                    else
                    {
                        dr3["贷方"] = DBNull.Value;
                    }

                    if (sum2 > 0)
                    {
                        dr3["余额"] = sum2;
                        dr3["方向"] = "借";

                    }
                    else if (sum2 < 0)
                    {
                        dr3["余额"] = -sum2;
                        dr3["方向"] = "贷";

                    }
                    else
                    {
                        dr3["余额"] = sum2;
                        dr3["方向"] = "平";


                    }
                    dt4.Rows.Add(dr3);

                }

            }
            return dt4;
        }
        #endregion
        #region GET_TABLEINFO
        public DataTable GET_TABLEINFO1(DataTable dt)
        {

            DataTable dt4 = this.GetTableInfo_O();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {

                    DataRow dr1 = dt4.NewRow();
                    if (dt.Rows[i]["期初借方"].ToString() != "" || dt.Rows[i]["期初贷方"].ToString() != "")
                    {
                        dr1["摘要"] = "上年结转";
                        dr1["科目代码"] = dt.Rows[i]["科目代码"].ToString();
                        dr1["科目名称"] = dt.Rows[i]["科目名称"].ToString();
                        dr1["凭证号"] = dt.Rows[i]["凭证号"].ToString();

                        dr1["日期"] = dt.Rows[i]["结转日期"].ToString();
                        decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0, d5 = 0, d6 = 0;
                        if (!string.IsNullOrEmpty(dt.Rows[i]["借方原币金额"].ToString()))
                        {
                            d1 = decimal.Parse(dt.Rows[i]["借方原币金额"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dt.Rows[i]["贷方原币金额"].ToString()))
                        {
                            d2 = decimal.Parse(dt.Rows[i]["贷方原币金额"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dt.Rows[i]["期初借方"].ToString()))
                        {
                            d3 = decimal.Parse(dt.Rows[i]["期初借方"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dt.Rows[i]["期初贷方"].ToString()))
                        {
                            d4 = decimal.Parse(dt.Rows[i]["期初贷方"].ToString());
                        }

                        d5 = d2 - d1 + d3 - d4;
                        d6 = d1 - d2 - d3 + d4;
                        if (d5 > 0)
                        {
                            dr1["余额"] = d5;
                            dr1["方向"] = dt.Rows[i]["方向"].ToString();
                        }

                        if (d6 > 0)
                        {
                            dr1["余额"] = d6;
                            dr1["方向"] = "贷";
                        }


                        dt4.Rows.Add(dr1);
                    }

                    DataRow dr2 = dt4.NewRow();
                    dr2["摘要"] = "本期合计";
                    if (dt.Rows[i]["期初借方"].ToString() != "" || dt.Rows[i]["期初贷方"].ToString() != "")
                    {
                    }
                    else
                    {
                        dr2["科目代码"] = dt.Rows[i]["科目代码"].ToString();
                        dr2["科目名称"] = dt.Rows[i]["科目名称"].ToString();
                        dr2["凭证号"] = dt.Rows[i]["凭证号"].ToString();
                    }
                    dr2["日期"] = dt.Rows[i]["本期结帐日期"].ToString();
                    if (!string.IsNullOrEmpty(dt.Rows[i]["借方原币金额"].ToString()))
                    {
                        dr2["借方"] = dt.Rows[i]["借方原币金额"].ToString();
                    }
                    else
                    {
                        dr2["借方"] = DBNull.Value;
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["贷方原币金额"].ToString()))
                    {
                        dr2["贷方"] = dt.Rows[i]["贷方原币金额"].ToString();
                    }
                    else
                    {
                        dr2["贷方"] = DBNull.Value;
                    }

                    dt4.Rows.Add(dr2);


                    DataRow dr3 = dt4.NewRow();
                    dr3["摘要"] = "原年累计";
                    /*dr3["科目代码"] = dt.Rows[i]["科目代码"].ToString();
                    dr3["科目名称"] = dt.Rows[i]["科目名称"].ToString();
                    dr3["凭证号"] = dt.Rows[i]["凭证号"].ToString();*/
                    dr3["日期"] = dt.Rows[i]["原期结帐日期"].ToString();
                    if (!string.IsNullOrEmpty(dt.Rows[i]["借方原币金额"].ToString()))
                    {

                        dr3["借方"] = decimal.Parse(dt.Rows[i]["借方原币金额"].ToString());
                    }
                    else
                    {
                        dr3["借方"] = DBNull.Value;
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["贷方原币金额"].ToString()))
                    {
                        dr3["贷方"] = dt.Rows[i]["贷方原币金额"].ToString();
                    }
                    else
                    {
                        dr3["贷方"] = DBNull.Value;
                    }

                    dt4.Rows.Add(dr3);

                }

            }
            return dt4;
        }
        #endregion
        #region GET_COLUMNS_INFO()
        public DataTable GET_COLUMNS_INFO()
        {
            DataTable dtt = new DataTable();
            dtt.Columns.Add("分类科目", typeof(string));
            dtt.Columns.Add("分类科目名称", typeof(string));
            dtt.Columns.Add("科目代码", typeof(string));
            dtt.Columns.Add("科目名称", typeof(string));
            dtt.Columns.Add("年初数", typeof(decimal));
            dtt.Columns.Add("期末数", typeof(decimal));
            dtt.Columns.Add("期初借方余额", typeof(decimal));
            dtt.Columns.Add("期初贷方余额", typeof(decimal));
            dtt.Columns.Add("借方余额", typeof(decimal));
            dtt.Columns.Add("贷方余额", typeof(decimal));
            dtt.Columns.Add("会计年度", typeof(string));
            dtt.Columns.Add("帐期", typeof(string));
            dtt.Columns.Add("币别编号", typeof(string));
            dtt.Columns.Add("币别", typeof(string));
            dtt.Columns.Add("名称", typeof(string));
            dtt.Columns.Add("汇率", typeof(decimal));
            dtt.Columns.Add("状态", typeof(string));
            dtt.Columns.Add("方向", typeof(string));
            dtt.Columns.Add("期初借方原币", typeof(decimal));
            dtt.Columns.Add("期初借方本币", typeof(decimal));
            dtt.Columns.Add("期初贷方原币", typeof(decimal));
            dtt.Columns.Add("期初贷方本币", typeof(decimal));
            dtt.Columns.Add("借方原币金额", typeof(decimal));
            dtt.Columns.Add("借方本币金额", typeof(decimal));
            dtt.Columns.Add("贷方原币金额", typeof(decimal));
            dtt.Columns.Add("贷方本币金额", typeof(decimal));
            dtt.Columns.Add("制单人", typeof(string));
            dtt.Columns.Add("制单日期", typeof(string));
            dtt.Columns.Add("项次", typeof(string));
            dtt.Columns.Add("摘要", typeof(string));
            dtt.Columns.Add("会计科目", typeof(string));
            dtt.Columns.Add("单价", typeof(string));
            dtt.Columns.Add("数量", typeof(string));
            return dtt;
        }
        #endregion
        #region GET_SUM
        public DataTable GET_SUM(string sqlcondition,string BILLNAME)
        {


            DataTable dtt = this.GET_COLUMNS_INFO();
            DataTable dtx6 = new DataTable();
            DataTable dtx61 =bc.getdt(sql + sqlcondition);
            DataView dvx = new DataView(dtx61);
            dvx.RowFilter = @"分类科目名称 IN (
'主营业务收入',
'其他业务收入',
'营业外收入',

'主营业务成本',
'营业税金及附加',
'其他业务成本',
'销售费用',
'管理费用',
'财务费用',
'营业外支出',
'所得税费用',

'投资收益',

'利润分配',
'本年利润'
)";
            if (BILLNAME == "资产负债表")
            {
                dtx6 = dtx61;
            }
            if(BILLNAME =="结转损益")
            {
                dtx6 = dvx.ToTable();
            }
            DataTable dtx7 = new DataTable();
            if (dtx6.Rows.Count > 0)
            {
                    foreach (DataRow dr1 in dtx6.Rows)
                    {
                        DataView dv = new DataView(dtt);
                        if (BILLNAME == "资产负债表")
                        {
                         
                            dv.RowFilter = "分类科目=" + dr1["分类科目"].ToString() + "";
                        }
                        if (BILLNAME == "结转损益")
                        {
                            dv.RowFilter = "科目代码=" + dr1["科目代码"].ToString() + "";
                        }
                      
                        DataTable  dtx8 = dv.ToTable();
                        if (dtx8.Rows.Count > 0)
                        {
                        }
                        else
                        {
                            //MessageBox.Show(dr1["分类科目"].ToString() + "," + dr1["科目代码"].ToString() + "," + dr1["期初借方原币"].ToString());
                            decimal d1 = 0, d2 = 0, d3 = 0,d4=0,d5=0,d6=0,d7=0,d8=0,d9 = 0, d10 = 0,d11=0,d12=0;
                            string v1, v2, v3, v4, v5, v6, v7, v8;
                            DataRow dr = dtt.NewRow();
                            dr["分类科目"] = dr1["分类科目"].ToString();
                            dr["科目代码"] = dr1["科目代码"].ToString();
                            dr["科目名称"] = dr1["科目名称"].ToString();
                            dr["方向"] = dr1["方向"].ToString();
                            dr["分类科目名称"] = dr1["分类科目名称"].ToString();
                            DataView dv1=new DataView (dtx6);
                            if (BILLNAME == "资产负债表")
                            {
                                dv1.RowFilter = "分类科目=" + dr1["分类科目"].ToString() + " AND 状态 IN ('INITIAL')";
                            }
                            if (BILLNAME == "结转损益")
                            {
                                dv1.RowFilter = "科目代码=" + dr1["科目代码"].ToString() + " AND 状态 IN ('INITIAL')";
                            }
                            dtx7=dv1.ToTable ();
                            if (dtx7.Rows.Count > 0)
                            {
                                v1 = Convert.ToString(dtx7.Compute("SUM(期初借方原币)", ""));
                                v2 = Convert.ToString(dtx7.Compute("SUM(期初借方本币)", ""));
                                v3 = Convert.ToString(dtx7.Compute("SUM(期初贷方原币)", ""));
                                v4 = Convert.ToString(dtx7.Compute("SUM(期初贷方本币)", ""));
                                if (!string.IsNullOrEmpty(v1))
                                {
                                    d1 = Convert.ToDecimal(dtx7.Compute("SUM(期初借方原币)", ""));
                                }
                                if (!string.IsNullOrEmpty(v2))
                                {
                                    d2 = Convert.ToDecimal(dtx7.Compute("SUM(期初借方本币)", ""));
                                
                                }
                                if (!string.IsNullOrEmpty(v3))
                                {
                                    d3 = Convert.ToDecimal(dtx7.Compute("SUM(期初贷方原币)", ""));
                                }
                                if (!string.IsNullOrEmpty(v3))
                                {
                                    d4 = Convert.ToDecimal(dtx7.Compute("SUM(期初贷方本币)", ""));
                                }
                              
                            }
                            dr["期初借方原币"] = d1;
                            dr["期初借方本币"] = d2;
                            dr["期初贷方原币"] = d3;
                            dr["期初贷方本币"] = d4;
                            DataView dv2 = new DataView(dtx6);
                            if (BILLNAME == "资产负债表")
                            {
                                dv2.RowFilter = "分类科目=" + dr1["分类科目"].ToString() + " AND 状态 IN ('OPEN','CARRY')";
                            }
                            if (BILLNAME == "结转损益")
                            {
                                dv2.RowFilter = "科目代码=" + dr1["科目代码"].ToString() + " AND 状态 IN ('OPEN','CARRY')";
                            }
                            DataTable  dtx9 = dv2.ToTable();
                            if (dtx9.Rows.Count > 0)
                            {
                                v5 = Convert.ToString(dtx9.Compute("SUM(借方原币金额)", ""));
                                v6 = Convert.ToString(dtx9.Compute("SUM(借方本币金额)", ""));
                                v7 = Convert.ToString(dtx9.Compute("SUM(贷方原币金额)", ""));
                                v8 = Convert.ToString(dtx9.Compute("SUM(贷方本币金额)", ""));
                                if (!string.IsNullOrEmpty(v5))
                                {
                                    d5 = Convert.ToDecimal(dtx9.Compute("SUM(借方原币金额)", ""));
                                }
                                if (!string.IsNullOrEmpty(v6))
                                {
                                    d6 = Convert.ToDecimal(dtx9.Compute("SUM(借方本币金额)", ""));

                                }
                                if (!string.IsNullOrEmpty(v7))
                                {
                                    d7 = Convert.ToDecimal(dtx9.Compute("SUM(贷方原币金额)", ""));
                                }
                                if (!string.IsNullOrEmpty(v8))
                                {
                                    d8 = Convert.ToDecimal(dtx9.Compute("SUM(贷方本币金额)", ""));
                                }
                         
                            
                            }
                            dr["借方原币金额"] = d5;
                            dr["借方本币金额"] = d6;
                            dr["贷方原币金额"] = d7;
                            dr["贷方本币金额"] = d8;
                            string v9 = dr["方向"].ToString();
                            if (v9 == "借")
                            {
                                d9 = d2 - d4;
                                d10 = d9 + d6 - d8;
                                d11 = d9 + d6;
                                d12 = d9 + d8;
                            }
                            else if (v9 == "贷")
                            {
                                d9 = d4 - d2;
                                d10 = d9 + d8 - d6;
                                d11 = d9 + d6;
                                d12 = d9 + d8;
                            }
                          
                            dr["年初数"] = d9;
                            dr["期末数"] = d10;
                            dr["期初借方余额"] = d2;
                            dr["期初贷方余额"] = d4;
                            dr["借方余额"] = d11;
                            dr["贷方余额"] = d12;
                            dtt.Rows.Add(dr);
                        
                        }
                    }
                
            }
            return dtt;
        }
        #endregion
        #region BALANCE
        public DataTable BALANCE()
        {

            DataTable dtX = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE WHERE PARENT_NODEID IS NULL");
            DataTable dt4 = this.GET_COLUMNS_INFO();
            if (dtX.Rows.Count > 0)
            {
               
                foreach (DataRow dr in dtX.Rows)
                {

                    DataRow dr1 = dt4.NewRow();
                    dr1["分类科目"] = dr["ACCODE"].ToString();
                    dr1["分类科目名称"] = dr["ACNAME"].ToString();
                    dr1["科目代码"] = dr["ACCODE"].ToString();
                    dr1["科目名称"] = dr["ACNAME"].ToString();
                    dt4.Rows.Add(dr1);

                }
            }
            return dt4;
        }
        #endregion
        #region UPDATE_COURSE_BALANCE
        public DataTable UPDATE_COURSE_BALANCE(string sqlcondition)
        {

            DataTable dtX = GET_SUM(sqlcondition,"资产负债表");
            DataTable dt4 = BALANCE();
            if (dtX.Rows.Count > 0)
            {
                if (dt4.Rows.Count > 0)
                {
                    foreach (DataRow dr1 in dtX.Rows)
                    {

                        foreach (DataRow dr in dt4.Rows)
                        {

                            if (dr["分类科目"].ToString ()== dr1["分类科目"].ToString ())
                            {

                                
                                dr["年初数"] = dr1["年初数"].ToString();
                                dr["期末数"] = dr1["期末数"].ToString();
                                dr["期初借方余额"] = dr1["期初借方余额"].ToString();
                                dr["期初贷方余额"] = dr1["期初贷方余额"].ToString();
                                dr["借方余额"] = dr1["借方余额"].ToString();
                                dr["贷方余额"] = dr1["贷方余额"].ToString();
                                break;
                            }

                        }

                    }
                }
            }
            return dt4;
        }
        #endregion
        #region ExcelPrint
        public void ExcelPrint(string sqlcondition,string Printpath)
        {
           
            DataTable dt2 = this.UPDATE_COURSE_BALANCE(sqlcondition);
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;


            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);
            worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            application.Visible = false;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;

            DataView dv = new DataView(dt2);
            dv.RowFilter = "分类科目名称 IN ('银行存款','库存现金','其他货币资金')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[7,3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[7,5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('交易性金融资产')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[8, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[8, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应收票据')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[9, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[9, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应收股利')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[10, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[10, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应收利息')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[11, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[11, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应收利息')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[11, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[11, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应收账款','预收账款','坏账准备')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "应收账款")
                    {
                        if (!string.IsNullOrEmpty(dr["期初借方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初借方余额"].ToString());
                     
                        }
                        if (!string.IsNullOrEmpty(dr["借方余额"].ToString()))
                        {
                        
                            d2 = d2 + decimal.Parse(dr["借方余额"].ToString());
                        }
                      
                    }
                    if (dr["分类科目名称"].ToString() == "预收账款")
                    {
                        if (!string.IsNullOrEmpty(dr["期初借方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初借方余额"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["借方余额"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["借方余额"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "坏账准备")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 - decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 - decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                   
                }
             
                worksheet.Cells[12, 3] = d1;
                worksheet.Cells[12, 5] = d2;
            }
            dv.RowFilter = "分类科目名称 IN ('其他应收款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[13, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[13, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('其他应收款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[13, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[13, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('预付账款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[14, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[14, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = @"分类科目名称 IN (
'材料采购',
'在途物资',
'原材料',
'库存商品',
'周转材料',
'委托加工物资',
'委托代销商品',
'发出商品',
'生产成本',
'存货跌价准备',
'材料成本差异',
'商品进销差价')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "材料采购")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
             
                    if (dr["分类科目名称"].ToString() == "在途物资")
                    {
                    

                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 /decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 / decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "原材料")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "库存商品")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "周转材料")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "委托加工物资")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "委托代销商品")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }

                    if (dr["分类科目名称"].ToString() == "发出商品")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "生产成本")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "存货跌价准备")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 - decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 - decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "材料成本差异")
                    {
                        if (!string.IsNullOrEmpty(dr["期初借方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初借方余额"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["借方余额"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["借方余额"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "商品进销差价")
                    {
                        if (!string.IsNullOrEmpty(dr["期初借方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初借方余额"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["借方余额"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["借方余额"].ToString());
                        }
                    }
                  
                }

                worksheet.Cells[15, 3] = d1;
                worksheet.Cells[15, 5] = d2;
            }
            /*dv.RowFilter = "分类科目名称 IN ('持有至到期投资','长期应收款','长期待摊费用')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "持有至到期投资")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "长期应收款")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "长期待摊费用")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }

                }

                worksheet.Cells[16, 3] = d1;
                worksheet.Cells[16, 5] = d2;
            }*/
            dv.RowFilter = "分类科目名称 IN ('可供出售金融资产')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[20, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[20, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('持有至到期投资','持有至到期投资减值准备')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "持有至到期投资")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "持有至到期投资减值准备")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 - decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 - decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                 

                }

                worksheet.Cells[21, 3] = d1;
                worksheet.Cells[21, 5] = d2;
            }
            dv.RowFilter = "分类科目名称 IN ('投资性房地产')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[22, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[22, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('长期股权投资','长期股权投资减值准备')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "长期股权投资")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "长期股权投资减值准备")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 - decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 - decimal.Parse(dr["期末数"].ToString());
                        }
                    }


                }

                worksheet.Cells[21, 3] = d1;
                worksheet.Cells[21, 5] = d2;
            }
            dv.RowFilter = "分类科目名称 IN ('固定资产')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[26, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[26, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('累计折旧')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[27, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[27, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('固定资产减值准备')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[29, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[29, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('固定资产清理')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[30, 3] = dt.Compute("SUM(期初借方余额)", "");
                worksheet.Cells[30, 5] = dt.Compute("SUM(借方余额)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('在建工程')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[31, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[31, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('无形资产','累计摊销','无形资产减值准备')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "无形资产")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "累计摊销")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 - decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 - decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "无形资产减值准备")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 - decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 - decimal.Parse(dr["期末数"].ToString());
                        }
                    }

                }

                worksheet.Cells[35, 3] = d1;
                worksheet.Cells[35, 5] = d2;
            }
            dv.RowFilter = "分类科目名称 IN ('递延所得税资产')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[36, 3] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[36, 5] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('短期借款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[7, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[7, 13] = dt.Compute("SUM(期末数)", "");
              
            }
            dv.RowFilter = "分类科目名称 IN ('应付票据')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[8, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[8, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应付账款','预付账款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "应付账款")
                    {
                        if (!string.IsNullOrEmpty(dr["期初贷方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初贷方余额"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["贷方余额"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["贷方余额"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "预付账款")
                    {
                        if (!string.IsNullOrEmpty(dr["期初贷方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初贷方余额"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["贷方余额"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["贷方余额"].ToString());
                        }
                    }
                }
                worksheet.Cells[9, 12] = d1;
                worksheet.Cells[9, 13] = d2;
            }
            dv.RowFilter = "分类科目名称 IN ('预收账款','应收账款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "预收账款")
                    {
                        if (!string.IsNullOrEmpty(dr["期初贷方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初贷方余额"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["贷方余额"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["贷方余额"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "应收账款")
                    {
                        if (!string.IsNullOrEmpty(dr["期初贷方余额"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["期初贷方余额"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["贷方余额"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["贷方余额"].ToString());
                        }
                    }
                }
                worksheet.Cells[10, 12] = d1;
                worksheet.Cells[10, 13] = d2;
            }
            dv.RowFilter = "分类科目名称 IN ('其他应付款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[11, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[11, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应付职工薪酬')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[12, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[12, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应付股利')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[13, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[13, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应交税费')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[14, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[14, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('长期借债')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[19, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[19, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('应付债卷')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[20, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[20, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('长期应付款')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[21, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[21, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('递延收益')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[24, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[24, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('实收资本')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[27, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[27, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('资本公积')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[28, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[28, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('盈余公积')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[29, 12] = dt.Compute("SUM(年初数)", "");
                worksheet.Cells[29, 13] = dt.Compute("SUM(期末数)", "");
            }
            dv.RowFilter = "分类科目名称 IN ('本年利润','利润分配')";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {
                decimal d1 = 0, d2 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["分类科目名称"].ToString() == "本年利润")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                    }
                    if (dr["分类科目名称"].ToString() == "利润分配")
                    {
                        if (!string.IsNullOrEmpty(dr["年初数"].ToString()))
                        {
                            d1 = d1 + decimal.Parse(dr["年初数"].ToString());

                        }
                        if (!string.IsNullOrEmpty(dr["期末数"].ToString()))
                        {

                            d2 = d2 + decimal.Parse(dr["期末数"].ToString());
                        }
                       
                    }
                }
                worksheet.Cells[31, 12] = d1;
                worksheet.Cells[31, 13] = d2;
                
            }
            workbook.Save();
            application.Quit();
            worksheet = null;
            workbook = null;
            application = null;
            GC.Collect();
            System.Diagnostics.Process.Start(Printpath);
        }
        #endregion
    }
}
