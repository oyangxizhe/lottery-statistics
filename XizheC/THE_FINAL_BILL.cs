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
    public class THE_FINAL_BILL
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




        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        List<string> list = new List<string>();
        List<string> listo = new List<string>();
        THE_BALANCE_SHEET the_balance_sheet = new THE_BALANCE_SHEET();
      
        string[] xw=new string[]{
"流动资产",
"货币资金",
"交易性金融资产"
        };
        
        string[] xwo = new string[]{
            "",
            ""};
        public THE_FINAL_BILL()
        {
            IFExecution_SUCCESS = true;
          
            
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

            dt = basec.getdts(the_balance_sheet.getsql + sql1);
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
  
        #region GET_SUM
        public DataTable GET_SUM(string sqlcondition)
        {


            DataTable dtt = the_balance_sheet.GET_SUM(sqlcondition, "结转损益");
        
            
            return dtt;
        }
        #endregion
        #region BALANCE
        public DataTable BALANCE()
        {

            DataTable dtX = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE WHERE PARENT_NODEID IS NULL");
            DataTable dt4 = the_balance_sheet.GET_COLUMNS_INFO() ;
            if (dtX.Rows.Count > 0)
            {
               
                foreach (DataRow dr in dtX.Rows)
                {

                    DataRow dr1 = dt4.NewRow();
                    dr1["科目"] = dr["ACCODE"].ToString();
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

            DataTable dtX = GET_SUM(sqlcondition);
            DataTable dt4 = BALANCE();
            if (dtX.Rows.Count > 0)
            {
                if (dt4.Rows.Count > 0)
                {
                    foreach (DataRow dr1 in dtX.Rows)
                    {

                        foreach (DataRow dr in dt4.Rows)
                        {

                            if (dr["科目"].ToString ()== dr1["科目"].ToString ())
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
        #region MAKER_VOUCHER
        public void   MAKER_VOUCHER(string sqlcondition)
        {

            CVOUCHER vou = new CVOUCHER();
            PERIOD period = new PERIOD();
            DataTable dt2 = this.CLASSIFY_VOUCHER(sqlcondition);

            if (vou.NUM_ID != "")
            {
                vou.VOUCHER_DATE = period.CURRENT_PERIOD_EXPIRATION_DATE;
                vou.ACCOUNTING_PERIOD_EXPIRATION_DATE = period.CURRENT_PERIOD_EXPIRATION_DATE;
                vou.save("VOUCHER_MST", "VOUCHER_DET", "VOID", vou.NUM_ID, dt2);
                MessageBox.Show("当期损益已经结转！生产的结转凭证单号为："+vou.NUM_ID , "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion
        #region CLASSIFY_VOUCHER
        public DataTable  CLASSIFY_VOUCHER(string sqlcondition)
        {
            PERIOD period = new PERIOD();
            CVOUCHER vou = new CVOUCHER();
            DataTable dtt = the_balance_sheet.GET_COLUMNS_INFO();
            DataTable dt2 = the_balance_sheet.GET_SUM(sqlcondition, "结转损益");
            DataView dv = new DataView(dt2);
            dv.RowFilter = @"分类科目名称 IN (
'主营业务收入',
'其他业务收入',
'营业外收入')";
            dt = dv.ToTable();
            int i = 1;
            if (dt.Rows.Count > 0)
            {
                
                foreach (DataRow dr1 in dt.Rows)
                {

                    DataRow dr = dtt.NewRow();
                    dr["会计科目"] = dr1["科目代码"].ToString();
                    dr["科目名称"] = dr1["科目名称"].ToString();
                    dr["分类科目名称"] = dr1["分类科目名称"].ToString();
                    dr["借方本币金额"] = dr1["期末数"].ToString();
                    dr["项次"] = i;
                    dtt.Rows.Add(dr);
                    i = i + 1;
                   
                }
                DataRow dr2 = dtt.NewRow();
                dr2["会计科目"] = "4103";
                dr2["贷方本币金额"] = dt.Compute("SUM(期末数)", "");
                dr2["摘要"] = "结转收入";
                dr2["项次"] = i;
                dtt.Rows.Add(dr2);
                i = i + 1;
      
            }

            dv.RowFilter = @"分类科目名称 IN (
'主营业务成本',
'营业税金及附加',
'其他业务成本',
'销售费用',
'管理费用',
'财务费用',
'营业外支出',
'所得税费用')";
            dt = dv.ToTable();
           
            if (dt.Rows.Count > 0)
            {

                foreach (DataRow dr1 in dt.Rows)
                {

                    DataRow dr = dtt.NewRow();
                    dr["会计科目"] = dr1["科目代码"].ToString();
                    dr["科目名称"] = dr1["科目名称"].ToString();
                    dr["分类科目名称"] = dr1["分类科目名称"].ToString();
                    dr["贷方本币金额"] = dr1["期末数"].ToString();
                    dr["项次"] = i;
                    dtt.Rows.Add(dr);
                    i = i + 1;

                }
                DataRow dr2 = dtt.NewRow();


                dr2["科目名称"] = "本年利润";
                dr2["会计科目"] = "4103";
                dr2["借方本币金额"] = dt.Compute("SUM(期末数)", "");
                dr2["项次"] = i;
                dr2["摘要"] = "结转成本 费用 营业税金 所得税";
                dtt.Rows.Add(dr2);
                i = i + 1;

            }
            dv.RowFilter = @"分类科目名称 IN (
'投资收益'
)";
            dt = dv.ToTable();
            if (dt.Rows.Count > 0)
            {

                foreach (DataRow dr1 in dt.Rows)
                {

                    DataRow dr = dtt.NewRow();
                    dr["会计科目"] = dr1["科目代码"].ToString();
                    dr["科目名称"] = dr1["科目名称"].ToString();
                    dr["分类科目名称"] = dr1["分类科目名称"].ToString();
                    if (decimal.Parse(dr1["期末数"].ToString()) > 0)
                    {
                        dr["借方本币金额"] = dr1["期末数"].ToString();
                    }
                    else
                    {
                        dr["贷方本币金额"] = dr1["期末数"].ToString();
                    }
                    dr["项次"] = i;
                    dtt.Rows.Add(dr);
                    i = i + 1;

                }
                DataRow dr2 = dtt.NewRow();
                dr2["会计科目"] = "4103";
                if (decimal.Parse(dt.Rows [0]["期末数"].ToString()) > 0)
                {
                    dr2["贷方本币金额"] = dt.Rows[0]["期末数"].ToString();
                   
                }
                else
                {
                    dr2["借方本币金额"] = dt.Rows[0]["期末数"].ToString();
                  
                }
          
                dr2["项次"] = i;
                dr2["摘要"] = "结转投资收益";
                dtt.Rows.Add(dr2);
                i = i + 1;

            }
        
            if (period.GETPERIOD != "12")
            {
                

            }
            else
            {

                dt2 = the_balance_sheet.UPDATE_COURSE_BALANCE(" WHERE  B.FINANCIAL_YEAR='" + period.FINANCIAL_YEAR +
                    "'  AND B.STATUS IN ('OPEN','INITIAL','CARRY')  ORDER BY C.ACCODE ASC ");
                dv = new DataView(dt2);
                dv.RowFilter = @"分类科目名称 IN (
'利润分配',
'本年利润'
)";
                dt = dv.ToTable();
                if (dt.Rows.Count > 0)
                {
                    decimal d1 = 0, d2 = 0,d3=0;
                    string v1, v2;
                    DataView dv1 = new DataView(dtt);
                    dv1.RowFilter = "科目名称='本年利润'";
                    DataTable dt3 = dv1.ToTable();
                    if (dt3.Rows.Count > 0)
                    {
                        v1 = Convert.ToString(dt3.Compute("SUM(借方本币金额)", ""));
                        v2 = Convert.ToString(dt3.Compute("SUM(贷方本币金额)", ""));
                        if (!string.IsNullOrEmpty(v1))
                        {
                            d1 = decimal.Parse(v1);
                        }
                        if (!string.IsNullOrEmpty(v2))
                        {
                            d2 = decimal.Parse(v2);
                        }
                   

                    }
         
                    foreach (DataRow dr1 in dt.Rows)
                    {
                        
                        if (dr1["分类科目名称"].ToString() == "本年利润")
                        {
                            d3 = decimal.Parse(dr1["期末数"].ToString());
                            DataRow dr = dtt.NewRow();
                            dr["会计科目"] = dr1["科目代码"].ToString();
                            dr["科目名称"] = dr1["科目名称"].ToString();
                            dr["分类科目名称"] = dr1["分类科目名称"].ToString();
                            if (d3+d2-d1 > 0)
                            {
                                dr["借方本币金额"] = d3+d2-d1;
                            }

                            else
                            {
                                dr["贷方本币金额"] = d3 + d2 - d1;
                            }
                            dr["项次"] = i;
                            dtt.Rows.Add(dr);
                            i = i + 1;
                        }
                        if (dr1["分类科目名称"].ToString() == "利润分配")
                        {
                           
                            DataRow dr2 = dtt.NewRow();
                            dr2["会计科目"] = dr1["科目代码"].ToString();
                            dr2["科目名称"] = dr1["科目名称"].ToString();
                            dr2["分类科目名称"] = dr1["分类科目名称"].ToString();
                            if (d3 + d2 - d1 > 0)
                            {
                                dr2["贷方本币金额"] = d3 + d2 - d1;

                            }
                            else
                            {
                                dr2["借方本币金额"] = d3 + d2 - d1;

                            }

                            dr2["项次"] = i;
                            dr2["摘要"] = "结转利润分配";
                            dtt.Rows.Add(dr2);
                            i = i + 1;
                        }


                    }


                }

               
            }
            return dtt;
     
        }
        #endregion
    }
}
