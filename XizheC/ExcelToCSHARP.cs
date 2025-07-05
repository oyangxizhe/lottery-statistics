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
    public class ExcelToCSHARP
    {
        #region nature
        private string _sql;
        public string sql
        {
            set { _sql = value; }
            get { return _sql; ; }

        }
        public string sqlo { set; get; }
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
        private string _DBID;
        public string DBID
        {
            set { _DBID = value; }
            get { return _DBID; }
        }
        private string _ACID;
        public string ACID
        {
            set { _ACID = value; }
            get { return _ACID; }
        }
        private string _ACCODE;
        public string ACCODE
        {
            set { _ACCODE = value; }
            get { return _ACCODE; }


        }
        private string _EMID;
        public  string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _PERIOD;
        public string PERIOD
        {
            set { _PERIOD = value; }
            get { return _PERIOD; }

        }
        private string _RED_BALL_ONE;
        public string RED_BALL_ONE
        {
            set { _RED_BALL_ONE = value; }
            get { return _RED_BALL_ONE; }

        }
        private string _RED_BALL_TWO;
        public string RED_BALL_TWO
        {
            set { _RED_BALL_TWO = value; }
            get { return _RED_BALL_TWO; }

        }
        private string _RED_BALL_THREE;
        public string RED_BALL_THREE
        {
            set { _RED_BALL_THREE = value; }
            get { return _RED_BALL_THREE; }

        }
        private string _RED_BALL_FOUR;
        public string RED_BALL_FOUR
        {
            set { _RED_BALL_FOUR = value; }
            get { return _RED_BALL_FOUR; }

        }
        private string _RED_BALL_FIVE;
        public string RED_BALL_FIVE
        {
            set { _RED_BALL_FIVE = value; }
            get { return _RED_BALL_FIVE; }

        }
        private string _RED_BALL_SIX;
        public string RED_BALL_SIX
        {
            set { _RED_BALL_SIX = value; }
            get { return _RED_BALL_SIX; }

        }
        private string _BLUE_BALL;
        public string BLUE_BALL
        {
            set { _BLUE_BALL = value; }
            get { return _BLUE_BALL; }

        }
        private bool _IfFirstDetailCourse;
        public bool IfFirstDetailCourse
        {
            set { _IfFirstDetailCourse = value; }
            get { return _IfFirstDetailCourse; }
        }
        private bool _IFCONSULENZA;
        public bool IFCONSULENZA
        {
            set { _IFCONSULENZA = value; }
            get { return _IFCONSULENZA; }
        }
        private string _hint;
        public string hint
        {
            set { _hint = value; }
            get { return _hint; }
        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        #endregion
        string setsql = @"
SELECT 
PERIOD AS 期数,
RED_BALL_ONE AS 红球一,
RED_BALL_TWO AS 红球二,
RED_BALL_THREE AS 红球三,
RED_BALL_FOUR AS 红球四,
RED_BALL_FIVE AS 红球五,
RED_BALL_SIX AS 红球六,
BLUE_BALL AS 蓝球
FROM DOUBLE_BALL A
";
        string sql1 = @"INSERT INTO DOUBLE_BALL(
PERIOD,
RED_BALL_ONE,
RED_BALL_TWO,
RED_BALL_THREE,
RED_BALL_FOUR,
RED_BALL_FIVE,
RED_BALL_SIX,
BLUE_BALL
) 
VALUES 
(
@PERIOD,
@RED_BALL_ONE,
@RED_BALL_TWO,
@RED_BALL_THREE,
@RED_BALL_FOUR,
@RED_BALL_FIVE,
@RED_BALL_SIX,
@BLUE_BALL
)

";
        /*string sql2 = @"UPDATE DOUBLE_BALL SET 
DBID=@DBID,
PERIOD=@PERIOD,
RED_BALL_ONE=@RED_BALL_ONE,
RED_BALL_TWO=@RED_BALL_TWO,
RED_BALL_THREE=@RED_BALL_THREE,
RED_BALL_FOUR=@RED_BALL_FOUR,
RED_BALL_FIVE=@RED_BALL_FIVE,
RED_BALL_SIX=@RED_BALL_SIX,
BLUE_BALL=@BLUE_BALL,
DATE=@DATE,
MAKERID=@MAKERID
";*/
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        StringBuilder sqb = new StringBuilder();
        public ExcelToCSHARP()
        {
            IFExecution_SUCCESS = true;
            sql = setsql;
            sqlo = sql1;
        }

     
        #region importExcelToDataSet
        public static DataSet importExcelToDataSet(string FilePath, string tablename)
        {
            string strConn;
            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + FilePath + ";Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            OleDbDataAdapter myCommand = new OleDbDataAdapter("SELECT * FROM [" + tablename + "] ", strConn);
            DataSet myDataSet = new DataSet();
            try
            {
                myCommand.Fill(myDataSet);
            }
            catch (Exception ex)
            {
                MessageBox.Show("error," + ex.Message);
            }
            return myDataSet;
        }
        #endregion
        #region GetExcelFirstTableName
        public static string GetExcelFirstTableName(string excelFileName)
        {
            string tableName = null;
            if (File.Exists(excelFileName))
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet." +
                  "OLEDB.4.0;Extended Properties=\"Excel 8.0\";Data Source=" + excelFileName))
                {
                    conn.Open();
                    DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    tableName = dt.Rows[0][2].ToString().Trim();

                }
            }
            return tableName;
        }
        #endregion

        
 

        #region save
        public void save(string ACID, string s,string ACCODE, string ACNAME, string COURSE_TYPE, string BALANCE_DIRECTION,string CYCODE,string COURSE_NATURE)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.getOnlyString("SELECT ACCODE FROM Accountant_Course WHERE  ACID='" + ACID + "'");
            string v2 = bc.getOnlyString("SELECT ACNAME FROM Accountant_Course WHERE  ACID='" + ACID + "'");
            //string v3 = "NULL";
            //string varMakerID;
            if (!bc.exists("SELECT ACID FROM Accountant_Course WHERE ACID='" + ACID + "'"))
            {
                if (bc.exists("SELECT * FROM Accountant_Course WHERE ACCODE='" + ACCODE + "'"))
                {
                    IFExecution_SUCCESS = false;
                   
                    hint = "科目代码已经存在于系统！";

                }
                else if (bc.exists("SELECT * FROM Accountant_Course WHERE  ACNAME='" + ACNAME + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "科目名称已经存在于系统！";

                }
                else
                {
                    IFExecution_SUCCESS = true;

                    //SQlcommandE(sql1, ACID, ACCODE, ACNAME, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE,COURSE_NATURE );
                    ADD_OR_UPDATE = "ADD";
                }

            }
        
            else if (v1 != ACCODE && v2 == ACNAME)
            {
                if (bc.exists("SELECT * FROM Accountant_Course WHERE ACCODE='" + ACCODE + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "科目代码已经存在于系统！";

                }
                else
                {
                    IFExecution_SUCCESS = true;
                    //SQlcommandE(sql2 + " WHERE ACID='" + ACID + "'", ACID, ACCODE, ACNAME, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                    ADD_OR_UPDATE = "UPDATE";

                }
            }
            else if (v1 == ACCODE && v2 != ACNAME)
            {
                if (bc.exists("SELECT * FROM Accountant_Course WHERE ACNAME='" + ACNAME + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "科目名称已经存在于系统！";

                }
                else
                {
                    IFExecution_SUCCESS = true;
                    //SQlcommandE(sql2 + " WHERE ACID='" + ACID + "'", ACID, ACCODE, ACNAME, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                    ADD_OR_UPDATE = "UPDATE";

                }
            }
            else if (v1 != ACCODE && v2 != ACNAME)
            {
                if (bc.exists("SELECT * FROM Accountant_Course WHERE ACCODE='" + ACCODE + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "科目代码已经存在于系统！";

                }
                else if (bc.exists("SELECT * FROM Accountant_Course WHERE  ACNAME='" + ACNAME + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "科目名称已经存在于系统！";

                }
                else
                {
                    /*IFExecution_SUCCESS = true;
                    SQlcommandE(sql2 + " WHERE ACID='" + ACID + "'", ACID, ACCODE, ACNAME, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                    ADD_OR_UPDATE = "UPDATE";
                    */
                }
            }
            else
            {
                /*IFExecution_SUCCESS = true;
                SQlcommandE(sql2 + " WHERE ACID='" + ACID + "'", ACID, ACCODE, ACNAME, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                ADD_OR_UPDATE = "UPDATE";*/


            }
        }
        #endregion
       
       
   
     
   
 
   
 
        #region CheckKeyInValueIfExistsDetailCourse
        public int CheckKeyInValueIfExistsDetailCourse(string TABLENAME, string COLUMN_NAME, string COLUMN_VALUE, string REMARK,string REMARKT)
        {
            int ju = 0;
            int len = COLUMN_VALUE.Length;
            int len1 = len + 3;
            DataTable dt = bc.getdt("SELECT *  FROM " + TABLENAME + " WHERE SUBSTRING(" + COLUMN_NAME + ",1," + len + 
                ")='"+COLUMN_VALUE+"'"+" AND LEN("+COLUMN_NAME+")="+len1);
           
            if (dt.Rows.Count == 1)
            {
                ju = 1;
                MessageBox.Show(REMARK + " " + COLUMN_VALUE + REMARKT , "提示", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            }
            else if (dt.Rows.Count > 1)
            {
                ju = 2;
                MessageBox.Show(REMARK + " " + COLUMN_VALUE + REMARKT, "提示", MessageBoxButtons.OK,
               MessageBoxIcon.Information);
            }
 
            return ju;
        }
        #endregion
    
      
        #region CHECK_DATATABLE_IF_EXISTS_DETAIL_COURSE()
        public bool CHECK_DATATABLE_IF_EXISTS_DETAIL_COURSE(DataTable dt)
        {
            bool b = false;

            for (int k = 0; k < dt.Rows.Count; k++)
            {
                if (juage(k,dt))
                {
                    b = true;
                    break;
                }
            }
            return b;
        }
        #endregion
        #region juage()
        private bool juage(int k,DataTable dt)
        {
            bool b = false;
            string v1 = dt.Rows[k]["科目代码"].ToString();
            string v2 = dt.Rows[k]["累计借方"].ToString();
            string v3 = dt.Rows[k]["累计贷方"].ToString();
            string v4 = dt.Rows[k]["期初借方"].ToString();
            string v5 = dt.Rows[k]["期初贷方"].ToString();

            if ((v2 != "" || v3 != "" || v4 != "" || v5 != "") &&
                CheckKeyInValueIfExistsDetailCourse("ACCOUNTANT_COURSE", "ACCODE", v1, "科目代码", "存在明细科目，需使用明细科目记帐！") == 1)
            {
                b = true;
            }
            return b;
        }
        #endregion
        #region showdata
        public void showdata(string path)
        {
            DataSet ds = new DataSet();
            string tablename = ExcelToCSHARP.GetExcelFirstTableName(path);
            ds = ExcelToCSHARP.importExcelToDataSet(path, tablename);
            DataTable dt = ds.Tables[0];
            //dt = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "F1 IS NOT NULL");
     
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //MessageBox.Show(dt.Rows[i][3].ToString());
                //DBID  = bc.numYM(10, 4, "0001", "select * from double_ball", "DBID", "DB");
                PERIOD = dt.Rows[i][0].ToString();
                if (dt.Rows[i][1].ToString().Length == 20)
                { //string a = "07,09,11,15,18,25|07";
                    RED_BALL_ONE = dt.Rows[i][1].ToString().Substring(0, 2);
                    RED_BALL_TWO = dt.Rows[i][1].ToString().Substring(3, 2);
                    RED_BALL_THREE = dt.Rows[i][1].ToString().Substring(6, 2);
                    RED_BALL_FOUR = dt.Rows[i][1].ToString().Substring(9, 2);
                    RED_BALL_FIVE = dt.Rows[i][1].ToString().Substring(12, 2);
                    RED_BALL_SIX = dt.Rows[i][1].ToString().Substring(15, 2);//160122
                    BLUE_BALL = dt.Rows[i][1].ToString().Substring(18, 2);
                    /*sqb = new StringBuilder();
                    sqb.AppendFormat("第 {0} 期 ", dt.Rows[i][0].ToString());
                    sqb.AppendFormat("红球一为：{0} ", RED_BALL_ONE);
                    sqb.AppendFormat("红球二为：{0} ", RED_BALL_TWO);
                    sqb.AppendFormat("红球三为：{0} ", RED_BALL_THREE);
                    sqb.AppendFormat("红球四为：{0} ", RED_BALL_FOUR);
                    sqb.AppendFormat("红球五为：{0} ", RED_BALL_FIVE);
                    sqb.AppendFormat("红球六为：{0} ", RED_BALL_SIX);
                    sqb.AppendFormat("蓝球为：{0} ", BLUE_BALL);
                    MessageBox.Show(sqb.ToString());*/
                 
                }
                
                if (dt.Rows[i][1].ToString().Length != 20)
                {

                }
                else if (JuageACCODEFormat(i))
                {
                }
                else
                {   //注意这里如若导入的数据量多要通过stringbuilder拼接sql语句然后批量写入数据库，sql语句的执行不要写在for循环内部
                    //我这里因为每次不写多太多号码所以没有做批量写入。量多频繁调数据库连接不行。
                    SQlcommandE(sql1);
                }
            }

        }
        #endregion
        #region JuageACCODEFormat()
        public bool JuageACCODEFormat(int i)
        {

            bool b = false;
 
            if (bc.exists ("SELECT * FROM DOUBLE_BALL WHERE PERIOD='"+PERIOD +"'"))
            {

                b = true;
                //MessageBox.Show("第" + i + "行" + "期数已经存在系统！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else   if (RED_BALL_ONE == "")
            {

                b = true;
                MessageBox.Show("第" + i + "行" + "红色球1不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
           
            }
            else if (bc.yesno1(RED_BALL_ONE) == 0)
            {
                b = true;
                MessageBox.Show("第" + i + "行" + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (RED_BALL_TWO == "")
            {

                b = true;
                MessageBox.Show("第" + i + "行" + "红色球2不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (bc.yesno1(RED_BALL_TWO) == 0)
            {
                b = true;
                MessageBox.Show("第" + i + "行" + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (RED_BALL_THREE == "")
            {

                b = true;
                MessageBox.Show("第" + i + "行" + "红色球3不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (bc.yesno1(RED_BALL_THREE) == 0)
            {
                b = true;
                MessageBox.Show("第" + i + "行" + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (RED_BALL_FOUR == "")
            {

                b = true;
                MessageBox.Show("第" + i + "行" + "红色球4不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (bc.yesno1(RED_BALL_FOUR) == 0)
            {
                b = true;
                MessageBox.Show("第" + i + "行" + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (RED_BALL_FIVE == "")
            {

                b = true;
                MessageBox.Show("第" + i + "行" + "红色球5不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (bc.yesno1(RED_BALL_FIVE) == 0)
            {
                b = true;
                MessageBox.Show("第" + i + "行" + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (RED_BALL_SIX == "")
            {

                b = true;
                MessageBox.Show("第" + i + "行" + "红色球6不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (bc.yesno1(RED_BALL_SIX) == 0)
            {
                b = true;
                MessageBox.Show("第" + i + "行" + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (BLUE_BALL == "")
            {

                b = true;
                MessageBox.Show("第" + i + "行" + "蓝色球不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (bc.yesno1(BLUE_BALL) == 0)
            {
                b = true;
                MessageBox.Show("第" + i + "行" + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return b;
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            string varMakerID = "123";
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            //sqlcom.Parameters.Add("@DBID", SqlDbType.VarChar, 20).Value = DBID;
            sqlcom.Parameters.Add("@PERIOD", SqlDbType.VarChar, 20).Value =PERIOD ;
            sqlcom.Parameters.Add("@RED_BALL_ONE", SqlDbType.VarChar, 20).Value =RED_BALL_ONE ;
            sqlcom.Parameters.Add("@RED_BALL_TWO", SqlDbType.VarChar, 20).Value =RED_BALL_TWO ;
            sqlcom.Parameters.Add("@RED_BALL_THREE", SqlDbType.VarChar, 20).Value = RED_BALL_THREE ;
            sqlcom.Parameters.Add("@RED_BALL_FOUR", SqlDbType.VarChar, 20).Value =RED_BALL_FOUR ;
            sqlcom.Parameters.Add("@RED_BALL_FIVE", SqlDbType.VarChar, 20).Value =RED_BALL_FIVE ;
            sqlcom.Parameters.Add("@RED_BALL_SIX", SqlDbType.VarChar, 20).Value = RED_BALL_SIX ;
            sqlcom.Parameters.Add("@BLUE_BALL", SqlDbType.VarChar, 20).Value =BLUE_BALL ;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
         
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
    }
}
