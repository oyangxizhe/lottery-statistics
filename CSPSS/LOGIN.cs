using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using XizheC;

namespace CSPSS
{
    public partial class LOGIN : Form
    {
        private static string _USID;
        public static string USID
        {
            set { _USID = value; }
            get { return _USID; }

        }
        private static string _UNAME;
        public static string UNAME
        {
            set { _UNAME = value; }
            get { return _UNAME; }

        }
        private static string _EMID;
        public static string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private static string _ENAME;
        public static string ENAME
        {
            set { _ENAME = value; }
            get { return _ENAME; }

        }
        private static string _DEPART;
        public static string DEPART
        {
            set { _DEPART = value; }
            get { return _DEPART; }

        }
      
      
        public byte[] PWD;
        basec bc = new basec();
        CUSER cuser = new CUSER();
        public LOGIN()
        {
            InitializeComponent();
        }
        DataTable dt = new DataTable();
       
        private void LOGIN_Load(object sender, EventArgs e)
        {
            
            dt = bc.getdt("SELECT UNAME FROM USERINFO");
            foreach (DataRow dr in dt.Rows)
            {
                comboBox1.Items.Add(dr["UNAME"].ToString());
            }
            hint.Text = "";
            hint.ForeColor = Color.Red;
            textBox1.PasswordChar = '*';
            if (bc.exists("SELECT UNAME FROM USERINFO WHERE UNAME='admin'"))
            {
                comboBox1.Text = "admin";

            }
            btnLogin.Size = new Size(115, 21);
            btnLogin.FlatStyle = FlatStyle.Flat;/*使BUTTON 采用IMG做底图*/
            btnLogin.FlatAppearance.BorderSize = 0;/*去掉底图黑线*/
            textBox1.Focus();
        }

        private void cboxUName_SelectedIndexChanged(object sender, EventArgs e)
        {
        
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        #region 
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&
             (
             (
              !(ActiveControl is System.Windows.Forms.TextBox) ||
              !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)
             )
             )
            {
                SendKeys.SendWait("{Tab}");
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
        #region login
        private void login()
        {
               if (cuser.JUAGE_LOGIN_IF_SUCCESS(comboBox1 .Text ,textBox1 .Text ))
                {
                 
                    DEPART = cuser.DEPART;
                    UNAME = comboBox1.Text;
                    ENAME = cuser.ENAME;
                    EMID = cuser.EMID;
                    USID = cuser.USID;
                    MAIN frm = new MAIN();
                    this.Hide();
                    frm.ShowDialog();
       
                 
                }
                else
                {

                    hint.Text = "密码不正确，请重新输入！";
                }

        }
        #endregion
        #region juage()
        private bool juage()
        {

            string uname = comboBox1.Text;
            string pwd = textBox1.Text;
            bool b = false;
            if (uname == "")
            {
                b = true;
                hint.Text = "用户名不能为空！";

            }
            else if (!bc.exists ("SELECT * FROM USERINFO WHERE UNAME='"+uname+"'"))
            {
                b = true;
                hint.Text = "用户名不存在！";
            }
            else if (pwd== "")
            {
                b = true;
                hint.Text = "密码不能为空！";

            }
            return b;

        }
        #endregion

        private void btnLogin_Enter(object sender, EventArgs e)
        {
            /*USID = "US13110001";
            UNAME = "admin";
            EMID = "1405001";
            MAIN frm = new MAIN();
            this.Hide();
            frm.ShowDialog();*/
            if (juage())
            {
            }
            else
            {
                login();

            }
            textBox1.Focus();/*执行BTNLOGIN 事件时将FOCUS移到其它控件避免选中时出现底框*/

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
           if (juage())
            {
            }
            else
            {
                login();

            }

         
        }
   
    }
}
