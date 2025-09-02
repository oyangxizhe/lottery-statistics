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
    public partial class MAIN : Form
    {
         DataTable dt = new DataTable();
         DataTable dt2 = new DataTable();
         basec bc = new basec();
         CUSER cuser = new CUSER();
         CEMPLOYEE_INFO cemplyee_info = new CEMPLOYEE_INFO();
         Color c2 = System.Drawing.ColorTranslator.FromHtml("#4a7bb8");
         Color c3 = System.Drawing.ColorTranslator.FromHtml("#24ade5");
         CVOUCHER cvoucher = new CVOUCHER();
         CDEPART cdepart = new CDEPART();
         CPOSITION cposition = new CPOSITION();
         CUSER_GROUP cuser_group = new CUSER_GROUP();
        public MAIN()
        {
            InitializeComponent();
        }
        private void MAIN_Load(object sender, EventArgs e)
        {
            this.Text = "彩票分析系统 Version 1.0.16";
            dt = bc.getdt("SELECT * from RightList where USID = '"+LOGIN .USID +"'");
            SHOW_TREEVIEW(dt);
            menuStrip1.Font = new Font("宋体", 9);
            this.WindowState = FormWindowState.Maximized;
            toolStripStatusLabel1.Text = "||当前用户：" + LOGIN.UNAME;
            toolStripStatusLabel2.Text = "||所属部门：" + LOGIN.DEPART;
            toolStripStatusLabel3.Text = "||登录时间：" + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
            listView1.BackColor = c2;
            listBox3.BackColor = c2;
            listBox3.Height = 84;
            groupBox1.BackColor = c2;
            listView1.ForeColor = Color.White;
            listView1.Font = new Font("新宋体", 11);

            listView1.Location = new Point(1, 75);
            listView2.BorderStyle = BorderStyle.None;
            //listView1.BorderStyle = BorderStyle.None;
            listView2.Location = new Point(195, 75);
            listBox3.Location = new Point(1, -35);
            listView1.Height = 660;
            listView2.Height = 660;
            listView1.Width = 194;
             //listView1 .BackgroundImage  = Image.FromFile(System .IO.Path.GetFullPath("Image/1.png"));
            groupBox1.Height = 675;

            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/1.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/2.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/3.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/4.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/5.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/6.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/7.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/8.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/9.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/10.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/11.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/12.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/13.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/14.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/15.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/16.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/17.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/18.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/19.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/20.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/21.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/22.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/23.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/24.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/25.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/26.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/27.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/28.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/29.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/30.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/31.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/32.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/33.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/34.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/35.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/36.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/37.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/38.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/39.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/40.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/41.png")));
            imageList1.Images.Add(Image.FromFile(System.IO.Path.GetFullPath("Image/42.png")));

            imageList1.ColorDepth = ColorDepth.Depth32Bit;/*防止图片失真*/
            listView1.View = View.SmallIcon;
            listView2.View = View.LargeIcon;
            imageList1.ImageSize = new Size(48, 48);/*set imglist size*/
            listView1.SmallImageList = imageList1;
            listView2.LargeImageList = imageList1;
            
        }
        #region show_treeview
        private void SHOW_TREEVIEW(DataTable dt)
        {
           
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "PARENT_NODEID=0");

            if (dt.Rows.Count > 0)
            {
                for(int i=0;i<dt.Rows.Count ;i++)
                {
                    ListViewItem lvi = listView1.Items.Add(dt.Rows[i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
        
               DataTable  dtx = bc.GET_DT_TO_DV_TO_DT(dt, "", "NODE_NAME='彩票分析'");
                if (dtx.Rows.Count > 0)
                {
                    click(dtx.Rows[0]["NODE_NAME"].ToString());
                    listView1.Items[0].BackColor = c3;
                }
                else
                {

                    click(dt.Rows[0]["NODE_NAME"].ToString());
                    listView1.Items[1].BackColor = c3;
                }
               
                
            }
        }
        #endregion

        #region show_treeview_O
        private void SHOW_TREEVIEW_O(string NODEID)
        {

            dt2 = bc.getdt("SELECT * FROM RIGHTLIST WHERE PARENT_NODEID='" + NODEID  + "'AND  USID = '" + LOGIN.USID + "'");
            if (dt2.Rows.Count > 0)
            {
                for(int i=0;i<dt2.Rows.Count ;i++)
                {
                    ListViewItem lvi = listView2.Items.Add(dt2.Rows [i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt2.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
            }
        }
        #endregion

         private void 退出系统ToolStripMenuItem1_Click(object sender, EventArgs e)
         {
             if (MessageBox.Show("确定要退出本系统吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
             {
                 Application.Exit();
             }
             else
             {
                 MAIN fmain = new MAIN();
                 fmain.Show();
             }
         }
         private void listView1_Click(object sender, EventArgs e)
         {
            
             string v1 = listView1.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/
             click(v1);
            
         }
         private void click(string NODE_NAME)
         {
             listView2.Items.Clear();
             string id = bc.getOnlyString("SELECT NODEID FROM RIGHTLIST WHERE NODE_NAME='" + NODE_NAME + "'");
             SHOW_TREEVIEW_O(id);

             foreach (ListViewItem lvi in listView1.Items)
             {
                 if (lvi.Selected)
                 {
                     lvi.BackColor = c3;
                     pictureBox1.Focus();/*SELECTED AFTER MOVE FOCUS*/
                 }
                 else
                 {
                     lvi.BackColor = c2;
                 }

             }

         }
         private void listView2_Click(object sender, EventArgs e)
         {
             string v1 = listView2.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/

             #region v1
            if (v1 == "双色球分析")
             {
                 CSPSS.BASE_INFO.DOUBLE_BALL FRM = new CSPSS.BASE_INFO.DOUBLE_BALL();

                 FRM.ShowDialog();

             }
         
             else if (v1 == "用户帐户")
             {
                 CSPSS.USER_MANAGE.USER_INFO FRM = new CSPSS.USER_MANAGE.USER_INFO();
                 FRM.IDO = cuser.GETID();
                 FRM.ADD_OR_UPDATE = "ADD";
                 FRM.ShowDialog();

             }
             else if (v1 == "更改密码")
             {
                 CSPSS.USER_MANAGE.EDIT_PWD FRM = new CSPSS.USER_MANAGE.EDIT_PWD();
                 FRM.ShowDialog();

             }
             else if (v1 == "权限管理")
             {
                 CSPSS.USER_MANAGE.EDIT_RIGHT FRM = new CSPSS.USER_MANAGE.EDIT_RIGHT();
                 FRM.ShowDialog();

             }
             else if (v1 == "用户组信息")
             {
                 CSPSS.USER_MANAGE.USER_GROUP FRM = new CSPSS.USER_MANAGE.USER_GROUP();
                 FRM.IDO = cuser_group.GETID();
                 FRM.ShowDialog();

             }
             #endregion
         }
    }
}
