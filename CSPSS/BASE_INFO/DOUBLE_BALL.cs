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
using OfficeOpenXml;
using System.IO;

namespace CSPSS.BASE_INFO
{
    public partial class DOUBLE_BALL : Form
    {
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dtx = new DataTable();
        basec bc = new basec();
        private bool _IF_COMPLETED;
        public bool IF_COMPLETED
        {
            set { _IF_COMPLETED = value; }
            get { return _IF_COMPLETED; }
        }
        private string _PERIOD;
        public string PERIOD
        {
            set { _PERIOD = value; }
            get { return _PERIOD; }

        }
        private string _DBID;
        public string DBID
        {
            set { _DBID = value; }
            get { return _DBID; }
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
        ExcelToCSHARP etc = new ExcelToCSHARP();
        protected int M_int_judge, i, look;
        protected int getdata;
       
        Color c2 = System.Drawing.ColorTranslator.FromHtml("#990033");
        CDOUBLE_BALL cdouble_ball = new CDOUBLE_BALL();
        LOADING loading = new LOADING();
        public DOUBLE_BALL()
        {
            InitializeComponent();
        }
        #region init
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DOUBLE_BALL));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.chk2 = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
            this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.btnToExcel = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnToCSharp = new System.Windows.Forms.Button();
            this.hint = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.PictureBox();
            this.btnExit = new System.Windows.Forms.PictureBox();
            this.btnSearch = new System.Windows.Forms.PictureBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.Location = new System.Drawing.Point(6, 20);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(545, 225);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DataSourceChanged += new System.EventHandler(this.dataGridView1_DataSourceChanged);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.chk2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dtpEndDate);
            this.groupBox1.Controls.Add(this.dtpStartDate);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnToExcel);
            this.groupBox1.Location = new System.Drawing.Point(3, 127);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(936, 112);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "导出信息";
            // 
            // button2
            // 
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("宋体", 9F);
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button2.Location = new System.Drawing.Point(730, 20);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(50, 64);
            this.button2.TabIndex = 53;
            this.button2.Text = "导出出现次数";
            this.button2.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("宋体", 9F);
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button1.Location = new System.Drawing.Point(803, 22);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(50, 64);
            this.button1.TabIndex = 52;
            this.button1.Text = "导出前两位号";
            this.button1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // chk2
            // 
            this.chk2.AutoSize = true;
            this.chk2.Location = new System.Drawing.Point(152, 31);
            this.chk2.Name = "chk2";
            this.chk2.Size = new System.Drawing.Size(15, 14);
            this.chk2.TabIndex = 51;
            this.chk2.UseVisualStyleBackColor = true;
            this.chk2.CheckedChanged += new System.EventHandler(this.chk2_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(407, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(11, 12);
            this.label2.TabIndex = 49;
            this.label2.Text = "~";
            // 
            // dtpEndDate
            // 
            this.dtpEndDate.CustomFormat = "yyyy/MM/dd";
            this.dtpEndDate.Enabled = false;
            this.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpEndDate.Location = new System.Drawing.Point(471, 22);
            this.dtpEndDate.Name = "dtpEndDate";
            this.dtpEndDate.Size = new System.Drawing.Size(121, 21);
            this.dtpEndDate.TabIndex = 48;
            // 
            // dtpStartDate
            // 
            this.dtpStartDate.CustomFormat = "yyyy/MM/dd";
            this.dtpStartDate.Enabled = false;
            this.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpStartDate.Location = new System.Drawing.Point(232, 22);
            this.dtpStartDate.Name = "dtpStartDate";
            this.dtpStartDate.Size = new System.Drawing.Size(121, 21);
            this.dtpStartDate.TabIndex = 47;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(173, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 46;
            this.label1.Text = "期数期间";
            // 
            // btnToExcel
            // 
            this.btnToExcel.FlatAppearance.BorderSize = 0;
            this.btnToExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToExcel.Font = new System.Drawing.Font("宋体", 9F);
            this.btnToExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnToExcel.Image")));
            this.btnToExcel.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnToExcel.Location = new System.Drawing.Point(659, 20);
            this.btnToExcel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnToExcel.Name = "btnToExcel";
            this.btnToExcel.Size = new System.Drawing.Size(50, 64);
            this.btnToExcel.TabIndex = 11;
            this.btnToExcel.Text = "导出中奖号码";
            this.btnToExcel.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnToExcel.UseVisualStyleBackColor = false;
            this.btnToExcel.Click += new System.EventHandler(this.btnToExcel_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(857, 95);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(29, 12);
            this.label11.TabIndex = 29;
            this.label11.Text = "退出";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(771, 95);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(29, 12);
            this.label12.TabIndex = 28;
            this.label12.Text = "搜索";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.btnToCSharp);
            this.groupBox2.Controls.Add(this.hint);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Controls.Add(this.btnAdd);
            this.groupBox2.Controls.Add(this.btnExit);
            this.groupBox2.Controls.Add(this.btnSearch);
            this.groupBox2.Location = new System.Drawing.Point(3, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(936, 121);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "菜单栏";
            // 
            // btnToCSharp
            // 
            this.btnToCSharp.FlatAppearance.BorderSize = 0;
            this.btnToCSharp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToCSharp.Font = new System.Drawing.Font("宋体", 9F);
            this.btnToCSharp.Image = ((System.Drawing.Image)(resources.GetObject("btnToCSharp.Image")));
            this.btnToCSharp.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnToCSharp.Location = new System.Drawing.Point(443, 28);
            this.btnToCSharp.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnToCSharp.Name = "btnToCSharp";
            this.btnToCSharp.Size = new System.Drawing.Size(50, 64);
            this.btnToCSharp.TabIndex = 403;
            this.btnToCSharp.Text = "导入";
            this.btnToCSharp.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnToCSharp.UseVisualStyleBackColor = false;
            this.btnToCSharp.Click += new System.EventHandler(this.btnToCSharp_Click);
            // 
            // hint
            // 
            this.hint.AutoSize = true;
            this.hint.Location = new System.Drawing.Point(407, 95);
            this.hint.Name = "hint";
            this.hint.Size = new System.Drawing.Size(29, 12);
            this.hint.TabIndex = 402;
            this.hint.Text = "hint";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(341, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 401;
            this.label4.Text = "label4";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(28, 95);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(29, 12);
            this.label17.TabIndex = 24;
            this.label17.Text = "新增";
            // 
            // btnAdd
            // 
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.InitialImage = null;
            this.btnAdd.Location = new System.Drawing.Point(12, 20);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(60, 60);
            this.btnAdd.TabIndex = 16;
            this.btnAdd.TabStop = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnExit
            // 
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.InitialImage = null;
            this.btnExit.Location = new System.Drawing.Point(843, 20);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(60, 60);
            this.btnExit.TabIndex = 19;
            this.btnExit.TabStop = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Image = ((System.Drawing.Image)(resources.GetObject("btnSearch.Image")));
            this.btnSearch.InitialImage = null;
            this.btnSearch.Location = new System.Drawing.Point(757, 20);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(60, 60);
            this.btnSearch.TabIndex = 18;
            this.btnSearch.TabStop = false;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dataGridView1);
            this.groupBox3.Location = new System.Drawing.Point(3, 245);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(557, 251);
            this.groupBox3.TabIndex = 35;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "各期中奖号码";
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.dataGridView2);
            this.groupBox4.Location = new System.Drawing.Point(566, 245);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(373, 251);
            this.groupBox4.TabIndex = 36;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "中奖号码出现次数";
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView2.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView2.Location = new System.Drawing.Point(6, 20);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 23;
            this.dataGridView2.Size = new System.Drawing.Size(361, 225);
            this.dataGridView2.TabIndex = 1;
            // 
            // groupBox5
            // 
            this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox5.Controls.Add(this.dataGridView3);
            this.groupBox5.Location = new System.Drawing.Point(3, 496);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(936, 116);
            this.groupBox5.TabIndex = 37;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "出现次数排前两位号";
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToAddRows = false;
            this.dataGridView3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView3.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView3.Location = new System.Drawing.Point(6, 16);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.RowTemplate.Height = 23;
            this.dataGridView3.Size = new System.Drawing.Size(924, 92);
            this.dataGridView3.TabIndex = 2;
            // 
            // DOUBLE_BALL
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(942, 616);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "DOUBLE_BALL";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "双色球出球次数统计";
            this.Load += new System.EventHandler(this.DOUBLE_BALL_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.groupBox5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion

        private void dowork(object sender, DoWorkEventArgs e)
        {
           
     
            for (i = 0; i < 600; i++)
            {
                System.Threading.Thread.Sleep(100);
                if (IF_COMPLETED)
                {
                    groupBox1.Visible = true;
                    groupBox2.Visible = true;
                    groupBox3.Visible = true;
                    groupBox4.Visible = true;
                    groupBox5.Visible = true;
                    loading.Close();
                    break;
                }

            }

        }
        private void DOUBLE_BALL_Load(object sender, EventArgs e)
        {
            //string a = "07,09,11,15,18,25|07";
            //MessageBox.Show(a.Length.ToString ());
            Control.CheckForIllegalCrossThreadCalls = false;//避免出现线程间操作无效: 从不是创建控件“progressBar1”的线程访问它 160120
            //loading.Show();
            IF_COMPLETED = false;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            groupBox5.Visible = false;
            BackgroundWorker backgroundwork = new BackgroundWorker();
            backgroundwork.RunWorkerAsync();
            backgroundwork.DoWork += new DoWorkEventHandler(dowork);
            label1.Visible = false;
            label2.Visible = false;
            dtpStartDate.Visible = false;
            dtpEndDate.Visible = false;
            label4.Visible = false;
            chk2.Visible = false;
            label17.Visible = false;
            label12.Visible = false;
            btnAdd.Visible = false;
            btnSearch.Visible = false;

            btnToCSharp.Visible = true;
       
           /* DateTime date1 = Convert.ToDateTime("2017-12-21 23:59:59");
            DateTime date2 = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            
            if (date2 > date1)
            {
                //hint.Text = "软件使用已经到期，如需要继续使用请联系供应商再次授权获取更新包，谢谢！";
                MessageBox.Show("软件使用已经到期，如需要继续使用请联系供应商再次授权获取更新包，谢谢！");
                this.Close();
            }
            else
            {
               
          
            }*/
            Bind();
            IF_COMPLETED = true;


         
        }
        #region Bind
        public void Bind()
        {
            string s = bc.GET_SQLCONNECTION_STRING();
            this.WindowState = FormWindowState.Maximized;
            dt = bc.getdt(etc.sql+" ORDER BY PERIOD DESC");
            dataGridView1.DataSource = dt;
            dataGridView2.DataSource = cdouble_ball.GetTableInfo();
            dataGridView3.DataSource = cdouble_ball.GetTableInfo_t();
            dgvStateControl();
            hint.Text = "";
            hint.ForeColor = Color.Red;     
        }
        #endregion
    
        #region dgvStateControl
        private void dgvStateControl()
        {
            dataGridView1.Columns["期数"].Width = 60;
            dataGridView1.Columns["红球一"].Width = 60;
            dataGridView1.Columns["红球二"].Width = 60;
            dataGridView1.Columns["红球三"].Width = 60;
            dataGridView1.Columns["红球四"].Width = 60;
            dataGridView1.Columns["红球五"].Width = 60;
            dataGridView1.Columns["红球六"].Width = 60;
            dataGridView1.Columns["蓝球"].Width = 60;

            dataGridView2.Columns["红球一"].Width = 60;
            dataGridView2.Columns["红球二"].Width = 60;
            dataGridView2.Columns["红球三"].Width = 60;
            dataGridView2.Columns["红球四"].Width = 60;
            dataGridView2.Columns["红球五"].Width = 60;
            dataGridView2.Columns["红球六"].Width = 60;
            dataGridView2.Columns["蓝球"].Width = 60;

            dataGridView3.Columns["序号"].Width = 60;
            dataGridView3.Columns["红球一"].Width = 60;
            dataGridView3.Columns["红球二"].Width = 60;
            dataGridView3.Columns["红球三"].Width = 60;
            dataGridView3.Columns["红球四"].Width = 60;
            dataGridView3.Columns["红球五"].Width = 60;
            dataGridView3.Columns["红球六"].Width = 60;
            dataGridView3.Columns["蓝球"].Width = 60;
            /*dataGridView1.Columns["单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["支出金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["收入金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["公司余额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;*/
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
    

                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }

            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i == 6)
                {
                    dataGridView1.Columns[i].ReadOnly = true;
                }
                else
                {
                    dataGridView1.Columns[i].ReadOnly = true;
                }
                if (i == 0)
                {
                    dataGridView1.Columns[i].Visible = true;

                }
            }

            dataGridView2.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols2 = dataGridView2.Columns.Count;
            for (i = 0; i < numCols2; i++)
            {

                dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
    

                dataGridView2.EnableHeadersVisualStyles = false;
                dataGridView2.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView2.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }

            for (i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i == 6)
                {
                    dataGridView2.Columns[i].ReadOnly = true;
                }
                else
                {
                    dataGridView2.Columns[i].ReadOnly = true;
                }
                if (i == 0)
                {
                    dataGridView2.Columns[i].Visible = true;

                }
                if (dataGridView2.Columns[i].Name == "序号")
                {
                    dataGridView2.Columns[i].Visible = false;
                }
            }
            dataGridView3.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            numCols2 = dataGridView3.Columns.Count;
            for (i = 0; i < numCols2; i++)
            {

                dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
    

                dataGridView3.EnableHeadersVisualStyles = false;
                dataGridView3.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView3.Columns.Count; i++)
            {
                dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView3.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }

            for (i = 0; i < dataGridView3.Columns.Count; i++)
            {
                dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i == 6)
                {
                    dataGridView3.Columns[i].ReadOnly = true;
                }
                else
                {
                    dataGridView3.Columns[i].ReadOnly = true;
                }
                if (i == 0)
                {
                    dataGridView3.Columns[i].Visible = true;

                }
            }
        
        }
        #endregion

        #region add

        #endregion
  


        #region override enter
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


        #region doubleclick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

      
        }
        #endregion
        public void a2()
        {

            getdata = 1;

        }
        public void a3()
        {

            getdata = 2;

        }
   
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
               
                bc.dgvtoExcel_for_epplus(dataGridView1, "彩票各期中奖号码");
                
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.0000";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string a1 = bc.numYMD(12, 4, "0001", "select * from DOUBLE_BALL_MST", "VOID", "VO");
            if (a1 == "Exceed Limited")
            {
                MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
          
            try
            {
                search();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #region search()
        public void search()
        {
            string v1 = dtpStartDate.Value.ToString("yyyy/MM/dd 00:00:00").Replace("-", "/");
            string v2 = dtpEndDate.Value.ToString("yyyy/MM/dd 23:59:59").Replace("-", "/");

            string v3 = dtpStartDate.Value.ToString("yyyy/MM/dd").Replace("-", "/");
            string v4 = dtpEndDate.Value.ToString("yyyy/MM/dd").Replace("-", "/");
            dt = basec.getdts( " ORDER BY C.ACNAME ASC");



            if (chk2.Checked)
            {



            }

        

    
        }
        #endregion

        #region search_o()
        public void search_o(string sql)
        {
            string sqlo =" ORDER BY A.VOID ASC";
            string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='"+LOGIN .USID +"'");
            if (v7 == "Y")
            {
             
                dt = bc.getdt(sql+sqlo);
                
            }
            else if (v7 == "GROUP")
            {

                dt = bc.getdt(sql + @" AND B.MAKERID IN (SELECT EMID FROM USERINFO A WHERE USER_GROUP IN 
 (SELECT USER_GROUP FROM USERINFO WHERE USID='"+LOGIN .USID +"'))"+sqlo );
            }
            else
            {
                dt = bc.getdt(sql + " AND B.MAKERID='"+LOGIN .EMID +"'"+sqlo );

            }
          
         
           
            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;
                dgvStateControl();
            }
            else
            {


                hint.Text = "找不到所要搜索项！";
                dataGridView1.DataSource = dt;

            }

        }
        #endregion
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            PERIOD period = new PERIOD();
            MessageBox.Show(period.NEXT_FINANCIAL_YEAR + "," + period.NEXT_PERIOD+","+period .NEXT_PERIOD_t );
        }

        private void chk2_CheckedChanged(object sender, EventArgs e)
        {
            if (chk2.Checked)
            {
                dtpStartDate.Enabled = true;
                dtpEndDate.Enabled = true;
            }
            else
            {
                dtpStartDate.Enabled = false;
                dtpEndDate.Enabled = false;

            }
        }

        private void btnToCSharp_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfv = new OpenFileDialog();
            if (opfv.ShowDialog() == DialogResult.OK)
            {
                /*DataSet ds = new DataSet();
                string tablename = ExcelToCSHARP.GetExcelFirstTableName(opfv .FileName );
                ds = ExcelToCSHARP.importExcelToDataSet(opfv .FileName , tablename);
                DataTable dt = ds.Tables[0];
                dataGridView1.DataSource = dt;*/
                string path = opfv.FileName;
                etc.EMID = "123";
                etc.showdata_for_epplus(path);
                Bind();
            }
            try
            {
          
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    

    
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel_for_epplus(dataGridView3, "导出前两位号码");

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel_for_epplus(dataGridView2, "彩票各期中奖次数");

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    

  
    }
}
