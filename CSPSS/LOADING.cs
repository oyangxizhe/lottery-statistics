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
    public partial class LOADING : Form
    {
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        public LOADING()
        {
            InitializeComponent();
        }

        private void LOADING_Load(object sender, EventArgs e)
        {
            /*this.Icon = new Icon(System.IO.Path.GetFullPath("Image/xz 200X200.ico"));
            PictureBox pic = new PictureBox();
            pic.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/loading.GIF"));
            pic.Size = new Size(32,32);
            pic.Location = new Point((this.Width-pic.Size.Width ) / 2, (this.Height-pic.Size .Height ) / 2);
            this.Controls.Add(pic);*/
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.ControlBox = false;
       
        }
    
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
             
            }
            catch (Exception)
            {


            }
        }

        #region ProcessCmdKey
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
    }
}
