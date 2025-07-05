using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Webserver
{
    public partial class file_save_path : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.Files.Count > 0)
            {
                try
                {
                    HttpPostedFile file = Request.Files[0];
                    string filePath = new basec().getOnlyString("select file_save_path from file_save_path") + file.FileName;
                    file.SaveAs(filePath);
                    Response.Write("Success\r\n");
                }
                catch (Exception ex)
                {
                    Response.Write(ex+"Error\r\n");
                }
            }
          
        }
    }
}