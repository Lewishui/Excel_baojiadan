
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ThoughtWorks.QRCode.Codec;
using ZXing;

namespace Web
{
    public partial class _Default : System.Web.UI.Page
    {
        private string path;
        Bitmap scanBitmap;

        protected void Page_Load(object sender, EventArgs e)
        {
          
     
            //path = @"D:\Devlop\Test\C#二维码生成\二维码识别\jiashizheng.png";

            if (!Page.IsPostBack)
            {
                //Response.Redirect("frmPrice.aspx");
                Response.Redirect("~/Myadmin/login.aspx");
            }
        }

        protected void btnDelete_Click(object sender, EventArgs e)
        {




        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            Bitmap bt = new Bitmap(@"D:\Devlop\Test\C#二维码生成\二维码识别\jiashizheng.png");

            BarcodeReader reader = new BarcodeReader();
            reader.Options.CharacterSet = "UTF-8";
            reader.Options.TryHarder = true;
            reader.Options.PureBarcode = true;
            var result = reader.Decode(bt);
            //return (result == null) ? null : result.Text;
            if (result != null)
                this.txbModelName.Text = result.Text;
        }

        protected void btnCopyPreDocVersion_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/WebForm1.aspx");

        }


    }
}