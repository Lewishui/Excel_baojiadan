using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Web.ErrorPage
{
    public partial class ErrorPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                string QiHao = Request.QueryString["error"];
                if (Application["error"] != null)
                    ErrorMessageLabel.Text = Application["error"].ToString();
                else if (Request.QueryString["error"] != null && Request.QueryString["error"] != "")
                    ErrorMessageLabel.Text = Request.QueryString["error"];

            }
        }
    }
}