using clsBuiness;
using SDZdb;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Web
{
    public partial class login : System.Web.UI.Page
    {
        public string alterinfo1;

        bool is_AdminIS = false;
        int logis = 0;
        public string user;
        public string pass;

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

                List<string> itemi = new List<string>();

                var myCol = System.Configuration.ConfigurationManager.AppSettings;
                for (int i = 0; i < myCol.Count; i++)
                {
                    //itemi.Add(myCol.Get(i));
                    itemi.Add(myCol.AllKeys[i]);

                }



            }
        }


        protected void HtmlBtn_Click(object sender, EventArgs e)
        {


            string username = Request.Form["username"];
            string txtSAPPassword = Request.Form["password"];

            user = username;
            pass = txtSAPPassword;

            NewMethoduserFind(username.Trim(), txtSAPPassword.Trim());


            //  
        }
        private bool NewMethoduserFind(string user, string pass)
        {

            try
            {
                bool isadmin = false;

                if (pass.Length > 0 && "123" == pass.Trim() && user.Length > 0 && "admin" == user.Trim())
                {
                    isadmin = true;

                    alterinfo1 = "登录成功";
                    HttpCookie cookie = new HttpCookie("MyCook");//初使化并设置Cookie的名称

                    cookie.Values.Set("isadmin", HttpUtility.UrlEncode(isadmin.ToString()));
                    cookie.Expires = System.DateTime.Now.AddYears(100);

                    Response.SetCookie(cookie);
                    HttpCookie cookie1 = Request.Cookies["MyCook"];


                    logis++;
                }
                else if (pass.Length > 0 && "123456" == pass.Trim() && user.Length > 0 && "user" == user.Trim())
                {
                    isadmin = false;

                    alterinfo1 = "登录成功";
                    HttpCookie cookie = new HttpCookie("MyCook");//初使化并设置Cookie的名称

                    cookie.Values.Set("isadmin", HttpUtility.UrlEncode(isadmin.ToString()));
                    cookie.Expires = System.DateTime.Now.AddYears(100);

                    Response.SetCookie(cookie);
                    HttpCookie cookie1 = Request.Cookies["MyCook"];


                    logis++;

                }
                if (logis == 0)
                {
                    pass = "";

                    alterinfo1 = "登录失败，请确认用户名和密码或联系系统管理员，谢谢";
                    return false;
                }
                else
                    Response.Redirect("~/frmPrice.aspx");
                return false;


            }
            catch (Exception ex)
            {
                //ProcessLogger.Fatal("0793212:System Login Start " + DateTime.Now.ToString());
                //MessageBox.Show("登录失败，验证用户信息异常！" + ex, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false; ;

                throw;
            }

        }

        protected void HtmlBtcreate_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/frmUserManger.aspx");
        }

        protected void Btchangepas_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Myadmin/changepassword.aspx");
        }

        protected void Btmain_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/frmReadIDCare.aspx");

        }

        protected void HtmlNOlogin_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/frmReadIDCare.aspx?dengluleibie=nologin");


        }


    }
}