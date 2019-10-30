using China_System.Common;
using SDZdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QCAuto
{
    public partial class 辣皇后Login : Form
    {
        protected  string ZFCEPath = "";
        protected string netuser = "";
        protected string netpassword = "";
        public string pass;
        public List<lhh_LoginList_info> InfoList;
        public 辣皇后Login(string testvalue)
        {
            InitializeComponent();
            this.Text = String.Format("Login  Version {0}", AssemblyVersion);


            label2.Text = testvalue;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0 && Yhm.Text.Length>0)
            {
                pass = this.textBox1.Text;
                List<lhh_LoginList_info> selectinfo = InfoList.FindAll(f => f.loginid == Yhm.Text.Trim() && f.pwd == pass);
                if (selectinfo.Count <= 0)
                {
                    MessageBox.Show("用户名或密码错误！");
                }
                else 
                {
                    DateTime DateNow =Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd")) ;
                    string starTime = selectinfo[0].startTime;
                    DateTime endTime =Convert.ToDateTime( selectinfo[0].endTime);
                    if (DateNow >= endTime) 
                    {
                        MessageBox.Show("您的使用期限已到！");
                    }else
                    {
                        MessageBox.Show("验证成功！点击确定后为您打开表格。");
                        this.DialogResult = System.Windows.Forms.DialogResult.OK;
                        this.Close();
                    }
                }
            }
            else
            {

                MessageBox.Show("请输入用户名和密码");

            }

        }

        

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }
        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }
        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }
        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
           

        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
           
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
          //  button1_Click(null, EventArgs.Empty);

        }

        private void Login_Load(object sender, EventArgs e)
        {
            InfoList =getlist();
            string[] name = new string[InfoList.Count];
            for (int i = 0; i < InfoList.Count; i++) 
            {
                name[i] = InfoList[i].loginid;
            }
            Yhm.DataSource=name;
        }
        protected List<lhh_LoginList_info> getlist() 
        {
            try
            {
                
                
                InfoList = new List<lhh_LoginList_info>();
                  System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    string[] ob = Regex.Split(ZFCEPath, @"\\", RegexOptions.IgnoreCase);
                    string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "辣皇后\\用户使用时间配置.xlsx");
                 
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, true, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["配置"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.get_Range(WS.Cells[2, 7], WS.Cells[WS.UsedRange.Rows.Count, 10]);
                    int rowCount = WS.UsedRange.Rows.Count - 2;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    clsCommHelp.CloseExcel(excelApp, analyWK);

                    for (int i = 1; i <= rowCount; i++)
                    {
                        //  bgWorker.ReportProgress(0, "读入数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                        lhh_LoginList_info temp = new lhh_LoginList_info();

                        #region 基础信息

                       if(o[i,1]!=null)
                           temp.loginid=o[i,1].ToString().Trim();
                        if(o[i,2]!=null)
                           temp.pwd=o[i,2].ToString().Trim();
                        if(o[i,3]!=null)
                           temp.startTime=o[i,3].ToString().Trim();
                        if (o[i, 4] != null)
                            temp.endTime =Convert.ToDateTime(o[i, 4].ToString());
                    





                        #endregion




                        InfoList.Add(temp);
                }



                return InfoList;
            }
            catch (Exception ex)
            {
                MessageBox.Show("表格存在异常,请参照原始表格格式修改:" + ex.Message);

                throw ex;
            }

        }
        }
    }

