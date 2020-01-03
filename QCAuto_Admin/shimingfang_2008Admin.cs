//using China_System.Common;
//using SDZdb;
using System;
using System.Collections;
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

namespace QCAuto_Admin
{
    public partial class shimingfang_2008Admin : Form
    {
        protected string ZFCEPath = "";
        protected string netuser = "";
        protected string netpassword = "";
        public string pass;

        bool islocalpath = false;

        public shimingfang_2008Admin()
        {
            InitializeComponent();
            this.Text = String.Format("Login  Version {0}", AssemblyVersion);


            //label2.Text = testvalue;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "shimingfang_2008");
            FilePath = FilePath.Replace("QCAuto_Admin", "QCAuto");
            string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\shimingfang_2008";
            if (Directory.Exists(strDesktopPath))
            {
                List<string> Alist = new List<string>();
                if (Directory.Exists(FilePath))
                {

                    Alist = GetBy_CategoryReportFileName(FilePath);
                    for (int i = 0; i < Alist.Count; i++)
                    {
                        FileInfo info = new FileInfo(Alist[i]);
                        if (info.Exists)
                        {
                            info.Attributes = FileAttributes.Normal;
                        }
                    }
                }
                CopyFolder(strDesktopPath, FilePath);
                #region 加密
                string s = DESEncrypt.Encrypt(System.IO.File.ReadAllText(FilePath + "\\user control.txt", Encoding.Default));
                // tb_Mi.Text = DESEncrypt.Encrypt(tb_Min.Text);
                // System.IO.StreamWriter sw = new System.IO.StreamWriter("myfile.txt", true);
                //StreamWrite sw = new StreamWrite("myfile.txt",true);
                string mi = FilePath + "\\user control.txt";
                System.IO.File.WriteAllText(mi, s, Encoding.Default);
                s = DESEncrypt.Encrypt(System.IO.File.ReadAllText(FilePath + "\\ip.txt", Encoding.Default));

                mi = FilePath + "\\ip.txt";
                System.IO.File.WriteAllText(mi, s, Encoding.Default);

                #endregion


                Alist = GetBy_CategoryReportFileName(FilePath);

                for (int i = 0; i < Alist.Count; i++)
                {
                    FileInfo info = new FileInfo(Alist[i]);
                    if (info.Exists)
                    {
                        info.Attributes = FileAttributes.Hidden;
                    }
                }
                MessageBox.Show("导入完成可以去客户机上激活！");
            }
            else
                MessageBox.Show("导入失败将配置文件按照使用说明放在电脑桌面路径下！");




        }




        public static void CopyFolder(string sourcePath, string destPath)
        {
            if (Directory.Exists(sourcePath))
            {
                if (!Directory.Exists(destPath))
                {
                    //目标目录不存在则创建
                    try
                    {
                        Directory.CreateDirectory(destPath);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("创建目标目录失败：" + ex.Message);
                    }
                }
                //获得源文件下所有文件
                List<string> files = new List<string>(Directory.GetFiles(sourcePath));
                files.ForEach(c =>
                {
                    string destFile = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                    File.Copy(c, destFile, true);//覆盖模式
                });
                //获得源文件下所有目录文件
                List<string> folders = new List<string>(Directory.GetDirectories(sourcePath));
                folders.ForEach(c =>
                {
                    string destDir = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                    //采用递归的方法实现
                    CopyFolder(c, destDir);
                });
            }
            else
            {
                throw new DirectoryNotFoundException("源目录不存在！");
            }
        }

        public List<string> GetBy_CategoryReportFileName(string dirPath)
        {

            List<string> FileNameList = new List<string>();
            ArrayList list = new ArrayList();

            if (Directory.Exists(dirPath))
            {
                list.AddRange(Directory.GetFiles(dirPath));
            }
            if (list.Count > 0)
            {
                foreach (object item in list)
                {
                    if (!item.ToString().Contains("~$"))
                        //FileNameList.Add(item.ToString().Replace(dirPath + "\\", ""));
                        FileNameList.Add(item.ToString());
                }
            }

            return FileNameList;
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

        }

        private void Local_IP()
        {



            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "shimingfang_2008\\ip.txt";
            string[] fileText = File.ReadAllLines(A_Path);
            if (fileText.Length > 0 && fileText[0] != null && fileText[0] != "")
            {
                if (fileText[0] != null && fileText[0] != "")
                {
                    ZFCEPath = fileText[3];
                    if (!ZFCEPath.Contains("\\"))
                    {
                        ZFCEPath = "" + AppDomain.CurrentDomain.BaseDirectory + "shimingfang_2008\\" + ZFCEPath;
                        islocalpath = true;

                    }
                }
                if (fileText.Length > 1 && fileText[1] != null && fileText[1] != "")
                    netuser = fileText[1];
                if (fileText.Length > 2 && fileText[2] != null && fileText[2] != "")
                    netpassword = fileText[2];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "shimingfang_2008");
            FilePath = FilePath.Replace("QCAuto_Admin", "QCAuto");

            string strDesktopPath = @"C:\Program Files\shimingfang_2008";


            CopyFolder(FilePath, strDesktopPath);
            //DelectDir(FilePath);

            MessageBox.Show("激活完成，请重启客户端登录使用！");
        }

        public static void DelectDir(string srcPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        //如果 使用了 streamreader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                        File.Delete(i.FullName);      //删除指定文件
                    }
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }
    }


}

