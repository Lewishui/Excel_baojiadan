using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QCAuto
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string pass = "";

            var form = new Login();

            if (form.ShowDialog() == DialogResult.OK)
            {
                pass = form.pass;


            }
            else
                Application.Exit();
            if (pass == null || pass == "")
                return;

            bat_dsoframer();

            #region Noway
            DateTime oldDate = DateTime.Now;
            DateTime dt3;
            string endday = DateTime.Now.ToString("yyyy/MM/dd");
            dt3 = Convert.ToDateTime(endday);
            DateTime dt2;
            dt2 = Convert.ToDateTime("2019/4/20");

            TimeSpan ts = dt2 - dt3;
            int timeTotal = ts.Days;
            if (timeTotal < 0)
            {
                MessageBox.Show("缺失系统文件，或电脑系统更新导致，请联系开发人员 !", "系统错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            #endregion


            //Application.Run(new frmPrice(pass));//报价单
            Application.Run(new frmJiaqizhuantongjibiao(pass));//加气砖统计表
        }
        private static void bat_dsoframer()
        {
            string c = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dsoframer.ocx");
            string destFile = @"C:\Windows\SysWOW64" + "\\dsoframer.ocx";
           // destFile = @"C:\Windows" + "\\dsoframer.ocx";

            int io = 0;

            if (File.Exists(destFile))
            {

            }
            else
            {

                File.Copy(c, destFile, true);//覆盖模式
                io = 1;
            }
            destFile = @"C:\windows\system32" + "\\dsoframer.ocx";


            if ( File.Exists(destFile))
            {

            }
            else
            {
                File.Copy(c, destFile, true);//覆盖模式
                io = 1;
            }

            //此方法不弹窗会静默执行
           if (io == 1)
                bat();
        }
        public static void bat()
        {
            try
            {
                string c = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dos64.bat");

                if (File.Exists(c))
                {
                    //System.Diagnostics.Process.Start(folderpath + "\\saptis.exe");

                    System.Diagnostics.Process p = new System.Diagnostics.Process();
                    p.StartInfo.WorkingDirectory = c;
                    p.StartInfo.UseShellExecute = true;
                    p.StartInfo.FileName = c;
                    p.Start();
                    p.WaitForExit();
                }
                c = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dos32.bat");


                if (File.Exists(c))
                {

                    System.Diagnostics.Process p = new System.Diagnostics.Process();
                    p.StartInfo.WorkingDirectory = c;
                    p.StartInfo.UseShellExecute = true;
                    p.StartInfo.FileName = c;
                    p.Start();
                    p.WaitForExit();
                }
                c = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qaz.bat");
                string cnew = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qaz.reg");


                if (File.Exists(c))
                {
                   // System.Diagnostics.Process.Start("regedit.exe", " /s " + cnew);

                    //System.Diagnostics.Process p = new System.Diagnostics.Process();
                    //p.StartInfo.WorkingDirectory = c;
                    //p.StartInfo.UseShellExecute = true;
                    //p.StartInfo.FileName = c;
                    //p.Start();
                    //p.WaitForExit();
                    string cmd = "reg import " + cnew;
                    string output = "";

                   RunCmd(cmd, out output);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("EX:数据库配置失败 ：" + ex);


                throw;
            }
        }
        public static void RunCmd(string cmd, out string output)
        {
            try
            {
                string CmdPath = @"C:\Windows\System32\cmd.exe";
                cmd = cmd.Trim().TrimEnd('&') + "&exit";//说明：不管命令是否成功均执行exit命令，否则当调用ReadToEnd()方法时，会处于假死状态
                using (Process p = new Process())
                {
                    p.StartInfo.FileName = CmdPath;
                    p.StartInfo.UseShellExecute = false;        //是否使用操作系统shell启动
                    p.StartInfo.RedirectStandardInput = true;   //接受来自调用程序的输入信息
                    p.StartInfo.RedirectStandardOutput = true;  //由调用程序获取输出信息
                    p.StartInfo.RedirectStandardError = true;   //重定向标准错误输出
                    p.StartInfo.CreateNoWindow = true;          //不显示程序窗口
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();//启动程序

                    //向cmd窗口写入命令
                    p.StandardInput.WriteLine(cmd);
                    p.StandardInput.AutoFlush = true;

                    //获取cmd窗口的输出信息
                    output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();//等待程序执行完退出进程
                    p.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("EX:数据库配置失败 ：" + ex);


                throw;
            }
        }
      

    }
}
