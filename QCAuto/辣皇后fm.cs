using China_System.Common;
using clsBuiness;
using ISR_System;
using Microsoft.Office.Interop.Excel;
using SDZdb;
using Spire.Xls;
//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using WEBIBM;
using Excel = Microsoft.Office.Interop.Excel;
namespace QCAuto
{
    public partial class 辣皇后fm : Form
    {
        private AxDSOFramer.AxFramerControl m_axFramerControl = new AxDSOFramer.AxFramerControl();

        private WbBlockNewUrl MyWebBrower;
        Thread thOpen;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;

        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        //private AxSHDocVw.AxWebBrowser axWebBrowser1;
        string ZFCEPath;
        List<cls_sixzhuanjiagebiao_info> MAPPINGResult;
        string FileDownURL = string.Empty;
        string FileName = "";
        private China_System.Common.clsCommHelp.SortableBindingList<cls_sixzhuanjiagebiao_info> sortabledinningsOrderList;
        Excel.Application oApp;
        Excel.Workbooks oBooks;
        Excel.Workbook oBook;
        Excel.Worksheet oSheet;
        object oWebBrowser;
        string pass;
        string netuser;
        string netpassword;
        int axFramerControl1_is = 0;
        private Form viewForm;
        private DateTime strFileName;
        private string publicPDFName;
        private bool isReadyForSearch = false;
        bool blFresh = false;
        #region Import API
        System.Timers.Timer aTimer = new System.Timers.Timer(50);//实例化Timer类，设置间隔时间为10000毫秒； 
        System.Timers.Timer t = new System.Timers.Timer(50);//实例化Timer类，设置间隔时间为10000毫秒； 

        private int ScreenStatus, intCnt;
        private bool RUNING = false;
        private const int WM_KEYDOWN = 0x100;
        private const int WM_KEYUP = 0x101;
        private const int VK_TAB = 0x9;
        private const int VK_CONTROL = 0x11;
        private const int VK_PRIOR = 0x21;
        private const int VK_UP = 0x26;
        private const int VK_HOME = 0x24;
        private const int BM_CLICK = 0xF5;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int SYSKEYDOWN = 0x104;
        private const int WM_SETTEXT = 0x000C;
        private bool WebSiteStatus = false;
        private bool IntialFinish = false;
        private IntPtr hwnd_main, hwnd_ReportTree, hwnd_ReportTree1, hwnd_Control;
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint = "GetParent")]
        public static extern IntPtr GetParent(IntPtr hwndChild);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);


        [DllImport("User32.dll ")]
        public static extern IntPtr GetDlgItem(IntPtr parent, long childe);


        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);
        #endregion


        public 辣皇后fm(string password)
        {
            try
            {
                //bat_dsoframer();

                InitializeComponent();


                pass = password;
                Local_IP();
                int ssd = 0;
                #region test
                tabControl1.TabPages[2].Parent = null;//调用的是 AxDSOFramer  也好用，但是打开保存后共享Excel就变位只读了
               // tabControl1.TabPages[2].Parent = null;//统计表wb 按钮好用
                //toolStripButton5.Visible = false;

                #endregion

                //  ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "全晟新材料\\jxc.xlsx");
                // ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "加气砖统计表\\2019年加气砖统计表.xlsx");
                //ssd = 1;
                if (ssd == 0)
                {
                    string[] ob = Regex.Split(ZFCEPath, @"\\", RegexOptions.IgnoreCase);
                    //bool status = SharedTool.connectState(@"\\192.168.1.2", @"administrator", "333333");
                    string ipadd = "\\\\" + ob[2];
                    if (!ZFCEPath.Contains("D:\\Devlop\\报价单\\ewm\\Excel_baojiadan\\QCAuto\\bin\\Debug"))
                    {
                        bool status = SharedTool.connectState(ipadd, @netuser, netpassword);

                        if (!File.Exists(ZFCEPath) && status != true)
                        {
                            MessageBox.Show("没有找到此路径或此文件，请保证共享文件存在!");
                            System.Environment.Exit(0);
                            return;
                        }
                    }
                }


                #region 好用但是弹窗有结果
                //string cmd = @"regsvr32 C:\Windows\SysWOW64\dsoframer.ocx";
                //string output = "";

                //RunCmd(cmd, out output);
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);

                throw;
            }
        }

        private static void bat_dsoframer()
        {
            string c = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dsoframer.ocx");
            string destFile = @"C:\Windows\SysWOW64" + "\\dsoframer.ocx";

            int io = 0;

            if (Directory.Exists(destFile))
            {
                File.Copy(c, destFile, true);//覆盖模式
                io = 1;
            }
            destFile = @"C:\windows\system32" + "\\dsoframer.ocx";


            if (Directory.Exists(destFile))
            {
                File.Copy(c, destFile, true);//覆盖模式
                io = 1;
            }

            //此方法不弹窗会静默执行
            if (io == 1)
                bat();
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
            }
            catch (Exception ex)
            {
                MessageBox.Show("EX:数据库配置失败 ：" + ex);


                throw;
            }
        }

        private void Local_IP()
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "辣皇后\\ip.txt";
            string[] fileText = File.ReadAllLines(A_Path);
            if (fileText.Length > 0 && fileText[0] != null && fileText[0] != "")
            {
                if (fileText[0] != null && fileText[0] != "")
                    ZFCEPath = fileText[0];
                if (fileText.Length > 1 && fileText[1] != null && fileText[1] != "")
                    netuser = fileText[1];
                if (fileText.Length > 2 && fileText[2] != null && fileText[2] != "")
                    netpassword = fileText[2];
            }
        }
        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            //  Inputexcel();

            Object refmissing = System.Reflection.Missing.Value;
            //this.axWebBrowser2.Navigate(ZFCEPath);
            //axWebBrowser2.Navigate(ZFCEPath, ref refmissing, ref refmissing, ref refmissing, ref refmissing);
            //   axWebBrowser2.ExecWB(SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, ref refmissing, ref refmissing);
            //  this.webBrowser1.Navigate(strFileName);
            //    object axWebBrowser = this.webBrowser1.ActiveXInstance;




        }

        //private void axWebBrowser2_NavigateComplete2(object sender, AxSHDocVw.DWebBrowserEvents2_NavigateComplete2Event e)
        //{

        //    ///   return;

        //    object o = e.pDisp;
        //    oWebBrowser = e.pDisp;
        //    try
        //    {

        //        Object oDocument = o.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, o, null);
        //        Object oApplication = o.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oDocument, null);
        //        Excel.Application eApp = (Excel.Application)oApplication;
        //        eApp.UserControl = true;
        //        //Inputexcel(eApp);
        //        //textexcel();


        //        #region 方法2
        //        //Object refmissing = System.Reflection.Missing.Value;
        //        //object[] args = new object[4];
        //        //args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;
        //        //args[1] = SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER;
        //        //args[2] = refmissing;
        //        //args[3] = refmissing;

        //        //object axWebBrowser = this.webBrowser1.ActiveXInstance;

        //        //axWebBrowser.GetType().InvokeMember("ExecWB",
        //        //    BindingFlags.InvokeMethod, null, axWebBrowser, args);


        //        //object Application = axWebBrowser.GetType().InvokeMember("Document",
        //        //    BindingFlags.GetProperty, null, axWebBrowser, null);

        //        //Excel.Workbook wbb = (Excel.Workbook)oApplication;
        //        //Excel.ApplicationClass excel = wbb.Application as Excel.ApplicationClass;
        //        //Excel.Workbook wb = excel.Workbooks[1];
        //        //Excel.Worksheet ws = wb.Worksheets[1] as Excel.Worksheet;
        //        //ws.Cells.Font.Name = "Verdana";
        //        //ws.Cells.Font.Size = 14;
        //        //ws.Cells.Font.Bold = true;
        //        //Excel.Range range = ws.Cells;

        //        //Excel.Range oCell = range[10, 10] as Excel.Range;
        //        //oCell.Value2 = "你好";
        //        #endregion


        //        #region inster tx
        //        //object objBooks = eApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, eApp, null);

        //        ////添加一个新的Workbook
        //        //object objBook = objBooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, objBooks, null);
        //        ////获取Sheet集
        //        //object objSheets = objBook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, objBook, null);

        //        ////获取第一个Sheet对象
        //        //object[] Parameters = new Object[1] { 1 };
        //        //object objSheet = objSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, objSheets, Parameters);

        //        //Parameters = new Object[2] { 1, 1 + 1 };
        //        //object objCells = objSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, objSheet, Parameters);
        //        ////向指定单元格填写内容值
        //        //Parameters = new Object[1] { "name" };
        //        //objCells.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objCells, Parameters);

        //        #endregion

        //        #region 一、首先简要回顾一下如何操作Excel表
        //        Workbooks workbooks = eApp.Workbooks;
        //        Excel.ApplicationClass excel = workbooks.Application as Excel.ApplicationClass;
        //        Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)workbooks.get_Item(1);
        //        Excel.Workbook wb = excel.Workbooks[1];
        //        //_Workbook workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        //        int c = workbooks.Count;
        //        _Workbook workbook = workbooks.Add(ZFCEPath);
        //        Sheets sheets = workbook.Worksheets;

        //        _Worksheet worksheet = (_Worksheet)sheets.get_Item(1);
        //        Range range1 = worksheet.get_Range("A1", Missing.Value);
        //        const int nCells = 2345;
        //        range1.Value2 = nCells;

        //        #endregion


        //        ExcelExit();

        //    }
        //    catch (Exception ex)
        //    {
        //        ExcelExit();

        //        throw;
        //    }
        //}
        //public void Inputexcel(Microsoft.Office.Interop.Excel.Application excelApp1)
        //{


        //    try
        //    {
        //        string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xls");
        //        //需要换 成日期的导出表
        //        System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
        //        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

        //        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        //        Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(ZFCEPath, Type.Missing, true, Type.Missing,
        //            "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //            Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        //        excelApp.Visible = true;
        //        excelApp.ScreenUpdating = true;

        //        Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
        //        Microsoft.Office.Interop.Excel.Range rng;
        //        //   rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 45]);
        //        rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
        //        int rowCount = WS.UsedRange.Rows.Count - 1;
        //        object[,] o = new object[1, 1];
        //        o = (object[,])rng.Value2;
        //        Microsoft.Office.Interop.Excel.AllowEditRanges ranges = WS.Protection.AllowEditRanges;
        //        ranges.Add("Information", WS.Range["B2:E6"], Type.Missing);

        //        WS.Protect("123", true);

        //        clsCommHelp.CloseExcel(excelApp, analyWK);

        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }

        //}

        public void textexcel()
        {

            Excel.Application app = new Excel.Application();
            if (app == null)
            {
                // statusBar1.Text = "ERROR: EXCEL couldn''t be started!";
                return;
            }

            app.Visible = true; //如果只想用程序控制该excel而不想让用户操作时候，可以设置为false
            app.UserControl = true;

            Workbooks workbooks = app.Workbooks;

            _Workbook workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet); //根据模板产生新的workbook
            // _Workbook workbook = workbooks.Add("c://a.xls"); //或者根据绝对路径打开工作簿文件a.xls

            Sheets sheets = workbook.Worksheets;
            _Worksheet worksheet = (_Worksheet)sheets.get_Item(1);
            if (worksheet == null)
            {
                //  statusBar1.Text = "ERROR: worksheet == null";
                return;
            }

            // This paragraph puts the value 5 to the cell G1
            Range range1 = worksheet.get_Range("A1", Missing.Value);
            if (range1 == null)
            {
                ///    statusBar1.Text = "ERROR: range == null";
                return;
            }
            const int nCells = 2345;
            range1.Value2 = nCells;

        }
        private void ExcelExit()
        {
            //if (this.axFramerControl1 != null)
            //    this.axFramerControl1.Close();
            //if (this.m_axFramerControl != null && axFramerControl1_is == 1)
            //{
            //    Save();
            //    this.m_axFramerControl.Close();
            //}
            if (this.webBrowser1.Document != null)
                this.webBrowser1.Stop();

            NAR(oSheet);
            if (oBook != null)
            {
                try
                {
                    oBook.Close(false);
                    NAR(oBook);
                    NAR(oBooks);
                    if (oApp != null)
                        oApp.Quit();
                }
                catch
                {


                }
            }
            if (oApp != null)
                NAR(oApp);
            Debug.WriteLine("Sleeping...");
            // System.Threading.Thread.Sleep(5000);
            Debug.WriteLine("End Excel");
            //webBrowser1.Stop();
            //webBrowser1.Dispose();
            System.Environment.Exit(0);



        }
        public void Save()
        {
            try
            {
                //先保存
                this.m_axFramerControl.Save(true, true, "", "");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void NAR(Object o)
        {
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(o); }
            catch { }
            finally { o = null; }
        }

        private void frmPrice_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (MessageBox.Show("是否已经保存桌面其他Excel文件, 继续 ?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {

            }
            else
                return;

            toolStripLabel3.Text = "刷新中,请稍等...";

            string folderpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClearTask.bat");

            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.WorkingDirectory = folderpath;
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = folderpath;
            p.Start();


            ExcelExit();


        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Object refmissing = System.Reflection.Missing.Value;
            //axWebBrowser2.ExecWB(SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, ref refmissing, ref refmissing);

        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否已经保存桌面其他Excel文件, 继续 ?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {

            }
            else
                return;

            toolStripLabel3.Text = "刷新中,请稍等...";

            string folderpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClearTask.bat");

            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.WorkingDirectory = folderpath;
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = folderpath;
            p.Start();

            ExcelExit();

            this.Close();

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否已经保存桌面其他Excel文件, 继续 ?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {

            }
            else
                return;

            toolStripLabel3.Text = "刷新中,请稍等...";

            string folderpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClearTask.bat");

            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.WorkingDirectory = folderpath;
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = folderpath;
            p.Start();

            if (this.tabControl1.SelectedIndex == 2)
                toolStripButton4_Click(null, EventArgs.Empty);
            //if (this.tabControl1.SelectedIndex == 1)
            else

                toolStripButton4_Click(null, EventArgs.Empty);
        }

        public List<cls_sixzhuanjiagebiao_info> GetKEYnfo(string Alist)
        {

            List<cls_sixzhuanjiagebiao_info> MAPPINGResult = new List<cls_sixzhuanjiagebiao_info>();
            try
            {
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {
                    string path = Alist;
                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 46]];
                    int rowCount = WS.UsedRange.Rows.Count;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;
                    clsCommHelp.CloseExcel(excelApp, analyWK);

                    for (int i = 2; i <= rowCount; i++)
                    {
                        cls_sixzhuanjiagebiao_info temp = new cls_sixzhuanjiagebiao_info();

                        #region 基础信息


                        //temp.touxing_B = "";
                        //if (o[i, 2] != null)
                        //    temp.touxing_B = o[i, 2].ToString().Trim();

                        //if (temp.touxing_B == null || temp.touxing_B == "")
                        //    continue;

                        //temp.zuzuangbizhong_C = "";
                        //if (o[i, 3] != null)
                        //    temp.zuzuangbizhong_C = o[i, 3].ToString().Trim();

                        ////卖场代码

                        //temp.shapianbizhong_D = "";
                        //if (o[i, 4] != null)
                        //    temp.shapianbizhong_D = o[i, 4].ToString().Trim();

                        //temp.dianpiandanjia_E = "";
                        //if (o[i, 5] != null)
                        //    temp.dianpiandanjia_E = o[i, 5].ToString().Trim();

                        //temp.handianpianqianjia_F = "";
                        //if (o[i, 6] != null)
                        //    temp.handianpianqianjia_F = o[i, 6].ToString().Trim();
                        //temp.handianpiandunjia_G = "";
                        //if (o[i, 7] != null)
                        //    temp.handianpiandunjia_G = o[i, 7].ToString().Trim();

                        //temp.guigexinghao_H = "";
                        //if (o[i, 8] != null)
                        //    temp.guigexinghao_H = o[i, 8].ToString().Trim();

                        ////卖场名称
                        //temp.guige_I = "";
                        //if (o[i, 9] != null)
                        //    temp.guige_I = o[i, 9].ToString().Trim();

                        //temp.bizhong_J = "";
                        //if (o[i, 10] != null)
                        //    temp.bizhong_J = o[i, 10].ToString().Trim();


                        //temp.ganjia_K = "";
                        //if (o[i, 11] != null)
                        //    temp.ganjia_K = o[i, 11].ToString().Trim();

                        //temp.dunjia_L = "";
                        //if (o[i, 12] != null)
                        //    temp.dunjia_L = o[i, 12].ToString().Trim();

                        //temp.yuanmei_M = "";
                        //if (o[i, 13] != null)
                        //    temp.yuanmei_M = o[i, 13].ToString().Trim();

                        //temp.gongxu6_N = "";
                        //if (o[i, 14] != null)
                        //    temp.gongxu6_N = o[i, 14].ToString().Trim();

                        //temp.yuanmei_O = "";
                        //if (o[i, 15] != null)
                        //    temp.yuanmei_O = o[i, 15].ToString().Trim();

                        //temp.gongxu5_P = "";
                        //if (o[i, 16] != null)
                        //    temp.gongxu5_P = o[i, 16].ToString().Trim();



                        //temp.shujin_Q = "";
                        //if (o[i, 17] != null)
                        //    temp.shujin_Q = o[i, 17].ToString().Trim();


                        //temp.yunfei_R = "";
                        //if (o[i, 18] != null)
                        //    temp.yunfei_R = o[i, 18].ToString().Trim();


                        //temp.gongxu4_S = "";
                        //if (o[i, 19] != null)
                        //    temp.gongxu4_S = o[i, 19].ToString().Trim();


                        //temp.gongxu3_T = "";
                        //if (o[i, 20] != null)
                        //    temp.gongxu3_T = o[i, 20].ToString().Trim();


                        //temp.gongxu2_U = "";
                        //if (o[i, 21] != null)
                        //    temp.gongxu2_U = o[i, 21].ToString().Trim();

                        //temp.chengpinsi_V = "";
                        //if (o[i, 22] != null)
                        //    temp.chengpinsi_V = o[i, 22].ToString().Trim();

                        //temp.shunhao_W = "";
                        //if (o[i, 23] != null)
                        //    temp.shunhao_W = o[i, 23].ToString().Trim();

                        //temp.panyuan2_X = "";
                        //if (o[i, 24] != null)
                        //    temp.panyuan2_X = o[i, 24].ToString().Trim();

                        //temp.gongxu1_Y = "";
                        //if (o[i, 25] != null)
                        //    temp.gongxu1_Y = o[i, 25].ToString().Trim();

                        //temp.panyuan1_Z = "";
                        //if (o[i, 26] != null)
                        //    temp.panyuan1_Z = o[i, 26].ToString().Trim();

                        //temp.shunhao_AA = "";
                        //if (o[i, 27] != null)
                        //    temp.shunhao_AA = o[i, 27].ToString().Trim();

                        //temp.panyuan_AB = "";
                        //if (o[i, 28] != null)
                        //    temp.panyuan_AB = o[i, 28].ToString().Trim();


                        #endregion


                        #region 2

                        #region 基础信息

                        temp.touxing_B = "";
                        if (o[i, 1] != null)
                            temp.touxing_B = o[i, 1].ToString().Trim();


                        temp.Order_id = "";
                        if (o[i, 2] != null)
                            temp.Order_id = o[i, 2].ToString().Trim();

                        temp.zuzuangbizhong_C = "";
                        if (o[i, 3] != null)
                            temp.zuzuangbizhong_C = o[i, 3].ToString().Trim();

                        if (temp.touxing_B == null || temp.touxing_B == "")
                            continue;

                        temp.shapianbizhong_D = "";
                        if (o[i, 4] != null)
                            temp.shapianbizhong_D = o[i, 4].ToString().Trim();

                        //卖场代码

                        temp.dianpiandanjia_E = "";
                        if (o[i, 5] != null)
                            temp.dianpiandanjia_E = o[i, 5].ToString().Trim();

                        temp.handianpianqianjia_F = "";
                        if (o[i, 6] != null)
                            temp.handianpianqianjia_F = String.Format("{0:N2}", Convert.ToDouble(Math.Round(Convert.ToDouble(o[i, 6].ToString()), 2).ToString())); //o[i, 6].ToString().Trim();

                        temp.guige_I = "";
                        if (o[i, 7] != null)
                            temp.guige_I = o[i, 7].ToString().Trim();
                        temp.bizhong_J = "";
                        if (o[i, 8] != null)
                            temp.bizhong_J = o[i, 8].ToString().Trim();

                        temp.ganjia_K = "";
                        if (o[i, 9] != null)
                            temp.ganjia_K = String.Format("{0:N2}", Convert.ToDouble(Math.Round(Convert.ToDouble(o[i, 9].ToString()), 2).ToString())); //o[i, 8].ToString().Trim();


                        temp.handianpiandunjia_G = "";
                        if (o[i, 10] != null)
                            temp.handianpiandunjia_G = String.Format("{0:N2}", Convert.ToDouble(Math.Round(Convert.ToDouble(o[i, 10].ToString()), 2).ToString())); //o[i, 8].ToString().Trim();
                        //o[i, 10].ToString().Trim();


                        temp.dunjia_L = "";
                        if (o[i, 11] != null)
                            temp.dunjia_L = o[i, 11].ToString().Trim();



                        #endregion
                        #endregion
                        MAPPINGResult.Add(temp);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return MAPPINGResult;

        }

        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }
        private void CheckBankCharge(object sender, DoWorkEventArgs e)
        {
            DateTime oldDate = DateTime.Now;

            AddOfficeControl(ZFCEPath);

            //InitOfficeControl(ZFCEPath);

            //this.axFramerControl1.Open(ZFCEPath);
            //this.axFramerControl1.ShowView(0);
            //var WorkBook = this.axFramerControl1.ActiveDocument as Workbook;
            //(WorkBook.Sheets[2] as Worksheet).Activate();


            DateTime FinishTime = DateTime.Now;  //   
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();


            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);
        }
        public void Init(string _sFilePath)
        {
            try
            {

                AddOfficeControl(_sFilePath);
                //这里一定要先把dso控件加到界面上才能初始化dso控件,
                //这个dso控件在没有被show出来之前是不能进行初始化操作的,很奇怪为什么作者这样考虑.....
                //InitOfficeControl(_sFilePath);
            }
            catch (Exception ex)
            {

                return;

                throw ex;
            }
        }
        private void AddOfficeControl(string ZFCEPath)
        {
            try
            {

                ((System.ComponentModel.ISupportInitialize)(this.m_axFramerControl)).BeginInit();
                this.Controls.Add(m_axFramerControl);
                ((System.ComponentModel.ISupportInitialize)(this.m_axFramerControl)).EndInit();



                this.panel1.Controls.Add(m_axFramerControl);

                m_axFramerControl.Dock = DockStyle.Fill;
                m_axFramerControl.Titlebar = false;

                //    ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "圣丰生产管理系统表\\圣丰生产管理系统表.xlsx");

                m_axFramerControl.Open(ZFCEPath, false, "Excel.Sheet", "", "");

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("服务器出现意外情况"))

                    throw ex;
            }
        }
        private void InitOfficeControl(string _sFilePath)
        {
            try
            {
                if (m_axFramerControl == null)
                {
                    throw new ApplicationException("请先初始化office控件对象！");
                }

                //this.m_axFramerControl.SetMenuDisplay(48);
                //这个方法很特别，一个组合菜单控制方法，我还没有找到参数的规律，有兴趣的朋友可以研究一下
                string sExt = System.IO.Path.GetExtension(_sFilePath).Replace(".", "");
                //this.m_axFramerControl.CreateNew(this.LoadOpenFileType(sExt));//创建新的文件

                this.m_axFramerControl.Open(_sFilePath, false, this.LoadOpenFileType(sExt), "", "");//打开文件

                //隐藏标题
                this.m_axFramerControl.Titlebar = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private string LoadOpenFileType(string _sExten)
        {
            try
            {
                string sOpenType = "";
                switch (_sExten.ToLower())
                {
                    case "xls":
                        sOpenType = "Excel.Sheet";
                        break;
                    case "doc":
                        sOpenType = "Word.Document";
                        break;
                    case "ppt":
                        sOpenType = "PowerPoint.Show";
                        break;
                    case "vsd":
                        sOpenType = "Visio.Drawing";
                        break;
                    default:
                        sOpenType = "Word.Document";
                        break;
                }
                return sOpenType;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void frmJiaqizhuantongjibiao_Load(object sender, EventArgs e)
        {
            //thOpen = new Thread(new ThreadStart(FOpen));
            //((System.ComponentModel.ISupportInitialize)(this.m_axFramerControl)).BeginInit();
            //m_axFramerControl.Dock = System.Windows.Forms.DockStyle.Fill;
            //m_axFramerControl.Enabled = true;
            //m_axFramerControl.Location = new System.Drawing.Point(0, 0);
            //this.tabControl1.TabPages[2].Name = "spc_Excel";


            //this.tabControl1.TabPages[2].Controls.Add(m_axFramerControl);
            ////m_axFramerControl.Titlebar = false;
            ////m_axFramerControl.Menubar = false;
            ////m_axFramerControl.Toolbars = true;

            //((System.ComponentModel.ISupportInitialize)(this.m_axFramerControl)).EndInit();


            ////启动现成加载EXCEL
            //thOpen.Start();
        }
        private void FOpen()
        {

            lock (m_axFramerControl)
            {
                try
                {
                    m_axFramerControl.Open(ZFCEPath, false, "Excel.Sheet", "", "");
                    //xBook = (Workbook)m_axFramerControl.ActiveDocument;
                    //// xSheet = (xBook.Worksheets[1]);
                    //xSheet = (Worksheet)xBook.ActiveSheet;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }

            }


        }

        private void toolStripButton2_Click_1(object sender, EventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
        void MyWebBrower_BeforeNewWindow2(object sender, WebBrowserExtendedNavigatingEventArgs e)
        {
            #region 在原有窗口导航出新页
            e.Cancel = true;//http://pro.wwpack-crest.hp.com/wwpak.online/regResults.aspx
            //MyWebBrower.Navigate(e.Url);
            #endregion
        }
        protected void AnalysisWebInfo2(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Workbook wbb = null;
                Object refmissing = System.Reflection.Missing.Value;

                object[] args = new object[4];

                args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;

                args[1] = SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER;

                args[2] = refmissing;

                args[3] = refmissing;
                object axWebBrowser = this.MyWebBrower.ActiveXInstance;

                axWebBrowser.GetType().InvokeMember("ExecWB", BindingFlags.InvokeMethod, null, axWebBrowser, args);


                object oApplication = axWebBrowser.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, axWebBrowser, null);

                wbb = (Microsoft.Office.Interop.Excel.Workbook)oApplication;

                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)wbb.Worksheets[1];
                oBook = wbb;

                //Microsoft.Office.Interop.Excel.Application ExcelApp;
                //ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                //  oApp = (Microsoft.Office.Interop.Excel.Application)oApplication;

                oSheet = WS;

            }
            catch (Exception ex)
            {
                MessageBox.Show("异常出现：可以关闭桌面左右Excel 然后点击【强制刷新按钮】重试" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;


                throw;
            }

        }
        private void toolStripButton3_Click_1(object sender, EventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted_1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Workbook wbb = null;
                Object refmissing = System.Reflection.Missing.Value;

                object[] args = new object[4];

                args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;

                args[1] = SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER;

                args[2] = refmissing;

                args[3] = refmissing;
                object axWebBrowser = this.webBrowser1.ActiveXInstance;

                axWebBrowser.GetType().InvokeMember("ExecWB", BindingFlags.InvokeMethod, null, axWebBrowser, args);


                object oApplication = axWebBrowser.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, axWebBrowser, null);

                wbb = (Microsoft.Office.Interop.Excel.Workbook)oApplication;

                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)wbb.Worksheets[1];
                oBook = wbb;

                //Microsoft.Office.Interop.Excel.Application ExcelApp;
                //ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                //oApp = (Microsoft.Office.Interop.Excel.Application)oApplication;

                oSheet = WS;
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("调用的目标发生了异常"))
                    MessageBox.Show("异常出现：可以关闭桌面左右Excel 然后点击【强制刷新按钮】重试" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

                throw;
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            try
            {

                MessageBox.Show("提示：" + "1.如上次系统非正常关闭请点击强制刷新按钮即可重新打开\r\n2.如个别系统有二次认证需要到桌面查看登陆！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);



                //ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "辣皇后\\Gongchuang.xlsm");

                //InitialWebbroswer1();
                timerstart();

                //return;

                this.panel1.Controls.Add(MyWebBrower);

                toolStripLabel2.Text = "读取中，请耐心等待...(打开快慢受网络状况和表格大小影响)";
                this.tabControl1.SelectedIndex = 1;

                this.webBrowser1.Navigate(ZFCEPath);

                toolStripLabel2.Text = "读取完成,马上显示";
                // 打开制定的本地文件
                //axf.Open( Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "辣皇后\\共创联盟系统(4-139).xlsm"));
                //制定用Word来打开c:\plain.txt文件

            }
            catch (Exception ex)
            {
                MessageBox.Show("异常：" + ex);

                throw;
            }

        }

        public void InitialWebbroswer1()
        {
            try
            {
                MyWebBrower = new WbBlockNewUrl();
                //不显示弹出错误继续运行框（HP方可）
                MyWebBrower.ScriptErrorsSuppressed = true;
                MyWebBrower.BeforeNewWindow += new EventHandler<WebBrowserExtendedNavigatingEventArgs>(MyWebBrower_BeforeNewWindow);
                MyWebBrower.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(AnalysisWebInfo1);
                MyWebBrower.Dock = DockStyle.Fill;
                //显示用的窗体
                viewForm = new Form();
                //viewForm.Icon=
                viewForm.ClientSize = new System.Drawing.Size(800, 600);
                viewForm.StartPosition = FormStartPosition.CenterScreen;
                viewForm.Controls.Clear();
                viewForm.Controls.Add(MyWebBrower);
                viewForm.FormClosing += new FormClosingEventHandler(viewForm_FormClosing);
                //显示窗体
                viewForm.Show();

                //string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xls");

                //MyWebBrower.Url = new Uri(ZFCEPath);

                Object refmissing = System.Reflection.Missing.Value;
                MyWebBrower.Navigate(ZFCEPath, refmissing.ToString());

                //MyWebBrower.Navigate(ZFCEPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        void MyWebBrower_BeforeNewWindow(object sender, WebBrowserExtendedNavigatingEventArgs e)
        {
            string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xls");

            #region 在原有窗口导航出新页
            e.Cancel = false;//http://pro.wwpack-crest.hp.com/wwpak.online/regResults.aspx
            //   MyWebBrower.Navigate(ZFCEPath);
            #endregion
        }
        private void viewForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //if (toolStripStatusLabel1.Text != " Search Finished  !")
            {
                if (MessageBox.Show("正在进行，是否中止?", "Sign Out", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    if (MyWebBrower != null)
                    {
                        if (MyWebBrower.IsBusy)
                        {
                            MyWebBrower.Stop();
                        }
                        MyWebBrower.Dispose();
                        MyWebBrower = null;
                    }
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }
        protected void AnalysisWebInfo1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

            timerstart();

        }

        private void timerstart()
        {
            int file = 0;

            {
                strFileName = DateTime.Now;//开始时间
                aTimer.Elapsed += new ElapsedEventHandler(steap_tocover);
                aTimer.Start();
                file = 1;
                //HtmlElement btnAdd = doc.GetElementById("addDiv").FirstChild;
                //btnAdd.InvokeMember("Click");

            }
        }
        private void stoprefeach(object sender, EventArgs e)
        {
            DateTime rq2 = DateTime.Now;  //结束时间
            int a = rq2.Second - strFileName.Second;
            if (a >= 3 || rq2.Second < strFileName.Second)
            {
                aTimer.Stop();
                return;
            }
        }
        private void steap_tocover(object sender, EventArgs e)
        {
            SaveAs();

        }

        #region AＰI

        //第一步保存
        //private void SaveAs(object sender, EventArgs e)
        private void SaveAs()
        {
            DateTime rq2 = DateTime.Now;  //结束时间               
            TimeSpan ts = rq2 - strFileName;
            int a = rq2.Second - strFileName.Second;
            //if (a >= 10 || rq2.Second < strFileName.Second)
            if (ts.Minutes > 3)
            {
                isReadyForSearch = true;
                aTimer.Stop();
                //  isOneFinished = true;
                return;

            }

            List<IntPtr> arrHwnd_Sap_before = getSAPWindow();


            // log4net.ILog objLogger = log4net.LogManager.GetLogger("SystemExceptionLogger");

            //bool blFresh = false;
            RUNING = true;
            try
            {
                //得到File Download窗体对象
                IntPtr h1 = FindWindow("#32770", "File Download - Security Warning");
                IntPtr h2 = FindWindow("#32770", "文件下载 - 安全警告");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "File Download");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "文件下载");

                if (h2.ToInt32() > 0 || h1.ToInt32() > 0)
                {
                    //得到Save按钮对象
                    IntPtr duixiang;

                    if (h1.ToInt32() > 0)
                    {
                        duixiang = WinAPIuser32.FindWindowEx(h1, 0, "Button", "&Save");
                    }
                    else
                    {
                        //duixiang = WinAPIuser32.FindWindowEx(h2, 0, "BUTTON", "&Save");
                        duixiang = WinAPIuser32.FindWindowEx(h2, 0, "BUTTON", "打开(&O)");
                    }
                    //objLogger.Fatal("BUTTON 保存(&S):" + duixiang.ToString());
                    //如果得到点击Save按钮
                    if (duixiang.ToInt32() > 0)
                    {
                        Thread.Sleep(600);
                        blFresh = true;
                        SendMessage(duixiang, 0xF5, 0, 0);
                        SendMessage(duixiang, 0xF5, 0, 0);
                        WinAPIuser32.SendMessage(duixiang, WM_LBUTTONUP, IntPtr.Zero, null);
                        WinAPIuser32.SendMessage(duixiang, WM_LBUTTONUP, IntPtr.Zero, null);
                        //设定File Download窗体是否点击变量为已点击

                    }
                    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "&Save");
                    if (hwnd2.ToInt32() <= 0)
                        hwnd2 = WinAPIuser32.FindWindow("#32770", "Save");
                    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());
                }

                RUNING = false;
                #region MyRegion
                //if (true)
                //{

                //    //得到Save As窗体对象
                //    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                //    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "Save As");
                //    if (hwnd2.ToInt32() <= 0)
                //        hwnd2 = WinAPIuser32.FindWindow("#32770", "名前を付けて保存");

                //    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());

                //    if (hwnd.ToInt32() > 0 || hwnd2.ToInt32() > 0)
                //    {
                //        //得到其下的一系列子窗体对象

                //        IntPtr htextbox;
                //        IntPtr htextbox1;

                //        if (hwnd.ToInt32() > 0)
                //        {
                //            htextbox = WinAPIuser32.FindWindowEx(hwnd, 0, "ComboBoxEx32", null);
                //            htextbox1 = WinAPIuser32.FindWindowEx(hwnd, 0, "DUIViewWndClassName", null);
                //        }
                //        else
                //        {
                //            htextbox = WinAPIuser32.FindWindowEx(hwnd2, 0, "ComboBoxEx32", null);
                //            htextbox1 = WinAPIuser32.FindWindowEx(hwnd2, 0, "DUIViewWndClassName", null);
                //        }

                //        //objLogger.Fatal("ComboBoxEx32" + htextbox.ToString() + htextbox1.ToString());

                //        if (htextbox.ToInt32() > 0)
                //        {
                //            IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox, 0, "ComboBox", null);
                //            if (htextbox4.ToInt32() > 0)
                //            {
                //                Thread.Sleep(1000);
                //                //得到子窗体中输入保存路径文本框

                //                IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                //                string strPath = FileDownURL + FileName;
                //                WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                //                //得到Save按钮对象
                //                IntPtr hbutton;
                //                if (hwnd.ToInt32() > 0)
                //                {
                //                    hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                //                }
                //                else
                //                {
                //                    hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                //                    if (hbutton.ToInt32() <= 0)
                //                        hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                //                }

                //                //如果得到点击Save按钮
                //                if (hbutton.ToInt32() > 0)
                //                {
                //                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                //                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                //                    //停止每5秒点击Save As窗体Timer控件
                //                    WebSiteStatus = true;
                //                }
                //            }
                //        }
                //        else if (htextbox1.ToInt32() > 0)
                //        {
                //            IntPtr htextbox2 = WinAPIuser32.FindWindowEx(htextbox1, 0, "DirectUIHWND", null);
                //            if (htextbox2.ToInt32() > 0)
                //            {
                //                IntPtr htextbox3 = WinAPIuser32.FindWindowEx(htextbox2, 0, "FloatNotifySink", null);
                //                if (htextbox3.ToInt32() > 0)
                //                {
                //                    IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox3, 0, "ComboBox", null);
                //                    if (htextbox4.ToInt32() > 0)
                //                    {
                //                        Thread.Sleep(1000);
                //                        //得到子窗体中输入保存路径文本框
                //                        // FileName = @"C:";


                //                        ///////////////////////////////////////////////////////////路径保存地址
                //                        FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources\\PDF\\");
                //                        string strPath = FileDownURL + FileName + publicPDFName + ".pdf";
                //                        if (File.Exists(strPath))
                //                        {
                //                            File.Delete(strPath);

                //                        }
                //                        IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                //                        WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                //                        //得到Save按钮对象
                //                        IntPtr hbutton;
                //                        if (hwnd.ToInt32() > 0)
                //                        {
                //                            hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                //                        }
                //                        else
                //                        {
                //                            hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                //                            if (hbutton.ToInt32() <= 0)
                //                                hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                //                        }
                //                        //如果得到点击Save按钮
                //                        if (hbutton.ToInt32() > 0)
                //                        {
                //                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                //                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                //                            //停止每5秒点击Save As窗体Timer控件
                //                            WebSiteStatus = true;
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }

                //} 
                #endregion

            }
            catch (Exception ex)
            {
                RUNING = false;
                throw;
            }
            //    isReadyForSearch = false;
            //    WebSiteStatus = false;
            //    aTimer.Stop();


        }

        //第二部保存
        private void OnTimedEvent(object sender, EventArgs e)
        {
            Thread.Sleep(3000);
            //IntPtr hwnd = FindWindow("#32770", "另存为");
            IntPtr hwnd = FindWindow("#32770", "Save As");
            if (int.Parse(hwnd.ToString()) == 0)
            {
                hwnd = FindWindow("#32770", "Save");
            }
            if (int.Parse(hwnd.ToString()) > 0)
            {
                IntPtr hbutton = GetDlgItem(hwnd, 1);
                if (int.Parse(hbutton.ToString()) > 0)
                {
                    SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                    SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                    aTimer.Stop();
                }
            }

        }


        public void SaveAPIButton()
        {
            bool RUNING = false;

            //  List<IntPtr> arrHwnd_Sap_before = getSAPWindow();


            // log4net.ILog objLogger = log4net.LogManager.GetLogger("SystemExceptionLogger");

            bool blFresh = false;
            RUNING = true;
            try
            {
                //得到File Download窗体对象
                IntPtr h1 = FindWindow("#32770", "File Download - Security Warning");
                IntPtr h2 = FindWindow("#32770", "文件下载 - 安全警告");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "ファイルのダウンロード");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "文件下载");

                if (h2.ToInt32() > 0 || h1.ToInt32() > 0)
                {
                    //得到Save按钮对象
                    IntPtr duixiang;

                    if (h1.ToInt32() > 0)
                    {
                        duixiang = WinAPIuser32.FindWindowEx(h1, 0, "Button", "&Save");
                    }
                    else
                    {
                        duixiang = WinAPIuser32.FindWindowEx(h2, 0, "BUTTON", "保存(&S)");
                    }
                    //objLogger.Fatal("BUTTON 保存(&S):" + duixiang.ToString());
                    //如果得到点击Save按钮
                    if (duixiang.ToInt32() > 0)
                    {
                        SendMessage(duixiang, 0xF5, 0, 0);
                        SendMessage(duixiang, 0xF5, 0, 0);
                        WinAPIuser32.SendMessage(duixiang, WM_LBUTTONUP, IntPtr.Zero, null);

                        //设定File Download窗体是否点击变量为已点击
                        blFresh = true;
                    }
                    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "Save As");
                    if (hwnd2.ToInt32() <= 0)
                        hwnd2 = WinAPIuser32.FindWindow("#32770", "名前を付けて保存");

                    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());
                }
                RUNING = false;

                #region MyRegion
                //if (true)
                //{
                //    //得到Save As窗体对象
                //    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                //    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "Save As");
                //    if (hwnd2.ToInt32() <= 0)
                //        hwnd2 = WinAPIuser32.FindWindow("#32770", "名前を付けて保存");

                //    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());

                //    if (hwnd.ToInt32() > 0 || hwnd2.ToInt32() > 0)
                //    {
                //        //得到其下的一系列子窗体对象

                //        IntPtr htextbox;
                //        IntPtr htextbox1;

                //        if (hwnd.ToInt32() > 0)
                //        {
                //            htextbox = WinAPIuser32.FindWindowEx(hwnd, 0, "ComboBoxEx32", null);
                //            htextbox1 = WinAPIuser32.FindWindowEx(hwnd, 0, "DUIViewWndClassName", null);
                //        }
                //        else
                //        {
                //            htextbox = WinAPIuser32.FindWindowEx(hwnd2, 0, "ComboBoxEx32", null);
                //            htextbox1 = WinAPIuser32.FindWindowEx(hwnd2, 0, "DUIViewWndClassName", null);
                //        }

                //        //objLogger.Fatal("ComboBoxEx32" + htextbox.ToString() + htextbox1.ToString());

                //        if (htextbox.ToInt32() > 0)
                //        {
                //            IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox, 0, "ComboBox", null);
                //            if (htextbox4.ToInt32() > 0)
                //            {
                //                Thread.Sleep(1000);
                //                //得到子窗体中输入保存路径文本框

                //                IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                //                string strPath = FileDownURL + FileName;
                //                WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                //                //得到Save按钮对象
                //                IntPtr hbutton;
                //                if (hwnd.ToInt32() > 0)
                //                {
                //                    hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                //                }
                //                else
                //                {
                //                    hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                //                    if (hbutton.ToInt32() <= 0)
                //                        hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                //                }

                //                //如果得到点击Save按钮
                //                if (hbutton.ToInt32() > 0)
                //                {
                //                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                //                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                //                    //停止每5秒点击Save As窗体Timer控件
                //                    WebSiteStatus = true;
                //                }
                //            }
                //        }
                //        else if (htextbox1.ToInt32() > 0)
                //        {
                //            IntPtr htextbox2 = WinAPIuser32.FindWindowEx(htextbox1, 0, "DirectUIHWND", null);
                //            if (htextbox2.ToInt32() > 0)
                //            {
                //                IntPtr htextbox3 = WinAPIuser32.FindWindowEx(htextbox2, 0, "FloatNotifySink", null);
                //                if (htextbox3.ToInt32() > 0)
                //                {
                //                    IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox3, 0, "ComboBox", null);
                //                    if (htextbox4.ToInt32() > 0)
                //                    {
                //                        Thread.Sleep(1000);
                //                        //得到子窗体中输入保存路径文本框
                //                        FileName = @"C:";
                //                        IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                //                        string strPath = FileDownURL + FileName;
                //                        WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                //                        //得到Save按钮对象
                //                        IntPtr hbutton;
                //                        if (hwnd.ToInt32() > 0)
                //                        {
                //                            hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                //                        }
                //                        else
                //                        {
                //                            hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                //                            if (hbutton.ToInt32() <= 0)
                //                                hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                //                        }
                //                        //如果得到点击Save按钮
                //                        if (hbutton.ToInt32() > 0)
                //                        {
                //                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                //                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                //                            //停止每5秒点击Save As窗体Timer控件
                //                            WebSiteStatus = true;
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }
                //} 
                #endregion
            }
            catch (Exception ex)
            {
                RUNING = false;
                throw;
            }
        }

        #region Windows API Function

        private List<IntPtr> getSAPWindow()
        {
            IntPtr hwnd_sap = IntPtr.Zero;
            List<IntPtr> arrHwnd = new List<IntPtr>();

            while (true)
            {
                hwnd_sap = FindWindowEx(IntPtr.Zero, hwnd_sap, "SAP_FRONTEND_SESSION", null);
                if (hwnd_sap != IntPtr.Zero)
                {
                    arrHwnd.Add(hwnd_sap);
                }
                else
                {

                    break;
                }
            }

            return arrHwnd;
        }

        private void monitorSAP()
        {
            if (hwnd_main == IntPtr.Zero)
            {
                //MessageBox.Show("找不到SAP主窗口！");
                return;
            }

            if (IsWindowVisible(hwnd_main))
            {
                if (ScreenStatus == 0)
                {
                    hwnd_ReportTree1 = findReportTree1();
                    //MessageBox.Show(hwnd_ReportTree1.ToInt32().ToString());
                    selectReport(intCnt);

                    IntPtr btnExecute = findExecuteButton();
                    SendMessage(btnExecute, BM_CLICK, 0, 0);
                    ScreenStatus = 1;
                }
                //clickYes(btnExecute);
            }
        }

        private IntPtr findExecuteButton()
        {
            IntPtr children = FindWindowEx(hwnd_main, IntPtr.Zero, null, "");
            while (children != IntPtr.Zero)
            {
                children = FindWindowEx(hwnd_main, children, null, "");
                int nRet;
                StringBuilder ClassName = new StringBuilder(100);
                //Get the window class name
                nRet = GetClassName(children, ClassName, ClassName.Capacity);
                Regex r = new Regex("Afx:[(a-z)|(A-Z)|(0-9)]{8}:8:[0-9]{8}:00000000:00000000");
                if (nRet != 0 && r.Match(ClassName.ToString()).Success)
                {
                    IntPtr hwnd_level2 = FindWindowEx(children, IntPtr.Zero, "Button", null);
                    if (hwnd_level2 != IntPtr.Zero)
                    {
                        return hwnd_level2;
                    }
                }
            }
            return IntPtr.Zero;
        }

        private void clickYes(IntPtr hwnd_Control)
        {
            IntPtr hwnd_Button = FindWindowEx(hwnd_Control, new IntPtr(0), "Button", null);
            SendMessage(hwnd_Button, BM_CLICK, 0, 0);
        }

        private void selectReport(int intCount)
        {
            sendPageUp();
            sendTab(intCount);
        }

        private void sendPageUp()
        {
            for (int i = 0; i < 5; i++)
            {
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_CONTROL, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_PRIOR, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_PRIOR, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_CONTROL, 0);
            }
        }

        private void sendTab(int intCount)
        {
            //Tab
            for (int i = 0; i < intCount; i++)
            {
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_TAB, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_TAB, 0);
            }
        }

        private IntPtr findReportTree1()
        {
            IntPtr hwnd_level1, hwnd_level2, hwnd_level3;
            hwnd_level1 = FindWindowEx(hwnd_main, IntPtr.Zero, "Docking Container Class", null);
            //MessageBox.Show("Docking Container Class:" + hwnd_level1.ToInt32().ToString());
            if (hwnd_level1 != IntPtr.Zero)
            {
                hwnd_level2 = FindWindowEx(hwnd_level1, IntPtr.Zero, "Shell Window Class", "Control  Container");
                //MessageBox.Show("Control  Container:" + hwnd_level2.ToInt32().ToString());
                if (hwnd_level2 != IntPtr.Zero)
                {
                    hwnd_level3 = FindWindowEx(hwnd_level2, IntPtr.Zero, "AfxOleControl80", null);
                    if (hwnd_level3 == IntPtr.Zero)
                        hwnd_level3 = FindWindowEx(hwnd_level2, IntPtr.Zero, "AfxOleControl90", null);
                    //MessageBox.Show("AfxOleControl80:" + hwnd_level3.ToInt32().ToString());
                    if (hwnd_level3 != IntPtr.Zero)
                    {
                        //MessageBox.Show("SAPTreeList:" + FindWindowEx(hwnd_level3, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist").ToInt32().ToString());
                        return FindWindowEx(hwnd_level3, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist");
                        //return FindWindowEx(hwnd_level2, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist");
                    }
                }
            }
            return IntPtr.Zero;
        }

        private IntPtr findControlWindow()
        {
            string strCaption = "Execute Project Report: Initial Screen";

            IntPtr hwnd_Child = IntPtr.Zero;
            while (true)
            {
                hwnd_Child = FindWindowEx(IntPtr.Zero, hwnd_Child, "#32770", strCaption);
                if (GetParent(hwnd_Child) == hwnd_main || hwnd_Child == IntPtr.Zero)
                {
                    break;
                }
            }
            return hwnd_Child;
        }
        #endregion
        #endregion

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            //自动点击弹出确认或弹出提示
            //IHTMLDocument2 vDocument = (IHTMLDocument2)webBrowser1.Document.DomDocument;
            //vDocument.parentWindow.execScript("function confirm(str){return true;} ", "javascript"); //弹出确认
            //vDocument.parentWindow.execScript("function alert(str){return true;} ", "javaScript");//弹出提示
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            string marcro = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "辣皇后\\Gongchuang.xlsm");
            //  marcro = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "圣丰生产管理系统表\\圣丰生产管理系统表.xlsx");
            //System.Diagnostics.Process.Start(marcro);
            //  AddOfficeControl(marcro);

            ZFCEPath = marcro;

            axFramerControl1_is = 1;

            this.tabControl1.SelectedIndex = 2;
            toolStripLabel1.Text = "读取中，请耐心等待...(打开快慢受网络情况影响)";

            toolStripButton3_Click_1(null, EventArgs.Empty);
            //  Init(marcro);
            //this.axFramerControl1.Open(ZFCEPath);
            toolStripLabel1.Text = "读取完成,马上显示";
            this.tabControl1.SelectedIndex = 2;

            MessageBox.Show("读取完成，请查看", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);


            return;

            this.tabControl1.SelectedIndex = 1;
            if (MessageBox.Show("为了保证数据丢失，是否已保存其他桌面Excel文件, 继续 ?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {

            }
            else
                return;

            toolStripLabel2.Text = "查询中,请稍等...";

            string folderpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClearTask.bat");

            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.WorkingDirectory = folderpath;
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = folderpath;
            p.Start();

            Thread.Sleep(2000);

            string c = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Gongchuang.xlsm");
            string destFile = ZFCEPath.Replace("xlsx", "xlsm");
            // destFile = @"C:\Windows" + "\\dsoframer.ocx";

            int io = 0;

            if (File.Exists(destFile))
            {
                toolStripLabel2.Text = "打开中,马上显示";
                File.Copy(destFile, c, true);//覆盖模式
                io = 1;

                if (io == 1)
                {

                    Thread.Sleep(8000);
                    System.Diagnostics.Process.Start(c);

                }
            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            try
            {
                this.tabControl1.SelectedIndex = 1;
                if (MessageBox.Show("是否已经保存其他桌面Excel文件, 继续 ?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {

                }
                else
                    return;

                toolStripLabel2.Text = "查询中,请稍等...";

                string folderpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClearTask.bat");

                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.WorkingDirectory = folderpath;
                p.StartInfo.UseShellExecute = true;
                p.StartInfo.FileName = folderpath;
                p.Start();

                MessageBox.Show("已关闭");

            }
            catch (Exception ex)
            {
                MessageBox.Show("异常：" + ex.Message);
                return;


            }
        }
    }
}
