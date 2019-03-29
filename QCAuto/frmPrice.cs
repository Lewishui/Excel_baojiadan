using China_System.Common;
using clsBuiness;
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
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace QCAuto
{
    public partial class frmPrice : Form
    {
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

        Excel.Application oApp;
        Excel.Workbooks oBooks;
        Excel.Workbook oBook;
        Excel.Worksheet oSheet;
        object oWebBrowser;
        string pass;
        public frmPrice(string password)
        {
            InitializeComponent();

            //clsAllnew BusinessHelp = new clsAllnew();

            //BusinessHelp.InputNewBaobiao_RawDatat();
            ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xlsx");

            tabControl1.TabPages[0].Parent = null;
            if (password == "123")
            {
                tabControl1.TabPages[0].Parent = null;

                toolStripButton10.Visible = false;


                toolStripButton11.Visible = false;
            }
            else
            {
                tabControl1.TabPages[1].Parent = null;

                toolStripButton1.Visible = false;

            }
            pass = password;


            Local_IP();
         //   ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xlsx");
            if (!File.Exists(ZFCEPath))
            {
                MessageBox.Show("没有找到此路径或此文件，请保证共享文件存在!");
                System.Environment.Exit(0);
                return;
            }

        }
        private void Local_IP()
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "IP.txt";
            string[] fileText = File.ReadAllLines(A_Path);
            if (fileText.Length > 0 && fileText[0] != null && fileText[0] != "")
            {
                if (fileText[0] != null && fileText[0] != "")
                    ZFCEPath = fileText[0];

            }
        }
        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            //  Inputexcel();

            Object refmissing = System.Reflection.Missing.Value;
            //this.axWebBrowser2.Navigate(ZFCEPath);
            axWebBrowser2.Navigate(ZFCEPath, ref refmissing, ref refmissing, ref refmissing, ref refmissing);
            //   axWebBrowser2.ExecWB(SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, ref refmissing, ref refmissing);
            //  this.webBrowser1.Navigate(strFileName);
            //    object axWebBrowser = this.webBrowser1.ActiveXInstance;




        }

        private void axWebBrowser2_NavigateComplete2(object sender, AxSHDocVw.DWebBrowserEvents2_NavigateComplete2Event e)
        {

            ///   return;

            object o = e.pDisp;
            oWebBrowser = e.pDisp;
            try
            {

                Object oDocument = o.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, o, null);
                Object oApplication = o.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oDocument, null);
                Excel.Application eApp = (Excel.Application)oApplication;
                eApp.UserControl = true;
                //Inputexcel(eApp);
                //textexcel();


                #region 方法2
                //Object refmissing = System.Reflection.Missing.Value;
                //object[] args = new object[4];
                //args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;
                //args[1] = SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER;
                //args[2] = refmissing;
                //args[3] = refmissing;

                //object axWebBrowser = this.webBrowser1.ActiveXInstance;

                //axWebBrowser.GetType().InvokeMember("ExecWB",
                //    BindingFlags.InvokeMethod, null, axWebBrowser, args);


                //object Application = axWebBrowser.GetType().InvokeMember("Document",
                //    BindingFlags.GetProperty, null, axWebBrowser, null);

                //Excel.Workbook wbb = (Excel.Workbook)oApplication;
                //Excel.ApplicationClass excel = wbb.Application as Excel.ApplicationClass;
                //Excel.Workbook wb = excel.Workbooks[1];
                //Excel.Worksheet ws = wb.Worksheets[1] as Excel.Worksheet;
                //ws.Cells.Font.Name = "Verdana";
                //ws.Cells.Font.Size = 14;
                //ws.Cells.Font.Bold = true;
                //Excel.Range range = ws.Cells;

                //Excel.Range oCell = range[10, 10] as Excel.Range;
                //oCell.Value2 = "你好";
                #endregion


                #region inster tx
                //object objBooks = eApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, eApp, null);

                ////添加一个新的Workbook
                //object objBook = objBooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, objBooks, null);
                ////获取Sheet集
                //object objSheets = objBook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, objBook, null);

                ////获取第一个Sheet对象
                //object[] Parameters = new Object[1] { 1 };
                //object objSheet = objSheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, objSheets, Parameters);

                //Parameters = new Object[2] { 1, 1 + 1 };
                //object objCells = objSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, objSheet, Parameters);
                ////向指定单元格填写内容值
                //Parameters = new Object[1] { "name" };
                //objCells.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objCells, Parameters);

                #endregion

                #region 一、首先简要回顾一下如何操作Excel表
                Workbooks workbooks = eApp.Workbooks;
                Excel.ApplicationClass excel = workbooks.Application as Excel.ApplicationClass;
                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)workbooks.get_Item(1);
                Excel.Workbook wb = excel.Workbooks[1];
                //_Workbook workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                int c = workbooks.Count;
                _Workbook workbook = workbooks.Add(ZFCEPath);
                Sheets sheets = workbook.Worksheets;

                _Worksheet worksheet = (_Worksheet)sheets.get_Item(1);
                Range range1 = worksheet.get_Range("A1", Missing.Value);
                const int nCells = 2345;
                range1.Value2 = nCells;

                #endregion





                ExcelExit();

            }
            catch (Exception ex)
            {
                ExcelExit();

                throw;
            }
        }
        public void Inputexcel(Microsoft.Office.Interop.Excel.Application excelApp1)
        {

            try
            {

                string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xls");

                //需要换 成日期的导出表
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(ZFCEPath, Type.Missing, true, Type.Missing,
                    "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                excelApp.Visible = true;
                excelApp.ScreenUpdating = true;

                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range rng;
                //   rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 45]);
                rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                int rowCount = WS.UsedRange.Rows.Count - 1;
                object[,] o = new object[1, 1];
                o = (object[,])rng.Value2;



                Microsoft.Office.Interop.Excel.AllowEditRanges ranges = WS.Protection.AllowEditRanges;
                ranges.Add("Information", WS.Range["B2:E6"], Type.Missing);

                WS.Protect("123", true);

                clsCommHelp.CloseExcel(excelApp, analyWK);

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

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
            NAR(oSheet);
            if (oBook != null)
            {
                oBook.Close(false);
                NAR(oBook);
                NAR(oBooks);
                if (oApp != null)
                    oApp.Quit();
            }
            if (oApp != null)
                NAR(oApp);
            Debug.WriteLine("Sleeping...");
            // System.Threading.Thread.Sleep(5000);
            Debug.WriteLine("End Excel");
            webBrowser1.Stop();
            webBrowser1.Dispose();
            System.Environment.Exit(0);



        }
        private void NAR(Object o)
        {
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(o); }
            catch { }
            finally { o = null; }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                #region 1
                //object o = oWebBrowser;


                //Object oDocument = o.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, o, null);
                //Object oApplication = o.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oDocument, null);
                ////Excel.Application eApp = (Excel.Application)oApplication;
                ////eApp.UserControl = true;


                //Object refmissing = System.Reflection.Missing.Value;
                //object[] args = new object[4];
                //args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;
                //args[1] = SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER;
                //args[2] = refmissing;
                //args[3] = refmissing;


                //object axWebBrowser = this.webBrowser1.ActiveXInstance;

                //axWebBrowser.GetType().InvokeMember("ExecWB",
                //    BindingFlags.InvokeMethod, null, axWebBrowser, args);


                //object Application = axWebBrowser.GetType().InvokeMember("Document",
                //    BindingFlags.GetProperty, null, axWebBrowser, null);

                //Excel.Workbook wbb = (Excel.Workbook)oApplication;
                //Excel.ApplicationClass excel = wbb.Application as Excel.ApplicationClass;
                //Excel.Workbook wb = excel.Workbooks[1];
                //Excel.Worksheet ws = wb.Worksheets[1] as Excel.Worksheet;
                //ws.Cells.Font.Name = "Verdana";
                //ws.Cells.Font.Size = 14;
                //ws.Cells.Font.Bold = true;
                //Excel.Range range = ws.Cells;

                //Excel.Range oCell = range[10, 10] as Excel.Range;
                //oCell.Value2 = "你好"; 
                #endregion



                #region is good

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


                Microsoft.Office.Interop.Excel.Range rng;

                rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                int rowCount = WS.UsedRange.Rows.Count;
                object[,] o = new object[1, 1];
                o = (object[,])rng.Value2;

                //Excel.Application eApp = (Excel.Application)oApplication;
                //  clsCommHelp.CloseExcel(eApp, wbb);
                // for (int i = 2; i <= rowCount; i++)
                {
                    cls_sixzhuanjiagebiao_info temp = new cls_sixzhuanjiagebiao_info();

                    //temp.touxing_B = "";
                    //if (o[i, 2] != null)
                    //    temp.touxing_B = o[i, 2].ToString().Trim();// o[i, 2].ToString().Trim();
                    if (pass == "123")
                    {
                        Microsoft.Office.Interop.Excel.Range cloumnrng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 1]];
                        cloumnrng.EntireColumn.Hidden = false;


                        cloumnrng = WS.Range[WS.Cells[1, 2], WS.Cells[WS.UsedRange.Rows.Count, 2]];
                        cloumnrng.EntireColumn.Hidden = true;

                        cloumnrng = WS.Range[WS.Cells[1, 3], WS.Cells[WS.UsedRange.Rows.Count, 3]];
                        cloumnrng.EntireColumn.Hidden = true;

                        cloumnrng = WS.Range[WS.Cells[1, 4], WS.Cells[WS.UsedRange.Rows.Count, 4]];
                        cloumnrng.EntireColumn.Hidden = true;

                        cloumnrng = WS.Range[WS.Cells[1, 5], WS.Cells[WS.UsedRange.Rows.Count, 5]];
                        cloumnrng.EntireColumn.Hidden = true;



                        cloumnrng = WS.Range[WS.Cells[1, 6], WS.Cells[WS.UsedRange.Rows.Count, 6]];
                        cloumnrng.EntireColumn.Hidden = true;


                        cloumnrng = WS.Range[WS.Cells[1, 8], WS.Cells[WS.UsedRange.Rows.Count, 8]];
                        cloumnrng.EntireColumn.Hidden = true;


                        cloumnrng = WS.Range[WS.Cells[1, 9], WS.Cells[WS.UsedRange.Rows.Count, 9]];
                        cloumnrng.EntireColumn.Hidden = true;


                        cloumnrng = WS.Range[WS.Cells[1, 10], WS.Cells[WS.UsedRange.Rows.Count, 10]];
                        cloumnrng.EntireColumn.Hidden = true;


                        cloumnrng = WS.Range[WS.Cells[1, 11], WS.Cells[WS.UsedRange.Rows.Count, 11]];
                        cloumnrng.EntireColumn.Hidden = true;

                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel.Range cloumnrng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 1]];
                        cloumnrng.EntireColumn.Hidden = false;


                        cloumnrng = WS.Range[WS.Cells[1, 2], WS.Cells[WS.UsedRange.Rows.Count, 2]];
                        cloumnrng.EntireColumn.Hidden = false;

                        cloumnrng = WS.Range[WS.Cells[1, 3], WS.Cells[WS.UsedRange.Rows.Count, 3]];
                        cloumnrng.EntireColumn.Hidden = false;

                        cloumnrng = WS.Range[WS.Cells[1, 4], WS.Cells[WS.UsedRange.Rows.Count, 4]];
                        cloumnrng.EntireColumn.Hidden = false;

                        cloumnrng = WS.Range[WS.Cells[1, 5], WS.Cells[WS.UsedRange.Rows.Count, 5]];
                        cloumnrng.EntireColumn.Hidden = false;


                        cloumnrng = WS.Range[WS.Cells[1, 6], WS.Cells[WS.UsedRange.Rows.Count, 6]];
                        cloumnrng.EntireColumn.Hidden = false;


                        cloumnrng = WS.Range[WS.Cells[1, 8], WS.Cells[WS.UsedRange.Rows.Count, 8]];
                        cloumnrng.EntireColumn.Hidden = false;


                        cloumnrng = WS.Range[WS.Cells[1, 9], WS.Cells[WS.UsedRange.Rows.Count, 9]];
                        cloumnrng.EntireColumn.Hidden = false;


                        cloumnrng = WS.Range[WS.Cells[1, 10], WS.Cells[WS.UsedRange.Rows.Count, 10]];
                        cloumnrng.EntireColumn.Hidden = false;


                        cloumnrng = WS.Range[WS.Cells[1, 11], WS.Cells[WS.UsedRange.Rows.Count, 11]];
                        cloumnrng.EntireColumn.Hidden = false;

                    }
                    oBook = wbb;
                    //oApp = (Excel.Application)oApplication;
                    //  Excel.Application eApp = (Excel.Application)oApplication;

                    oSheet = WS;

                }
                #endregion

                //设置工作表保护密码，并指定不受密码保护（即允许用户编辑）的单元格区域 
                //WS.Protect("123", WS.Range["B2:AZ65535"], "");

                //WS.Protect("123", true);
                //WS.Protect("123", SheetProtectionType.All);
                //Microsoft.Office.Interop.Excel.AllowEditRanges ranges = WS.Protection.AllowEditRanges;
                //ranges.Add("Information", WS.Range["B2:E6"], Type.Missing);



                ////初始化一个工作簿并加载一个实例
                //Spire.Xls.Workbook book = new Spire.Xls.Workbook();
                //book.LoadFromFile("test.xlsx");

                ////获取工作簿中第一个工作表
                //Spire.Xls.Worksheet sheet = book.Worksheets[0];

                //sheet.AddAllowEditRange("AAA", sheet.Range["B2:E6"], "");
                //sheet.Protect("AAA", SheetProtectionType.All); 


            }
            catch (Exception ex)
            {

                toolStripLabel2.Text = "请关闭所有Excel，然后点击【强制刷新】按钮，如还出现异常再次点击【强制刷新】按钮";

                MessageBox.Show("请关闭所有Excel，然后点击【强制刷新】按钮");
                return;

                throw;
            }
        }

        private void frmPrice_FormClosing(object sender, FormClosingEventArgs e)
        {
            ExcelExit();


        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Object refmissing = System.Reflection.Missing.Value;
            axWebBrowser2.ExecWB(SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, ref refmissing, ref refmissing);

        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            toolStripLabel2.Text = "读取中，请耐心等待...(打开快慢受网络情况影响)";

            this.webBrowser1.Navigate(ZFCEPath);

            toolStripLabel2.Text = "读取完成,马上显示";

            return;


            try
            {


                InitialBackGroundWorker();
                Control.CheckForIllegalCrossThreadCalls = false;
                bgWorker.DoWork += new DoWorkEventHandler(CheckBankCharge);

                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {


                }
            }
            catch (Exception ex)
            {
                throw ex;
            }


            // Inputexcel(null);

            //  webBrowser1 = new WebBrowser();

            //Object refmissing = System.Reflection.Missing.Value;



            //object axWebBrowser = this.webBrowser1.ActiveXInstance; 
        }
        private void CheckBankCharge(object sender, DoWorkEventArgs e)
        {
            DateTime oldDate = DateTime.Now;

            this.webBrowser1.Navigate(ZFCEPath);


            DateTime FinishTime = DateTime.Now;  //   
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();


            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            ExcelExit();

            this.Close();

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否已经保存其他桌面Excel文件，, 继续 ?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {

            }
            else
                return;

            toolStripLabel2.Text = "刷新中,请稍等...";

            string folderpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClearTask.bat");

            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.WorkingDirectory = folderpath;
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = folderpath;
            p.Start();

            toolStripButton10_Click(null, EventArgs.Empty);

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            toolStripLabel1.Text = "读取中，请耐心等待...(打开快慢受网络情况影响)";

            List<cls_sixzhuanjiagebiao_info> MAPPINGResult = GetKEYnfo(ZFCEPath);

            toolStripLabel1.Text = "读取完成";
            dataGridView1.AutoGenerateColumns = false;

            dataGridView1.DataSource = MAPPINGResult;

            this.tabControl1.SelectedIndex = 1;


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
                            temp.zuzuangbizhong_C = o[i,3].ToString().Trim();

                        if (temp.zuzuangbizhong_C == null || temp.zuzuangbizhong_C == "")
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
                        if (o[i,8] != null)
                            temp.bizhong_J = o[i, 8].ToString().Trim();

                        temp.ganjia_K = "";
                        if (o[i, 9] != null)
                            temp.ganjia_K = String.Format("{0:N2}", Convert.ToDouble(Math.Round(Convert.ToDouble(o[i,9].ToString()), 2).ToString())); //o[i, 8].ToString().Trim();


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



    }
}
