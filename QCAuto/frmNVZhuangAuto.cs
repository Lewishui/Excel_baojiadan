using China_System.Common;
using SDZdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QCAuto
{
    public partial class frmNVZhuangAuto : Form
    {
        string pass;
        private string instertext;
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        List<cls_Sheet0home_info> Result;
        string strFileName;


        public frmNVZhuangAuto(string password)
        {
            InitializeComponent();

            pass = password;

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

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            tbox.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                instertext = tbox.FileName;


                textBox1.Text = instertext;

            }
            if (instertext == null || instertext == "")
                return;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.Filter = "xlsx|*.xlsx";
            strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }


            instertext = textBox1.Text;
            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(NEWReadclaimreportfromServer);

                bgWorker.RunWorkerAsync();

                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();

                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    if (Result != null && Result.Count > 0)
                        this.label2.Text = "总计/条：" + Result.Count.ToString();


                    string ZFCEPath = Path.GetDirectoryName(strFileName); ;
                    System.Diagnostics.Process.Start("explorer.exe", ZFCEPath);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("存在异常:" + ex.Message);
               
                return;

                throw ex;
            }

        }

        private void NEWReadclaimreportfromServer(object sender, DoWorkEventArgs e)
        {
            try
            {
                Result = new List<cls_Sheet0home_info>();
                //导入程序集
                DateTime oldDate = DateTime.Now;

                Result = ReadfindngFile(instertext);


                DateTime FinishTime = DateTime.Now;
                TimeSpan s1 = DateTime.Now - oldDate;
                string timei = s1.Minutes.ToString() + ":" + s1.Seconds.ToString();
                string Showtime = clsShowMessage.MSG_029 + timei.ToString();
                bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

            }
            catch (Exception ex)
            {
                MessageBox.Show("存在异常:" + ex);
                return;
                throw;
            }
        }
        public List<cls_Sheet0home_info> ReadfindngFile(string instertext)
        {
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;

            try
            {
                List<cls_Sheet0home_info> Result = new List<cls_Sheet0home_info>();
                List<cls_Sheet0home_info> filterResult = new List<cls_Sheet0home_info>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(instertext, Type.Missing, true, Type.Missing,
                    "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[2];
                Microsoft.Office.Interop.Excel.Range rng;
                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                int rowCount = WS.UsedRange.Rows.Count - 1;
                object[,] o = new object[1, 1];
                o = (object[,])rng.Value2;

                for (int i = 1; i <= rowCount; i++)
                {
                    bgWorker.ReportProgress(0, "读入数据中 筛选库  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_Sheet0home_info temp = new cls_Sheet0home_info();

                    #region 基础信息

                    temp.guanjianci_A = "";
                    if (o[i, 1] != null)
                        temp.guanjianci_A = o[i, 1].ToString().Trim();

                    if (temp.guanjianci_A == "" || temp.guanjianci_A == null)
                        continue;
                    temp.beizhu1 = "";
                    if (o[i, 1] != null)
                        temp.beizhu1 = "行-" + i.ToString().Trim();

                    #endregion

                    filterResult.Add(temp);
                }

                WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                rowCount = WS.UsedRange.Rows.Count - 1;
                o = new object[1, 1];
                o = (object[,])rng.Value2;

                for (int i = 1; i <= rowCount; i++)
                {
                    bgWorker.ReportProgress(0, "读入数据中1  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_Sheet0home_info temp = new cls_Sheet0home_info();

                    #region 基础信息

                    temp.guanjianci_A = "";
                    if (o[i, 1] != null)
                        temp.guanjianci_A = o[i, 1].ToString().Trim();

                    if (temp.guanjianci_A == "" || temp.guanjianci_A == null)
                        continue;

                    temp.sousuorenqi_B = "";
                    if (o[i, 2] != null)
                        temp.sousuorenqi_B = o[i, 2].ToString().Trim();

                    temp.zaixianshangpinshu_C = "";
                    if (o[i, 3] != null)
                        temp.zaixianshangpinshu_C = o[i, 3].ToString().Trim();

                    temp.zhifuzhuanhuanlv_D = "";
                    if (o[i, 4] != null)
                        temp.zhifuzhuanhuanlv_D = o[i, 4].ToString().Trim();

                    temp.dianjilv_E = "";
                    if (o[i, 5] != null)
                        temp.dianjilv_E = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp.shangchengdianjizhanbi_F = "";
                    if (o[i, 6] != null)
                        temp.shangchengdianjizhanbi_F = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                    temp.sousuoci_G = "";
                    if (o[i, 7] != null)
                        temp.sousuoci_G = o[i, 7].ToString().Trim();
                    temp.anlanqi_H = "";
                    if (o[i, 8] != null)
                        temp.anlanqi_H = o[i, 8].ToString().Trim();


                    #endregion
                    if (temp.guanjianci_A == "viv0官方旗舰店手机")
                    { 
                    
                    
                    }
                 
                    List<cls_Sheet0home_info> filtered = filterResult.FindAll(s => temp.guanjianci_A.Contains(s.guanjianci_A));
                    if (filtered.Count > 0)
                    {
                        temp.beizhu1 = "无效" ;
                        WS.Cells[i+1, 12] = temp.beizhu1;
                    }
                    Result.Add(temp);
                }

                if (this.radioButton2.Checked == true)
                {
                    //excelApp.Visible = true;
                    //excelApp.ScreenUpdating = true;
                    List<cls_Sheet0home_info> DeleteList = Result.FindAll(s => s.beizhu1 != null && s.beizhu1.Contains("无效"));
                    string[] s11 = new string[DeleteList.Count];
                    for (int i = 0; i < DeleteList.Count; i++)
                        s11[i] = DeleteList[i].beizhu1;

                    WS.get_Range(WS.Cells[1, 1], WS.Cells[1, 10]).AutoFilter(9, s11, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                    WS.get_Range(WS.Cells[2, 1], WS.Cells[50000, 20]).EntireRow.Delete(0);

                    WS.get_Range(WS.Cells[1, 1], WS.Cells[1, 10]).AutoFilter(9, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, false);
                }
                //excelApp.Visible = true;
                //excelApp.ScreenUpdating = true;

                analyWK.SaveAs(strFileName, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                excelApp.DisplayAlerts = false;

                clsCommHelp.CloseExcel(excelApp, analyWK);

                return Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show("表格存在异常,请参照原始表格格式修改:" + ex.Message);
                return null;

                throw ex;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();


        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }


    }
}
