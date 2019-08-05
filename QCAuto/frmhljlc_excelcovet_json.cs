using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.IO;
using China_System.Common;
using SDZdb;
namespace QCAuto
{
    public partial class frmhljlc_excelcovet_json : Form
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
        List<cls_xiangmujihuazongbiao_info> Result;
        string strFileName;

        public frmhljlc_excelcovet_json(string password)
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
        protected string ObjToJSON(object obj)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
            Stream stream = new MemoryStream();
            serializer.WriteObject(stream, obj);
            stream.Position = 0;
            StreamReader streamReader = new StreamReader(stream);
            return streamReader.ReadToEnd();
        }

        public List<cls_xiangmujihuazongbiao_info> ReadfindngFile(string instertext)
        {

            try
            {
                List<cls_xiangmujihuazongbiao_info> Result = new List<cls_xiangmujihuazongbiao_info>();


                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(instertext, Type.Missing, true, Type.Missing,
                    "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range rng;
                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                int rowCount = WS.UsedRange.Rows.Count - 1;
                object[,] o = new object[1, 1];
                o = (object[,])rng.Value2;
                clsCommHelp.CloseExcel(excelApp, analyWK);

                for (int i = 1; i <= rowCount; i++)
                {
                    bgWorker.ReportProgress(0, "读入数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_xiangmujihuazongbiao_info temp = new cls_xiangmujihuazongbiao_info();

                    #region 基础信息

                    temp.xuhao_A = "";
                    if (o[i, 1] != null)
                        temp.xuhao_A = o[i, 1].ToString().Trim();


                    temp.tiaomaneirong_B = "";
                    if (o[i, 2] != null)
                        temp.tiaomaneirong_B = o[i, 2].ToString().Trim();

                    temp.tuhao_C = "";
                    if (o[i, 3] != null)
                        temp.tuhao_C = o[i, 3].ToString().Trim();

                    temp.mingcheng_D = "";
                    if (o[i, 4] != null)
                        temp.mingcheng_D = o[i, 4].ToString().Trim();
                    if (temp.mingcheng_D == "" || temp.mingcheng_D == null)
                        continue;

                    temp.caizhi_E = "";
                    if (o[i, 5] != null)
                        temp.caizhi_E = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp.shuliang_F = "";
                    if (o[i, 6] != null)
                        temp.shuliang_F = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                    temp.danwei_G = "";
                    if (o[i, 7] != null)
                        temp.danwei_G = o[i, 7].ToString().Trim();
                    temp.taoshu_H = "";
                    if (o[i, 8] != null)
                        temp.taoshu_H = o[i, 8].ToString().Trim();

                    temp.xiangmujiaoqi_I = "";
                    if (o[i, 9] != null)
                        temp.xiangmujiaoqi_I = o[i, 9].ToString().Trim();

                    temp.zongshuliang_J = "";
                    if (o[i, 10] != null)
                        temp.zongshuliang_J = o[i, 10].ToString().Trim();

                    temp.wuliuzhouqi_K = "";
                    if (o[i, 11] != null)
                        temp.wuliuzhouqi_K = o[i, 11].ToString().Trim();
                    temp.zhuangpeizhouqi_L = "";
                    if (o[i, 12] != null)
                        temp.zhuangpeizhouqi_L = o[i, 12].ToString().Trim();

                    temp.lingjianchengpinzhouqi_M = "";
                    if (o[i, 13] != null)
                        temp.lingjianchengpinzhouqi_M = clsCommHelp.objToDateTime(o[i, 13]);


                    temp.shifouxuyao_N = "";
                    if (o[i, 14] != null)
                        temp.shifouxuyao_N = o[i, 14].ToString().Trim();


                    temp.bianmianchulizhouqi_O = "";
                    if (o[i, 15] != null)
                        temp.bianmianchulizhouqi_O = o[i, 15].ToString().Trim();


                    temp.lingjianbanchengpinzhouqi_P = "";
                    if (o[i, 16] != null)
                        temp.lingjianbanchengpinzhouqi_P = clsCommHelp.objToDateTime(o[i, 16]);


                    temp.beizhu_Q = "";
                    if (o[i, 17] != null)
                        temp.beizhu_Q = o[i, 17].ToString().Trim();

                    temp.genchuineirong_R = "";
                    if (o[i, 18] != null)
                        temp.genchuineirong_R = o[i, 18].ToString().Trim();


                    temp.genchuijiedian_S = "";
                    if (o[i, 19] != null)
                        temp.genchuijiedian_S = clsCommHelp.objToDateTime(o[i, 19]);


                    temp.xiatushijian_T = "";
                    if (o[i, 20] != null)
                        temp.xiatushijian_T = clsCommHelp.objToDateTime(o[i, 20]);


                    temp.xiaruriqi_U = "";
                    if (o[i, 21] != null)
                        temp.xiaruriqi_U = o[i, 21].ToString().Trim();


                    temp.xiangmubiaohao_V = "";
                    if (o[i, 22] != null)
                        temp.xiangmubiaohao_V = o[i, 22].ToString().Trim();


                    temp.tuhao1_W = "";
                    if (o[i, 23] != null)
                        temp.tuhao1_W = o[i, 23].ToString().Trim();

                    #endregion

                    Result.Add(temp);
                }



                return Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show("表格存在异常,请参照原始表格格式修改:" + ex.Message);

                throw ex;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".txt";
            saveFileDialog.Filter = "txt|*.txt";
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

                }
            }
            catch (Exception ex)
            {
                return;

                throw ex;
            }
 

        }
        private void NEWReadclaimreportfromServer(object sender, DoWorkEventArgs e)
        {
            Result = new List<cls_xiangmujihuazongbiao_info>();
            //导入程序集
            DateTime oldDate = DateTime.Now;

            Result = ReadfindngFile(instertext);

            string str = ObjToJSON(Result);

            str = "{  " + "\"" + "imgListData" + "\"" + ": " + str + "    }";

            StreamWriter sw = new StreamWriter(strFileName);
            sw.WriteLine(str);
            sw.Flush();
            sw.Close();

            DateTime FinishTime = DateTime.Now;
            TimeSpan s1 = DateTime.Now - oldDate;
            string timei = s1.Minutes.ToString() + ":" + s1.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();

        }

    }
}
