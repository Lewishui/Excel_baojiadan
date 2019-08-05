using China_System.Common;
using SDZdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.IO;
namespace QCAuto
{
    public partial class frmGPS改造项目基本信息 : Form
    {
        private string instertext;
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;

        List<cls_GPS_info> Result;
        List<cls_gaizaoqianjinggao_info> gaizaoqianResult;
        public string folderpath;
        List<cls_gaizaoHOUjinggao_info> gaizaoHOUResult;
        List<cls_zongqingdan_zhibiao_info> zongqingdan_zhibiaoResult;
        List<String> folder_list = new List<String>();
        string pass;

        public frmGPS改造项目基本信息(string password)
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
        private void NEWReadclaimreportfromServer(object sender, DoWorkEventArgs e)
        {
            Result = new List<cls_GPS_info>();
           folder_list = new List<String>();
            //导入程序集
            DateTime oldDate = DateTime.Now;
            if (folderpath != null && folderpath != "")
                folder_list = director(folderpath);

            Result = ReadfindngFile(instertext);
            //获取文件夹的所有路径
            for (int i = 0; i < Result.Count; i++)
            {
                bgWorker.ReportProgress(0, "生成数据中  :  " + Result[i].zhandianmingcheng+"> " + i.ToString() + "/" + Result.Count.ToString());
                  
                var ex = folder_list.FindAll(v => v .Contains( Result[i].zhandianmingcheng));


                List<cls_gaizaoqianjinggao_info> cloumnlist1 = gaizaoqianResult.FindAll(s => s.guzangyuan != null && Result[i].zhandianmingcheng.Contains(s.guzangyuan));
                List<cls_gaizaoHOUjinggao_info> cloumnlist2 = gaizaoHOUResult.FindAll(s => s.guzangyuan != null && Result[i].zhandianmingcheng.Contains(s.guzangyuan));
                List<cls_zongqingdan_zhibiao_info> cloumnlist3 = zongqingdan_zhibiaoResult.FindAll(s => s.jizhanmingcheng != null && Result[i].zhandianmingcheng.Contains(s.jizhanmingcheng));

                string image1 = ex.Find(v => v.Contains("外"));
                string image2 = ex.Find(v => v.Contains("内"));
                string image3 = ex.Find(v => v.Contains("前"));
                string image4 = ex.Find(v => v.Contains("后"));
                FindclaimreportData_ByMu(Result[i], cloumnlist1, cloumnlist2, cloumnlist3, image1, image2, image3, image4);
            }
            DateTime FinishTime = DateTime.Now;
            TimeSpan s1 = DateTime.Now - oldDate;
            string timei = s1.Minutes.ToString() + ":" + s1.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

        }
        public List<String> director(string dirs)
        {
         
            //绑定到指定的文件夹目录
            DirectoryInfo dir = new DirectoryInfo(dirs);
            //检索表示当前目录的文件和子目录
            FileSystemInfo[] fsinfos = dir.GetFileSystemInfos();
            //遍历检索的文件和子目录
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                //判断是否为空文件夹　　
                if (fsinfo is DirectoryInfo)
                {
                    //递归调用
                    director(fsinfo.FullName);
                }
                else
                {
                    Console.WriteLine(fsinfo.FullName);
                    //将得到的文件全路径放入到集合中
                    folder_list.Add(fsinfo.FullName);
                }
            }
            return folder_list;

        }
        public List<cls_GPS_info> FindclaimreportData_ByMu(cls_GPS_info Item, List<cls_gaizaoqianjinggao_info> cloumnlist1, List<cls_gaizaoHOUjinggao_info> cloumnlist2, List<cls_zongqingdan_zhibiao_info> cloumnlist3, string image1, string image2, string image3, string image4)
        {
            try
            {
                object miss = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Word.Application appWord = null;//应用程序
                object Nothing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.DocumentClass doc = null;//文档

                appWord = new Microsoft.Office.Interop.Word.Application();

             //   appWord.Visible = true;

                object objTrue = true;

                object objFalse = false;


                object objTemplate = AppDomain.CurrentDomain.BaseDirectory + "GPS改造项目\\tel.docm";

                object objDocType = WdDocumentType.wdTypeDocument;

                doc = (DocumentClass)appWord.Documents.Add(ref objTemplate, ref objFalse, ref objDocType, ref objTrue);
               // appWord.Visible = true;
                object obDD_Name = "add1";//站点名称
                doc.Bookmarks.get_Item(ref obDD_Name).Range.Text = Item.zhandianmingcheng; // 

                //站点基本信息
                object dishi = "dishi";//地市
                doc.Bookmarks.get_Item(ref dishi).Range.Text = Item.dishi; //姓名

                object quyu = "quyu";//区域
                doc.Bookmarks.get_Item(ref quyu).Range.Text = Item.quyu;


                object changjia = "changjia";//厂家
                doc.Bookmarks.get_Item(ref changjia).Range.Text = Item.changjia;

                object ruchangshijian = "ruchangshijian";//入场时间
                doc.Bookmarks.get_Item(ref ruchangshijian).Range.Text = Item.ruchangshijian;

                object xianchangongchengshi = "xianchangongchengshi";//现场工程师
                doc.Bookmarks.get_Item(ref xianchangongchengshi).Range.Text = Item.xianchanggongchengsi;

                object lianxidianhua = "lianxidianhua";//联系电话
                doc.Bookmarks.get_Item(ref lianxidianhua).Range.Text = Item.lianxidianhua;

                object zhandianID = "zhandianID";//站点ID
                doc.Bookmarks.get_Item(ref zhandianID).Range.Text = Item.zhandianID;

                object zhandianmingcheng = "zhandianmingcheng";//站点名称
                doc.Bookmarks.get_Item(ref zhandianmingcheng).Range.Text = Item.zhandianmingcheng;

                object zhandianweidu = "zhandianweidu";//站点维度
                doc.Bookmarks.get_Item(ref zhandianweidu).Range.Text = Item.zhandianweidu;

                object zhandianjingdu = "zhandianjingdu";//站点经度
                doc.Bookmarks.get_Item(ref zhandianjingdu).Range.Text = Item.zhandianjingdu;
                object zhandiandizhi = "zhandiandizhi";//站点地址
                doc.Bookmarks.get_Item(ref zhandiandizhi).Range.Text = Item.zhandiandizhi;

                object biaotiriqi = "biaotiriqi";//开头的日期
                doc.Bookmarks.get_Item(ref biaotiriqi).Range.Text = Item.ruchangshijian;


                

                #region 图片
                string picfileName = @"D:\Devlop\VBA_tool\杂七杂八\admin651813235\新作坡村T\新作坡村T\外部.jpg";
                picfileName = image1;
                foreach (Bookmark bk in doc.Bookmarks)
                {
                    //图片1：基站外部环境
                    if (bk.Name == "jizhanwaibuhuanjingPIC" && File.Exists(picfileName))
                    {
                        bk.Select();
                        Selection sel = appWord.Selection;
                        //sel.InlineShapes.AddPicture(ZFCEPath);

                        object Anchor = appWord.Selection.Range;

                        object LinkToFile = false;
                        object SaveWithDocument = true;
                        //设置图片位置
                        appWord.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        InlineShape inlineShape = appWord.ActiveDocument.InlineShapes.AddPicture(picfileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);

                        inlineShape.Width = 124; // 图片宽度   
                        inlineShape.Height = 157; // 图片高度  

                    }
                    string picfileName2 = @"D:\Devlop\VBA_tool\杂七杂八\admin651813235\新作坡村T\新作坡村T\外部.jpg";
                    picfileName2 = image2;
                    //图片2：基站内部环境
                    if (bk.Name == "jizhanneibuhuanjingPIC" && File.Exists(picfileName2))
                    {

                        bk.Select();
                        Selection sel = appWord.Selection;
                        //sel.InlineShapes.AddPicture(ZFCEPath);

                        object Anchor = appWord.Selection.Range;

                        object LinkToFile = false;
                        object SaveWithDocument = true;
                        //设置图片位置
                        appWord.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        InlineShape inlineShape = appWord.ActiveDocument.InlineShapes.AddPicture(picfileName2, ref LinkToFile, ref SaveWithDocument, ref Anchor);

                        inlineShape.Width = 124; // 图片宽度   
                        inlineShape.Height = 157; // 图片高度  

                    }
                    //图片3：基站改造前图片
                    string picfileName3 = @"D:\Devlop\VBA_tool\杂七杂八\admin651813235\新作坡村T\新作坡村T\外部.jpg";
                    picfileName3 = image3;
                    if (bk.Name == "jizhangaizhaoqiantipianPIC" && File.Exists(picfileName3))
                    {

                        bk.Select();
                        Selection sel = appWord.Selection;
                        //sel.InlineShapes.AddPicture(ZFCEPath);

                        object Anchor = appWord.Selection.Range;

                        object LinkToFile = false;
                        object SaveWithDocument = true;
                        //设置图片位置
                        appWord.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        InlineShape inlineShape = appWord.ActiveDocument.InlineShapes.AddPicture(picfileName3, ref LinkToFile, ref SaveWithDocument, ref Anchor);

                        inlineShape.Width = 124; // 图片宽度   
                        inlineShape.Height = 157; // 图片高度  

                    }
                    //图片四：基站改造后图片
                    string picfileName4 = @"D:\Devlop\VBA_tool\杂七杂八\admin651813235\新作坡村T\新作坡村T\外部.jpg";
                    picfileName4 = image4;
                    if (bk.Name == "jizhangaizhaoHOUtipianPIC" && File.Exists(picfileName4))
                    {

                        bk.Select();
                        Selection sel = appWord.Selection;
                        //sel.InlineShapes.AddPicture(ZFCEPath);

                        object Anchor = appWord.Selection.Range;

                        object LinkToFile = false;
                        object SaveWithDocument = true;
                        //设置图片位置
                        appWord.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        InlineShape inlineShape = appWord.ActiveDocument.InlineShapes.AddPicture(picfileName4, ref LinkToFile, ref SaveWithDocument, ref Anchor);

                        inlineShape.Width = 124; // 图片宽度   
                        inlineShape.Height = 157; // 图片高度  

                    }
                }
                #endregion
                #region 改造前告警信息
                // 
                int gzqjgxx = 1;
                foreach (cls_gaizaoqianjinggao_info oyem in cloumnlist1)
                {
                    if (gzqjgxx > 3)
                        break;

                    //if (gzqjgxx == 1)
                    {
                        object xuliehao = "xuliehao" + gzqjgxx.ToString();//序列号
                        doc.Bookmarks.get_Item(ref xuliehao).Range.Text = oyem.xuliehao;
                        object wangyuanleixing = "wangyuanleixing" + gzqjgxx.ToString();//网元类型
                        doc.Bookmarks.get_Item(ref wangyuanleixing).Range.Text = oyem.wangyuanleixing;
                        object guzangyuan = "guzangyuan" + gzqjgxx.ToString();//故障源
                        doc.Bookmarks.get_Item(ref guzangyuan).Range.Text = oyem.guzangyuan;


                        object chanshengshijian = "chanshengshijian" + gzqjgxx.ToString();//产生时间
                        doc.Bookmarks.get_Item(ref chanshengshijian).Range.Text = oyem.chanshengshijian;
                        object yuanyin = "yuanyin" + gzqjgxx.ToString();//原因号
                        doc.Bookmarks.get_Item(ref yuanyin).Range.Text = oyem.yuanyin;
                        object yuanyinmiaoshu = "yuanyinmiaoshu" + gzqjgxx.ToString();//原因描述
                        doc.Bookmarks.get_Item(ref yuanyinmiaoshu).Range.Text = oyem.yuanyinmiaoshu;
                        object gaojingbianhao = "gaojingbianhao" + gzqjgxx.ToString();//告警编号
                        doc.Bookmarks.get_Item(ref gaojingbianhao).Range.Text = oyem.gaojingbianhao;
                        object gaojingmingcheng = "gaojingmingcheng" + gzqjgxx.ToString();//告警名称
                        doc.Bookmarks.get_Item(ref gaojingmingcheng).Range.Text = oyem.gaojingmingcheng;
                    }
                    gzqjgxx++;


                }
                #endregion

                #region 改造后告警信息
                //gzqjgxx = 1;
                foreach (cls_gaizaoHOUjinggao_info oyem in cloumnlist2)
                {
                    if (gzqjgxx > 6)
                        break;

                    //if (gzqjgxx == 1)
                    {
                        object xuliehao = "xuliehao" + gzqjgxx.ToString();//序列号
                        doc.Bookmarks.get_Item(ref xuliehao).Range.Text = oyem.xuliehao;
                        object wangyuanleixing = "wangyuanleixing" + gzqjgxx.ToString();//网元类型
                        doc.Bookmarks.get_Item(ref wangyuanleixing).Range.Text = oyem.wangyuanleixing;
                        object guzangyuan = "guzangyuan" + gzqjgxx.ToString();//故障源
                        doc.Bookmarks.get_Item(ref guzangyuan).Range.Text = oyem.guzangyuan;


                        object chanshengshijian = "chanshengshijian" + gzqjgxx.ToString();//产生时间
                        doc.Bookmarks.get_Item(ref chanshengshijian).Range.Text = oyem.chanshengshijian;
                        object yuanyin = "yuanyin" + gzqjgxx.ToString();//原因号
                        doc.Bookmarks.get_Item(ref yuanyin).Range.Text = oyem.yuanyin;
                        object yuanyinmiaoshu = "yuanyinmiaoshu" + gzqjgxx.ToString();//原因描述
                        doc.Bookmarks.get_Item(ref yuanyinmiaoshu).Range.Text = oyem.yuanyinmiaoshu;
                        object gaojingbianhao = "gaojingbianhao" + gzqjgxx.ToString();//告警编号
                        doc.Bookmarks.get_Item(ref gaojingbianhao).Range.Text = oyem.gaojingbianhao;
                        object gaojingmingcheng = "gaojingmingcheng" + gzqjgxx.ToString();//告警名称
                        doc.Bookmarks.get_Item(ref gaojingmingcheng).Range.Text = oyem.gaojingmingcheng;
                    }
                    gzqjgxx++;


                }
                #endregion

                #region 站点指标对比
                gzqjgxx = 1;
                foreach (cls_zongqingdan_zhibiao_info oyem in cloumnlist3)
                {
                    if (gzqjgxx > 2)
                        break;

                    if (oyem.shijian == "改造前")
                    {
                        object LTE_C = "LTE_C" + gzqjgxx.ToString();//LTE_RRC连接建立成功率[单位:%]
                        doc.Bookmarks.get_Item(ref LTE_C).Range.Text = oyem.LTE_C;
                        object LTE_D = "LTE_D" + gzqjgxx.ToString();//LTE_无线掉线率[单位:%]
                        doc.Bookmarks.get_Item(ref LTE_D).Range.Text = oyem.LTE_D;
                        object LTE_E = "LTE_E" + gzqjgxx.ToString();//LTE_切换成功率[单位:%]
                        doc.Bookmarks.get_Item(ref LTE_E).Range.Text = oyem.LTE_E;

                    }
                    if (oyem.shijian == "改造后")
                    {
                        object LTE_C = "LTE_C" + gzqjgxx.ToString();//LTE_RRC连接建立成功率[单位:%]
                        doc.Bookmarks.get_Item(ref LTE_C).Range.Text = oyem.LTE_C;
                        object LTE_D = "LTE_D" + gzqjgxx.ToString();//LTE_无线掉线率[单位:%]
                        doc.Bookmarks.get_Item(ref LTE_D).Range.Text = oyem.LTE_D;
                        object LTE_E = "LTE_E" + gzqjgxx.ToString();//LTE_切换成功率[单位:%]
                        doc.Bookmarks.get_Item(ref LTE_E).Range.Text = oyem.LTE_E;

                    }
                    gzqjgxx++;


                }
                #endregion

                object obtable = "obtable";
                string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
                object FileName = Path.Combine(file, Item.zhandianmingcheng+ DateTime.Now.ToString("yyyyMMdd-ss") + ".doc");
                //new
                FileName = Path.Combine(file,Item.ruchangshijian+"华为3G及4G共模基站GPS调整报告("+ Item.zhandianmingcheng + ")" + ".doc");

                doc.SaveAs(ref FileName, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                object missingValue = Type.Missing;

                object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;

                doc.Close(ref doNotSaveChanges, ref missingValue, ref missingValue);

                appWord.Application.Quit(ref miss, ref miss, ref miss);

                doc = null;

                appWord = null;
                return null;


            }
            catch (Exception ex)
            {
                MessageBox.Show("数据模板异常" + ex.Message);
                return null;
                throw;
            }

        }

        public List<cls_GPS_info> ReadfindngFile(string instertext)
        {

            try
            {
                List<cls_GPS_info> Result = new List<cls_GPS_info>();
                gaizaoqianResult = new List<cls_gaizaoqianjinggao_info>();
                gaizaoHOUResult = new List<cls_gaizaoHOUjinggao_info>();
                zongqingdan_zhibiaoResult = new List<cls_zongqingdan_zhibiao_info>();


                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(instertext, Type.Missing, true, Type.Missing,
                    "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["站点基本信息"];
                Microsoft.Office.Interop.Excel.Range rng;
                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                int rowCount = WS.UsedRange.Rows.Count - 1;
                object[,] o = new object[1, 1];
                o = (object[,])rng.Value2;
                //clsCommHelp.CloseExcel(excelApp, analyWK);

                for (int i = 2; i <= rowCount; i++)
                {
                    bgWorker.ReportProgress(0, "读入数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_GPS_info temp = new cls_GPS_info();

                    #region 基础信息

                    temp.duiying = "";
                    if (o[i, 1] != null)
                        temp.duiying = o[i, 1].ToString().Trim();


                    temp.zhandianmingcheng = "";
                    if (o[i, 2] != null)
                        temp.zhandianmingcheng = o[i, 2].ToString().Trim();

                    temp.dishi = "";
                    if (o[i, 3] != null)
                        temp.dishi = o[i, 3].ToString().Trim();

                    temp.quyu = "";
                    if (o[i, 4] != null)
                        temp.quyu = o[i, 4].ToString().Trim();
                    if (temp.quyu == "" || temp.quyu == null)
                        continue;

                    temp.changjia = "";
                    if (o[i, 5] != null)
                        temp.changjia = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp.ruchangshijian = "";
                    if (o[i, 6] != null)
                    {
                        temp.ruchangshijian = clsCommHelp.objToDateTime1(o[i, 6]);

                        DateTime dt3 = Convert.ToDateTime(temp.ruchangshijian);
                        temp.ruchangshijian = dt3.ToString("yyyy年MM月dd日");

                    }
           
                    temp.xianchanggongchengsi = "";
                    if (o[i, 7] != null)
                        temp.xianchanggongchengsi = o[i, 7].ToString().Trim();
                    temp.lianxidianhua = "";
                    if (o[i, 8] != null)
                        temp.lianxidianhua = o[i, 8].ToString().Trim();

                    temp.zhandianID = "";
                    if (o[i, 9] != null)
                        temp.zhandianID = o[i, 9].ToString().Trim();

                    temp.zhandianjingdu = "";
                    if (o[i, 10] != null)
                        temp.zhandianjingdu = o[i, 10].ToString().Trim();

                    temp.zhandianweidu = "";
                    if (o[i, 11] != null)
                        temp.zhandianweidu = o[i, 11].ToString().Trim();
                    temp.zhandiandizhi = "";
                    if (o[i, 12] != null)
                        temp.zhandiandizhi = o[i, 12].ToString().Trim();



                    #endregion

                    Result.Add(temp);
                }

                //
                #region 告警调整前

                WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["告警调整前"];

                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                rowCount = WS.UsedRange.Rows.Count - 1;
                o = new object[1, 1];
                o = (object[,])rng.Value2;
                //clsCommHelp.CloseExcel(excelApp, analyWK);

                for (int i = 2; i <= rowCount; i++)
                {
                    bgWorker.ReportProgress(0, "读入告警调整前数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_gaizaoqianjinggao_info temp = new cls_gaizaoqianjinggao_info();

                    #region 基础信息

                    temp.xuliehao = "";
                    if (o[i, 1] != null)
                        temp.xuliehao = o[i, 1].ToString().Trim();


                    temp.wangyuanleixing = "";
                    if (o[i, 2] != null)
                        temp.wangyuanleixing = o[i, 2].ToString().Trim();

                    temp.guzangyuan = "";
                    if (o[i, 3] != null)
                        temp.guzangyuan = o[i, 3].ToString().Trim();

                    temp.chanshengshijian = "";
                    if (o[i, 4] != null)
                        temp.chanshengshijian = o[i, 4].ToString().Trim();
                    if (temp.chanshengshijian == "" || temp.chanshengshijian == null)
                        continue;

                    temp.yuanyin = "";
                    if (o[i, 5] != null)
                        temp.yuanyin = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp.yuanyinmiaoshu = "";
                    if (o[i, 6] != null)
                        temp.yuanyinmiaoshu = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                    temp.gaojingbianhao = "";
                    if (o[i, 7] != null)
                        temp.gaojingbianhao = o[i, 7].ToString().Trim();
                    temp.gaojingmingcheng = "";
                    if (o[i, 8] != null)
                        temp.gaojingmingcheng = o[i, 8].ToString().Trim();



                    #endregion

                    gaizaoqianResult.Add(temp);
                }

                #endregion
                #region 告警调整后

                WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["告警调整后"];

                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                rowCount = WS.UsedRange.Rows.Count - 1;
                o = new object[1, 1];
                o = (object[,])rng.Value2;
                //clsCommHelp.CloseExcel(excelApp, analyWK);

                for (int i = 2; i <= rowCount; i++)
                {
                    bgWorker.ReportProgress(0, "读入告警调整后数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_gaizaoHOUjinggao_info temp = new cls_gaizaoHOUjinggao_info();

                    #region 基础信息

                    temp.xuliehao = "";
                    if (o[i, 1] != null)
                        temp.xuliehao = o[i, 1].ToString().Trim();


                    temp.wangyuanleixing = "";
                    if (o[i, 2] != null)
                        temp.wangyuanleixing = o[i, 2].ToString().Trim();

                    temp.guzangyuan = "";
                    if (o[i, 3] != null)
                        temp.guzangyuan = o[i, 3].ToString().Trim();

                    temp.chanshengshijian = "";
                    if (o[i, 4] != null)
                        temp.chanshengshijian = o[i, 4].ToString().Trim();
                    if (temp.chanshengshijian == "" || temp.chanshengshijian == null)
                        continue;

                    temp.yuanyin = "";
                    if (o[i, 5] != null)
                        temp.yuanyin = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp.yuanyinmiaoshu = "";
                    if (o[i, 6] != null)
                        temp.yuanyinmiaoshu = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                    temp.gaojingbianhao = "";
                    if (o[i, 7] != null)
                        temp.gaojingbianhao = o[i, 7].ToString().Trim();
                    temp.gaojingmingcheng = "";
                    if (o[i, 8] != null)
                        temp.gaojingmingcheng = o[i, 8].ToString().Trim();



                    #endregion

                    gaizaoHOUResult.Add(temp);
                }
                #endregion
                #region 总清单（指标）

                WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["总清单指标"];

                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                rowCount = WS.UsedRange.Rows.Count - 1;
                o = new object[1, 1];
                o = (object[,])rng.Value2;
                clsCommHelp.CloseExcel(excelApp, analyWK);

                for (int i = 2; i <= rowCount; i++)
                {
                    bgWorker.ReportProgress(0, "读入总清单指标数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_zongqingdan_zhibiao_info temp = new cls_zongqingdan_zhibiao_info();

                    #region 基础信息

                    temp.jizhanmingcheng = "";
                    if (o[i, 1] != null)
                        temp.jizhanmingcheng = o[i, 1].ToString().Trim();


                    temp.shijian = "";
                    if (o[i, 2] != null)
                        temp.shijian = o[i, 2].ToString().Trim();

                    temp.LTE_C = "";
                    if (o[i, 3] != null)
                        temp.LTE_C = o[i, 3].ToString().Trim();

                    temp.LTE_D = "";
                    if (o[i, 4] != null)
                        temp.LTE_D = o[i, 4].ToString().Trim();
                    if (temp.LTE_D == "" || temp.LTE_D == null)
                        continue;

                    temp.LTE_E = "";
                    if (o[i, 5] != null)
                        temp.LTE_E = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);

                    #endregion

                    zongqingdan_zhibiaoResult.Add(temp);
                }
                #endregion
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
            folderpath = textBox2.Text;
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



                    string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results"), "");
                    System.Diagnostics.Process.Start("explorer.exe", ZFCEPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("异常：" + ex);

                return ;

                throw ex;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择图片所在总文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                folderpath = dialog.SelectedPath;
                textBox2.Text = dialog.SelectedPath;
                folderpath = textBox2.Text;

            }
            else
                return;
        }

    }
}
