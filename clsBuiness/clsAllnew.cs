using China_System.Common;
using ISR_System;
using SDZdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Caching;
using System.Windows.Forms;

namespace clsBuiness
{
    public enum ProcessStatus
    {
        初始化,
        登录界面,
        确认YES,
        第一页面,
        第二页面,
        Filter下拉,
        关闭页面,
        结束页面

    }
    public class clsAllnew
    {

        string newsth;
        public BackgroundWorker bgWorker1;
        private ProcessStatus isrun = ProcessStatus.初始化;
        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }
        public ToolStripStatusLabel tsStatusLabel2 { get; set; }
        private DateTime StopTime;
        List<clsuserinfo> userinfo_webResult;
        public string ConStr;
        public string ConStrPIC;
        private System.Windows.Forms.PictureBox picboxPhoto;
        string mdbpath2_Ctirx = AppDomain.CurrentDomain.BaseDirectory + "\\Hello.jpg";//记录 Status  click 和选择哪个服务器
        public string servename;
        private WbBlockNewUrl MyWebBrower;
        private Form viewForm;
        #region dll

        public static Boolean IsConnected = false;
        public static Boolean IsAuthenticate = false;
        public static Boolean IsRead_Content = false;
        public static int Port = 0;
        public static int ComPort = 0;
        public const int cbDataSize = 128;
        public const int GphotoSize = 256 * 1024;

        [DllImport("termb.dll")]
        static extern int InitComm(int port);//连接身份证阅读器 

        [DllImport("termb.dll")]
        static extern int InitCommExt();//自动搜索身份证阅读器并连接身份证阅读器 

        [DllImport("termb.dll")]
        static extern int CloseComm();//断开与身份证阅读器连接 

        [DllImport("termb.dll")]
        static extern int Authenticate();//判断是否有放卡，且是否身份证 

        [DllImport("termb.dll")]
        public static extern int Read_Content(int index);//读卡操作,信息文件存储在dll所在下

        [DllImport("termb.dll")]
        public static extern int ReadContent(int index);//读卡操作,信息文件存储在dll所在下

        [DllImport("termb.dll")]
        static extern int GetSAMID(StringBuilder SAMID);//获取SAM模块编号

        [DllImport("termb.dll")]
        static extern int GetSAMIDEx(StringBuilder SAMID);//获取SAM模块编号（10位编号）

        [DllImport("termb.dll")]
        static extern int GetBmpPhoto(string PhotoPath);//解析身份证照片

        [DllImport("termb.dll")]
        static extern int GetBmpPhotoToMem(byte[] imageData, int cbImageData);//解析身份证照片

        [DllImport("termb.dll")]
        static extern int GetBmpPhotoExt();//解析身份证照片

        [DllImport("termb.dll")]
        static extern int Reset_SAM();//重置Sam模块

        [DllImport("termb.dll")]
        static extern int GetSAMStatus();//获取SAM模块状态 

        [DllImport("termb.dll")]
        static extern int GetCardInfo(int index, StringBuilder value);//解析身份证信息 

        [DllImport("termb.dll")]
        static extern int ExportCardImageV();//生成竖版身份证正反两面图片(输出目录：dll所在目录的cardv.jpg和SetCardJPGPathNameV指定路径)

        [DllImport("termb.dll")]
        static extern int ExportCardImageH();//生成横版身份证正反两面图片(输出目录：dll所在目录的cardh.jpg和SetCardJPGPathNameH指定路径) 

        [DllImport("termb.dll")]
        static extern int SetTempDir(string DirPath);//设置生成文件临时目录

        [DllImport("termb.dll")]
        static extern int GetTempDir(StringBuilder path, int cbPath);//获取文件生成临时目录

        [DllImport("termb.dll")]
        static extern void GetPhotoJPGPathName(StringBuilder path, int cbPath);//获取jpg头像全路径名 


        [DllImport("termb.dll")]
        static extern int SetPhotoJPGPathName(string path);//设置jpg头像全路径名

        [DllImport("termb.dll")]
        static extern int SetCardJPGPathNameV(string path);//设置竖版身份证正反两面图片全路径

        [DllImport("termb.dll")]
        static extern int GetCardJPGPathNameV(StringBuilder path, int cbPath);//获取竖版身份证正反两面图片全路径

        [DllImport("termb.dll")]
        static extern int SetCardJPGPathNameH(string path);//设置横版身份证正反两面图片全路径

        [DllImport("termb.dll")]
        static extern int GetCardJPGPathNameH(StringBuilder path, int cbPath);//获取横版身份证正反两面图片全路径

        [DllImport("termb.dll")]
        static extern int getName(StringBuilder data, int cbData);//获取姓名

        [DllImport("termb.dll")]
        static extern int getSex(StringBuilder data, int cbData);//获取性别

        [DllImport("termb.dll")]
        static extern int getNation(StringBuilder data, int cbData);//获取民族

        [DllImport("termb.dll")]
        static extern int getBirthdate(StringBuilder data, int cbData);//获取生日(YYYYMMDD)

        [DllImport("termb.dll")]
        static extern int getAddress(StringBuilder data, int cbData);//获取地址

        [DllImport("termb.dll")]
        static extern int getIDNum(StringBuilder data, int cbData);//获取身份证号

        [DllImport("termb.dll")]
        static extern int getIssue(StringBuilder data, int cbData);//获取签发机关

        [DllImport("termb.dll")]
        static extern int getEffectedDate(StringBuilder data, int cbData);//获取有效期起始日期(YYYYMMDD)

        [DllImport("termb.dll")]
        static extern int getExpiredDate(StringBuilder data, int cbData);//获取有效期截止日期(YYYYMMDD) 

        [DllImport("termb.dll")]
        static extern int getBMPPhotoBase64(StringBuilder data, int cbData);//获取BMP头像Base64编码 

        [DllImport("termb.dll")]
        static extern int getJPGPhotoBase64(StringBuilder data, int cbData);//获取JPG头像Base64编码

        [DllImport("termb.dll")]
        static extern int getJPGCardBase64V(StringBuilder data, int cbData);//获取竖版身份证正反两面JPG图像base64编码字符串

        [DllImport("termb.dll")]
        static extern int getJPGCardBase64H(StringBuilder data, int cbData);//获取横版身份证正反两面JPG图像base64编码字符串

        [DllImport("termb.dll")]
        static extern int HIDVoice(int nVoice);//语音提示。。仅适用于与带HID语音设备的身份证阅读器（如ID200）

        [DllImport("termb.dll")]
        static extern int IC_SetDevNum(int iPort, StringBuilder data, int cbdata);//设置发卡器序列号

        [DllImport("termb.dll")]
        static extern int IC_GetDevNum(int iPort, StringBuilder data, int cbdata);//获取发卡器序列号

        [DllImport("termb.dll")]
        static extern int IC_GetDevVersion(int iPort, StringBuilder data, int cbdata);//设置发卡器序列号 

        [DllImport("termb.dll")]
        static extern int IC_WriteData(int iPort, int keyMode, int sector, int idx, StringBuilder key, StringBuilder data, int cbdata, ref int snr);//写数据

        [DllImport("termb.dll")]
        static extern int IC_ReadData(int iPort, int keyMode, int sector, int idx, StringBuilder key, StringBuilder data, int cbdata, ref int snr);//du数据

        [DllImport("termb.dll")]
        static extern int IC_GetICSnr(int iPort, ref int snr);//读IC卡物理卡号 

        [DllImport("termb.dll")]
        static extern int IC_GetIDSnr(int iPort, StringBuilder data, int cbdata);//读身份证物理卡号 

        [DllImport("termb.dll")]
        static extern int getEnName(StringBuilder data, int cbdata);//获取英文名

        [DllImport("termb.dll")]
        static extern int getCnName(StringBuilder data, int cbdata);//获取中文名 

        [DllImport("termb.dll")]
        static extern int getPassNum(StringBuilder data, int cbdata);//获取港澳台居通行证号码

        [DllImport("termb.dll")]
        static extern int getVisaTimes();//获取签发次数

        #endregion
        public string rev_servename;

        public clsAllnew()
        {
            ConStr = System.Web.Configuration.WebConfigurationManager.AppSettings["服务器1"];
            if (ConStr != null && ConStr != "")
                ConStrPIC = ConStr.Replace("Provider=SQLOLEDB;", "");



        }
        public List<clsuserinfo> findUser(string findtext)
        {
            try
            {
                //  string strSelect = "select * from emw_user where name='" + findtext + "'";
                OleDbConnection aConnection = new OleDbConnection(ConStr);

                List<clsuserinfo> ClaimReport_Server = new List<clsuserinfo>();
                if (aConnection.State == ConnectionState.Closed)
                    aConnection.Open();

                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(findtext, aConnection);
                OleDbCommandBuilder mybuilder = new OleDbCommandBuilder(myDataAdapter);
                DataSet ds = new DataSet();
                myDataAdapter.Fill(ds, "emw_user");
                foreach (DataRow reader in ds.Tables["emw_user"].Rows)
                {
                    clsuserinfo item = new clsuserinfo();

                    if (reader["_id"].ToString() != "")
                        item.Order_id = reader["_id"].ToString();
                    if (reader["name"].ToString() != "")
                        item.name = reader["name"].ToString();
                    if (reader["password"].ToString() != "")
                        item.password = reader["password"].ToString();
                    if (reader["Createdate"].ToString() != "")
                        item.Createdate = reader["Createdate"].ToString();
                    if (reader["Btype"].ToString() != "")
                        item.Btype = reader["Btype"].ToString();

                    if (reader["denglushijian"].ToString() != "")
                        item.denglushijian = reader["denglushijian"].ToString();
                    if (reader["jigoudaima"].ToString() != "")
                        item.jigoudaima = reader["jigoudaima"].ToString();
                    if (reader["userTime"].ToString() != "")
                        item.userTime = reader["userTime"].ToString();

                    if (reader["AdminIS"].ToString() != "")
                        item.AdminIS = reader["AdminIS"].ToString();
                    if (reader["mibao"].ToString() != "")
                        item.mibao = reader["mibao"].ToString();


                    ClaimReport_Server.Add(item);

                    //这里做数据处理....
                }
                return ClaimReport_Server;

            }
            catch (Exception ex)
            {
                //  inputlog(ex.Message + "//" + ex.Source + "//" + ex.StackTrace);
                HttpContext.Current.Response.Redirect("~/ErrorPage/ErrorPage.aspx?Error=" + "无法与服务器建立连接，请确保数据库配置或网络畅通！");

                throw ex;
            }
        }
        private static void inputlog(string aainput)
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "bin\\log.txt";
            StreamWriter sw = new StreamWriter(A_Path);
            sw.WriteLine(aainput);
            sw.Flush();
            sw.Close();
        }
        public void createUser_Server(List<clsuserinfo> AddMAPResult)
        {




            //创建连接对象
            bool isok = false;
            OleDbConnection con = new OleDbConnection(ConStr);
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                //命令
                foreach (clsuserinfo item in AddMAPResult)
                {

                    string sql = "";
                    sql = "insert into emw_user(name,password,Createdate,Btype,denglushijian,jigoudaima,userTime,AdminIS,mibao) values ('" + item.name + "','" + item.password + "',N'" + item.Createdate + "','" + item.Btype + "','" + item.denglushijian + "','" + item.jigoudaima + "','" + item.userTime + "','" + item.AdminIS + "','" + item.mibao + "')";

                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    isok = true;

                }
                //con.Close();
                return;
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open) con.Close();
                if (con != null)
                    con.Dispose();
                return;

                throw;
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); con.Dispose(); }
        }
        public List<clsuserinfo> ReadUserlistfromServer()
        {
            string conditions = "select * from emw_user";//成功

            OleDbConnection aConnection = new OleDbConnection(ConStr);

            List<clsuserinfo> ClaimReport_Server = new List<clsuserinfo>();
            if (aConnection.State == ConnectionState.Closed)
                aConnection.Open();

            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(conditions, aConnection);
            OleDbCommandBuilder mybuilder = new OleDbCommandBuilder(myDataAdapter);
            DataSet ds = new DataSet();
            myDataAdapter.Fill(ds, "emw_user");
            foreach (DataRow reader in ds.Tables["emw_user"].Rows)
            {
                clsuserinfo item = new clsuserinfo();

                if (reader["_id"].ToString() != "")
                    item.Order_id = reader["_id"].ToString();
                if (reader["name"].ToString() != "")
                    item.name = reader["name"].ToString();
                if (reader["password"].ToString() != "")
                    item.password = reader["password"].ToString();
                if (reader["Createdate"].ToString() != "")
                    item.Createdate = reader["Createdate"].ToString();
                if (reader["Btype"].ToString() != "")
                    item.Btype = reader["Btype"].ToString();

                if (reader["denglushijian"].ToString() != "")
                    item.denglushijian = reader["denglushijian"].ToString();
                if (reader["jigoudaima"].ToString() != "")
                    item.jigoudaima = reader["jigoudaima"].ToString();
                if (reader["userTime"].ToString() != "")
                    item.userTime = reader["userTime"].ToString();

                if (reader["AdminIS"].ToString() != "")
                    item.AdminIS = reader["AdminIS"].ToString();

                if (reader["mibao"].ToString() != "")
                    item.mibao = reader["mibao"].ToString();

                ClaimReport_Server.Add(item);

                //这里做数据处理....
            }
            return ClaimReport_Server;

        }
        public bool deleteUSER(string name)
        {



            OleDbConnection con = new OleDbConnection(ConStr);
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                string sql2 = "delete from emw_user where   name='" + name + "'";

                OleDbCommand cmd = new OleDbCommand(sql2, con);
                cmd.ExecuteNonQuery();

                return true;
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open) con.Close();
                if (con != null)
                    con.Dispose();
                return false;

                throw;
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); con.Dispose(); }

        }
        public bool changeUserpassword_Server(List<clsuserinfo> AddMAPResult)
        {
            //创建连接对象
            bool isok = false;
            OleDbConnection con = new OleDbConnection(ConStr);
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                //命令
                foreach (clsuserinfo item in AddMAPResult)
                {

                    string sql = "";
                    string conditions = "";
                    if (item.password != null)
                    {
                        conditions += " password ='" + item.password + "'";
                    }
                    if (item.name != null)
                    {
                        conditions += " ,name ='" + item.name + "'";
                    }
                    if (item.Btype != null)
                    {
                        conditions += " ,Btype ='" + item.Btype + "'";
                    }
                    if (item.denglushijian != null)
                    {
                        conditions += " ,denglushijian ='" + item.denglushijian + "'";
                    }
                    if (item.Createdate != null)
                    {
                        conditions += " ,Createdate ='" + item.Createdate + "'";
                    }
                    if (item.AdminIS != null)
                    {
                        conditions += " ,AdminIS ='" + item.AdminIS + "'";
                    }
                    if (item.jigoudaima != null)
                    {
                        conditions += " ,jigoudaima ='" + item.jigoudaima + "'";
                    }
                    if (item.userTime != null)
                    {
                        conditions += " ,userTime ='" + item.userTime + "'";
                    }
                    if (item.mibao != null)
                    {
                        conditions += " ,mibao ='" + item.mibao + "'";
                    }


                    conditions = "update emw_user set  " + conditions + " where _id = " + item.Order_id + " ";
                    sql = conditions;

                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    isok = true;

                }
                //con.Close();
                return isok;
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open) con.Close();
                if (con != null)
                    con.Dispose();
                return false;

                throw;
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); con.Dispose(); }

        }


        #region 读取IC卡设备
        public List<clCard_info> Read_card()
        {
            #region 假数据
            //string image64 = ImgToBase64String(@"D:\Devlop\身份证阅读器二次开发软件说明\cardv.jpg");
            ////string m_strPath = Application.StartupPath;

            ////Base64ToImage(image64).Save(m_strPath + "\\Hello.jpg");
            //Base64ToImage(image64).Save(mdbpath2_Ctirx);


            //List<clCard_info> reads1 = new List<clCard_info>();


            //clCard_info item = new clCard_info();
            //item.daima_gonghao = "d1ll";
            //item.zhengjianhaoma = "12345";
            //item.tupian = item.zhengjianhaoma;
            //item.FData = image64;

            //reads1.Add(item);
            //return reads1; 
            #endregion
            try
            {

                int AutoSearchReader = InitCommExt();
                if (AutoSearchReader > 0)
                {
                    Port = AutoSearchReader;
                    IsConnected = true;
                    //textBox_Name.Text = AutoSearchReader.ToString();

                    StringBuilder sb = new StringBuilder(cbDataSize);
                    GetSAMID(sb);
                    //  MessageBox.Show("连接身份证阅读器成功,SAM模块编号:" + sb);
                    //button_Connect.Enabled = false;
                    //button_ReadCard.Enabled = true;
                    //button_DisConnect.Enabled = true;
                    #region 读取

                    List<clCard_info> reads = button_ReadCard();



                    #endregion



                    return reads;
                }
                else
                {
                    MessageBox.Show("检查是否正确连接设备");
                }
            }
            catch (Exception exx)
            {

                throw;
            }
            return null;
        }
        private List<clCard_info> button_ReadCard()
        {

            List<clCard_info> resulits = new List<clCard_info>();

            //卡认证
            int FindCard = Authenticate();

            int rs = Read_Content(1);

            if (rs != 1 && rs != 2 && rs != 3)
            {

                return null;
            }
            clCard_info item = new clCard_info();

            //读卡成功
            //姓名
            StringBuilder sb = new StringBuilder(cbDataSize);
            getName(sb, cbDataSize);
            item.mingcheng = sb.ToString();

            //民族/国家
            getNation(sb, cbDataSize);
            item.minzu = sb.ToString();

            //性别 
            getSex(sb, cbDataSize);
            item.xingbie = sb.ToString();

            //出生 
            getBirthdate(sb, cbDataSize);
            item.chushengriqi = sb.ToString();
            // sb.ToString().Substring(0, 4) - sb.ToString().Substring(4, 2) - sb.ToString().Substring(6, 2);

            //地址 
            getAddress(sb, cbDataSize);
            string ad = sb.ToString();
            item.jiatingzhuzhi = ad;

            //号码 
            getIDNum(sb, cbDataSize);
            item.zhengjianhaoma = sb.ToString();

            //机关 
            getIssue(sb, cbDataSize);
            //textBox_Issue.Text = sb.ToString();

            //有效期 
            getEffectedDate(sb, cbDataSize);
            string aa = sb.ToString();
            getExpiredDate(sb, cbDataSize);
            item.zhengjianyouxiao = aa + sb.ToString();

            //通行证号  
            getPassNum(sb, cbDataSize);
            //textBox_PassNum.Text = sb.ToString();

            //签证次数  
            //textBox_VisaTimes.Text = "" + getVisaTimes();

            //英文名 
            getEnName(sb, cbDataSize);
            //textBox_EnName.Text = sb.ToString();

            //中文名  
            getCnName(sb, cbDataSize);
            //textBox_CnName.Text = sb.ToString();

            //证件类型
            GetCardInfo(105, sb);
            if ("1" == sb.ToString())
            {
                //textBox_CardType.Text = "居民身份证";
            }
            else if ("3" == sb.ToString())
            {
                //textBox_CardType.Text = "港澳台居住证";
            }
            else
            {
                //textBox_CardType.Text = "外国人居住证";
            }
            //图片
            getJPGCardBase64H(sb, cbDataSize);
            item.tupian = item.zhengjianhaoma;
            item.FData = sb.ToString();

            resulits.Add(item);
            return resulits;

        }
        public string ImgToBase64String(string Imagefilename)
        {
            try
            {
                Bitmap bmp = new Bitmap(Imagefilename);
                MemoryStream ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                byte[] arr = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(arr, 0, (int)ms.Length);
                ms.Close();
                return Convert.ToBase64String(arr);
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public System.Drawing.Image Base64ToImage(string base64String)
        {
            byte[] imageBytes = Convert.FromBase64String(base64String);
            MemoryStream ms = new MemoryStream(imageBytes, 0, imageBytes.Length);
            ms.Write(imageBytes, 0, imageBytes.Length);
            System.Drawing.Image image = System.Drawing.Image.FromStream(ms, true);
            return image;
        }

        #endregion
        #region 写入金蝶数据库

        public void createICcard_info_Server(List<cls_order_info> AddMAPResult)
        {


            //创建连接对象
            bool isok = false;
            OleDbConnection con = new OleDbConnection(ConStr);
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                //命令
                foreach (cls_order_info item in AddMAPResult)
                {

                    string sql = "";
                    sql = "insert into NewMeasueDataTable(PotNo,DDate,AlCnt,Lsp,Djzsp,Djwd,Fzb,FeCnt,SiCnt,AlOCnt,CaFCnt,MgCnt,LDYJ,MLsp,LPW) values ('" + item.PotNo + "','" + item.DDate + "',N'" + item.AlCnt + "',N'" + item.Lsp + "','" + item.Djzsp + "','" + item.Djwd + "','" + item.Fzb + "','" + item.FeCnt + "','" + item.SiCnt + "','" + item.AlOCnt + "','" + item.CaFCnt + "','" + item.MgCnt + "','" + item.LDYJ + "','" + item.MLsp + "','" + item.LPW + "')";

                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    isok = true;

                }
                //con.Close();
                return;
            }
            catch (Exception ex)
            {

                if (con.State == ConnectionState.Open) con.Close();
                if (con != null)
                    con.Dispose();

                HttpContext.Current.Response.Redirect("~/ErrorPage/ErrorPage.aspx?Error=" + ex.ToString());

                throw ex;
                return;

                throw;
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); con.Dispose(); }
        }
        public void create_t_Item_info_Server(List<clt_Item_info> AddMAPResult)
        {


            //创建连接对象
            bool isok = false;
            OleDbConnection con = new OleDbConnection(ConStr);
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                //命令
                foreach (clt_Item_info item in AddMAPResult)
                {

                    string sql = "";
                    sql = "insert into t_Item(FItemID,FItemClassID,FExternID,FNumber,FParentID,FLevel,FDetail,FName,FUnUsed,FBrNo,FFullNumber,FDiff,FDeleted,FShortNumber,FFullName,FGRCommonID,FSystemType,FUseSign,FAccessory,FGrControl,FHavePicture) values ('" + item.FItemID + "','" + item.FItemClassID + "',N'" + item.FExternID + "',N'" + item.FNumber + "','" + item.FParentID + "','" + item.FLevel + "','" + item.FDetail + "','" + item.FName + "','" + item.FUnUsed + "','" + item.FBrNo + "','" + item.FFullNumber + "','" + item.FDiff + "','" + item.FDeleted + "','" + item.FShortNumber + "','" + item.FFullName + "','" + item.FGRCommonID + "','" + item.FSystemType + "','" + item.FUseSign + "','" + item.FAccessory + "','" + item.FGrControl + "','" + item.FHavePicture + "')";

                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    isok = true;

                }
                //con.Close();
                return;
            }
            catch (Exception ex)
            {

                if (con.State == ConnectionState.Open) con.Close();
                if (con != null)
                    con.Dispose();

                HttpContext.Current.Response.Redirect("~/ErrorPage/ErrorPage.aspx?Error=" + ex.ToString());

                throw ex;
                return;

                throw;
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); con.Dispose(); }
        }


        public List<cls_order_info> Readt_ItemServer(string conditions)
        {

            try
            {
                OleDbConnection aConnection = new OleDbConnection(ConStr);

                List<cls_order_info> ClaimReport_Server = new List<cls_order_info>();
                if (aConnection.State == ConnectionState.Closed)
                    aConnection.Open();

                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(conditions, aConnection);
                OleDbCommandBuilder mybuilder = new OleDbCommandBuilder(myDataAdapter);
                DataSet ds = new DataSet();
                myDataAdapter.Fill(ds, "NewMeasueDataTable");
                foreach (DataRow reader in ds.Tables["NewMeasueDataTable"].Rows)
                {
                    cls_order_info item = new cls_order_info();

                    if (reader["Order_id"].ToString() != "")
                        item.Order_id = reader["Order_id"].ToString();
                    if (reader["PotNo"].ToString() != "")
                        item.PotNo = reader["PotNo"].ToString();

                    if (reader["DDate"].ToString() != "")
                    {
                        item.DDate = reader["DDate"].ToString();
                        item.DDate = clsCommHelp.objToDateTime1(reader["DDate"].ToString());

                    }
                    if (reader["AlCnt"].ToString() != "")
                        item.AlCnt = reader["AlCnt"].ToString();
                    if (reader["Lsp"].ToString() != "")
                        item.Lsp = reader["Lsp"].ToString();

                    if (reader["Djzsp"].ToString() != "")
                        item.Djzsp = reader["Djzsp"].ToString();

                    if (reader["Djwd"].ToString() != "")
                        item.Djwd = reader["Djwd"].ToString();

                    if (reader["Fzb"].ToString() != "")
                        item.Fzb = reader["Fzb"].ToString();

                    if (reader["FeCnt"].ToString() != "")
                        item.FeCnt = reader["FeCnt"].ToString();


                    if (reader["SiCnt"].ToString() != "")
                        item.SiCnt = reader["SiCnt"].ToString();

                    if (reader["AlOCnt"].ToString() != "")
                        item.AlOCnt = reader["AlOCnt"].ToString();

                    if (reader["CaFCnt"].ToString() != "")
                        item.CaFCnt = reader["CaFCnt"].ToString();

                    if (reader["MgCnt"].ToString() != "")
                        item.MgCnt = reader["MgCnt"].ToString();

                    if (reader["LDYJ"].ToString() != "")
                        item.LDYJ = reader["LDYJ"].ToString();

                    if (reader["MLsp"].ToString() != "")
                        item.MLsp = reader["MLsp"].ToString();

                    if (reader["LPW"].ToString() != "")
                        item.LPW = reader["LPW"].ToString();

                    ClaimReport_Server.Add(item);


                }
                return ClaimReport_Server;
            }
            catch (Exception ex)
            {
                HttpContext.Current.Response.Redirect("~/ErrorPage/ErrorPage.aspx?Error=" + "网络访问较慢或网络不通无法访问 ：" + ex.ToString());

                // inputlog(ex.Message + "//" + ex.Source + "//" + ex.StackTrace);

                throw ex;
            }

        }
        public bool deleteCard(string sql2)
        {
            OleDbConnection con = new OleDbConnection(ConStr);
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();

                OleDbCommand cmd = new OleDbCommand(sql2, con);
                cmd.ExecuteNonQuery();

                return true;
            }
            catch (Exception ex)
            {

                throw ex;

                if (con.State == ConnectionState.Open) con.Close();
                if (con != null)
                    con.Dispose();

                HttpContext.Current.Response.Redirect("~/ErrorPage/ErrorPage.aspx?Error=" + ex.ToString());

                return false;

                throw;
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); con.Dispose(); }

        }

        public bool changeCardServer(List<cls_order_info> AddMAPResult)
        {
            //创建连接对象
            bool isok = false;
            OleDbConnection con = new OleDbConnection(ConStr);
            try
            {
                if (con.State == ConnectionState.Closed)
                    con.Open();
                //命令
                foreach (cls_order_info item in AddMAPResult)
                {

                    string sql = "";
                    string conditions = "";
                    if (item.PotNo != null)
                    {
                        conditions += " PotNo ='" + item.PotNo + "'";
                    }
                    if (item.DDate != null)
                    {
                        conditions += " ,DDate ='" + item.DDate + "'";
                    }
                    if (item.AlCnt != null)
                    {
                        conditions += " ,AlCnt ='" + item.AlCnt + "'";
                    }
                    if (item.Lsp != null)
                    {
                        conditions += " ,Lsp ='" + item.Lsp + "'";
                    }
                    if (item.Djzsp != null)
                    {
                        conditions += " ,Djzsp ='" + item.Djzsp + "'";
                    }
                    if (item.Djwd != null)
                    {
                        conditions += " ,Djwd ='" + item.Djwd + "'";
                    }
                    if (item.Fzb != null)
                    {
                        conditions += " ,Fzb ='" + item.Fzb + "'";
                    }
                    if (item.FeCnt != null)
                    {
                        conditions += " ,FeCnt ='" + item.FeCnt + "'";
                    }
                    if (item.SiCnt != null)
                    {
                        conditions += " ,SiCnt ='" + item.SiCnt + "'";
                    }
                    if (item.AlOCnt != null)
                    {
                        conditions += " ,AlOCnt ='" + item.AlOCnt + "'";
                    }
                    if (item.CaFCnt != null)
                    {
                        conditions += " ,CaFCnt ='" + item.CaFCnt + "'";
                    }
                    if (item.MgCnt != null)
                    {
                        conditions += " ,MgCnt ='" + item.MgCnt + "'";
                    }
                    if (item.LDYJ != null)
                    {
                        conditions += " ,LDYJ ='" + item.LDYJ + "'";
                    }
                    if (item.MLsp != null)
                    {
                        conditions += " ,MLsp ='" + item.MLsp + "'";
                    }
                    if (item.LPW != null)
                    {
                        conditions += " ,LPW ='" + item.LPW + "'";
                    }
                    conditions = "update NewMeasueDataTable set  " + conditions + " where Order_id = " + item.Order_id + " ";
                    sql = conditions;

                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    isok = true;

                }
                //con.Close();
                return isok;
            }
            catch (Exception ex)
            {

                if (con.State == ConnectionState.Open) con.Close();
                if (con != null)
                    con.Dispose();
                HttpContext.Current.Response.Redirect("~/ErrorPage/ErrorPage.aspx?Error=" + ex.ToString());

                return false;

                throw;
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); con.Dispose(); }

        }

        #endregion



        #region wbr
        public void InputNewBaobiao_RawDatat()
        {
            //axWebBrowser1_NavigateComplete2();

            InitialWebbroswer1();
            return;

            try
            {

                string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\null.xlsx");
                //需要换 成日期的导出表
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(ZFCEPath, Type.Missing, true, Type.Missing,
                    "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                excelApp.Visible = true;
                excelApp.ScreenUpdating = true;

                //Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["Vendor Statement(1)"];
                //Microsoft.Office.Interop.Excel.Range rng;
                //rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 45]);
                //int rowCount = WS.UsedRange.Rows.Count - 1;
                //object[,] o = new object[1, 1];
                //o = (object[,])rng.Value2;
                //clsCommHelp.CloseExcel(excelApp, analyWK);

            }
            catch (Exception ex)
            {
                throw ex;
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

                string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xls");

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
        void MyWebBrower_BeforeNewWindow(object sender, WebBrowserExtendedNavigatingEventArgs e)
        {
            string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xls");

            #region 在原有窗口导航出新页
            e.Cancel = false;//http://pro.wwpack-crest.hp.com/wwpak.online/regResults.aspx
         //   MyWebBrower.Navigate(ZFCEPath);
            #endregion
        }

        protected void AnalysisWebInfo1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string ZFCEPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\报价单.xls");
            //WbBlockNewUrl myDoc = sender as WbBlockNewUrl;
            //http://dciw-unilever.ihost.com/exprimo/
            //if (myDoc.Url.ToString().IndexOf(ZFCEPath) >= 0)
            //{
            //    HtmlElement submit = null;
            //    HtmlElementCollection a = myDoc.Document.GetElementsByTagName("a");
            //    int aaa = 0;
            //    foreach (HtmlElement item in a)
            //    {
            //        if (item.OuterHtml.IndexOf("a") > 0)
            //        {
            //            submit = item;
            //            break;
            //        }
            //    }
            //    //  submit.InvokeMember("Click");
            //}
        }



        #endregion




    }
}
