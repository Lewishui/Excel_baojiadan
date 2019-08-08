using China_System.Common;
using clsBuiness;
using MasterClassified;
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
using System.Text.RegularExpressions;
using System.Threading;

namespace QCAuto
{
    public partial class frm调拨系统 : Form
    {

        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        public string path;
        public string path2;
        private China_System.Common.clsCommHelp.SortableBindingList<cls_diaobo_info> sortabledinningsOrderList;
        List<cls_kucun_info> kucunResults;
        List<cls_xiaoshou_info> xiaoshouResults;

        private China_System.Common.clsCommHelp.SortableBindingList<cls_xiaoshou_info> xiaoshou_sortabledinningsOrderList;
        private China_System.Common.clsCommHelp.SortableBindingList<cls_kucun_info> kucun_sortabledinningsOrderList;
        public List<string> mdzz = new List<string>();

        List<cls_xiaoshou_info> ALllResults;
        int RowRemark = 0;
        int cloumn = 0;
        List<cls_diaobo_info> ShowResults;
        List<cls_diaobo_info> FilterShowResults;
        List<cls_diaobo_info> Show2Results;
        List<cls_xiaoshou_info> shaixuan_xiaoshouResults;
        List<cls_kucun_info> kuncun_xiaoshouResults;
        DataTable qtyTable;
        DataTable allqtyTable;
        int f_count;
        int cc;
        private bool IsRun = false;
        string[] mz = new string[0];
        DataTable shib = new DataTable();
        private Thread GetDataforRawDataThread;
        private System.Timers.Timer timerAlter_new;
        public frm调拨系统(string password)
        {
            InitializeComponent();

            //checkedListBox1.Items.Clear();

            //for (int j = 0; j < dataGridView.ColumnCount; j++)
            //{
            //    string MU = dataGridView.Columns[j].HeaderText;

            //    checkedListBox1.Items.Add(MU);
            //}
            NewMethod();


              
        }
        private void NewMethod()
        {
            timerAlter_new = new System.Timers.Timer(666);
            timerAlter_new.Elapsed += new System.Timers.ElapsedEventHandler(TimeControl);
            timerAlter_new.AutoReset = true;
            timerAlter_new.Start();
        }
        private void TimeControl(object sender, EventArgs e)
        {
            if (!IsRun)
            {
                IsRun = true;
                GetDataforRawDataThread = new Thread(TimeMethod);
                GetDataforRawDataThread.Start();
            }
        }
        private void TimeMethod()
        {
            bool istrue = true;
            clsmytest buiness = new clsmytest();

            bool istue = buiness.checkname("frm调拨系统", "yhltd");
            if (istue == false)
            {
                Control.CheckForIllegalCrossThreadCalls = false;
                this.Visible = false; 
                //MessageBox.Show("缺失系统文件，或电脑系统更新导致，请联系开发人员 !", "系统错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var form = new frmAlterinfo("缺失系统文件，或电脑系统更新导致，请联系开发人员 !");
             
                if (form.ShowDialog() == DialogResult.OK)
                {
                  
                }
              

                System.Environment.Exit(0);
            }

            IsRun = false;
        }
        public frm调拨系统()
        {
            // TODO: Complete member initialization
        }

        private void webBrowser1_DocumentCompleted_1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

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

        private void 读取_Click(object sender, EventArgs e)
        {
            var form = new frmImportFile();
            if (form.ShowDialog() == DialogResult.OK)
            {
                path = form.path;
                path2 = form.path2;
            }
            if (path2 == null || path2 == "" || path == null || path == "")
                return;



            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(KEYFile);

                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    InitializeDataGridView();

                    davshow(ShowResults);





                }
            }
            catch (Exception ex)
            {
                throw ex;
            }





        }


        private void InitializeDataGridView()
        {
            var counties = ShowResults.Select(s => new MockEntity { ShortName = s.kuanhao, FullName = s.kuanhao }).Distinct().ToList();
            counties.Insert(0, new MockEntity { ShortName = "全部", FullName = "全部" });

            //this.comboBox3.DisplayMember = "FullName";
            //this.comboBox3.ValueMember = "ShortName";
            //this.comboBox3.DataSource = counties;


            var counties1 = ShowResults.Select(s => new MockEntity { ShortName = s.kuanhao, FullName = s.kuanhao }).Distinct().ToList();
            counties1.Insert(0, new MockEntity { ShortName = "全部", FullName = "全部" });

            this.comboBox4.DisplayMember = "FullName";
            this.comboBox4.ValueMember = "ShortName";
            this.comboBox4.DataSource = counties1;

            var counties2 = ShowResults.Select(s => new MockEntity { ShortName = s.cima, FullName = s.cima }).Distinct().ToList();
            counties2.Insert(0, new MockEntity { ShortName = "全部", FullName = "全部" });

            this.comboBox5.DisplayMember = "FullName";
            this.comboBox5.ValueMember = "ShortName";
            this.comboBox5.DataSource = counties2;

            var counties3 = ShowResults.Select(s => new MockEntity { ShortName = s.kuanhao, FullName = s.kuanhao }).Distinct().ToList();
            counties3.Insert(0, new MockEntity { ShortName = "全部", FullName = "全部" });

            this.comboBox1.DisplayMember = "FullName";
            this.comboBox1.ValueMember = "ShortName";
            this.comboBox1.DataSource = counties3;



            var counties6 = ShowResults.Select(s => new MockEntity { ShortName = s.kuanhao, FullName = s.kuanhao }).Distinct().ToList();
            counties6.Insert(0, new MockEntity { ShortName = "全部", FullName = "全部" });


            this.comboBox2.DisplayMember = "FullName";
            this.comboBox2.ValueMember = "ShortName";
            this.comboBox2.DataSource = counties6;



            //var counties3 = ShowResults.Select(s => new MockEntity { ShortName = s.xingbie, FullName = s.kuanhao }).Distinct().ToList();
            //counties.Insert(0, new MockEntity { ShortName = "全部", FullName = "全部" });

            //this.comboBox6.DisplayMember = "FullName";
            //this.comboBox6.ValueMember = "ShortName";
            //this.comboBox6.DataSource = counties3;



            //单据类型	店铺名称	款号	款名	颜色	尺码	数量	吊牌价	性别	季节	类别	年份

            //checkedListBox1.Items.Add("单据类型");
            //checkedListBox1.Items.Add("店铺名称");
            //checkedListBox1.Items.Add("款号");
            //checkedListBox1.Items.Add("款名");
            //checkedListBox1.Items.Add("颜色");
            //checkedListBox1.Items.Add("尺码");
            //checkedListBox1.Items.Add("数量");
            //checkedListBox1.Items.Add("吊牌价");
            //checkedListBox1.Items.Add("性别");
            //checkedListBox1.Items.Add("季节");
            //checkedListBox1.Items.Add("年份");




        }

        private void davshow(List<cls_diaobo_info> ShowResults)
        {
            sortabledinningsOrderList = new China_System.Common.clsCommHelp.SortableBindingList<cls_diaobo_info>(ShowResults);
            this.bindingSource1.DataSource = this.sortabledinningsOrderList;
            //   this.dataGridView.DataSource = null;
            //this.dataGridView.AutoGenerateColumns = false;
            //this.dataGridView.DataSource = this.bindingSource1;
            //if (dataGridView.DataSource != null)
            //    toolStripLabel1.Text = "条数： " + dataGridView.RowCount;
            //int s = this.tabControl1.SelectedIndex;
            //if (s == 0)
            //{
            //    if (dataGridView.DataSource != null)
            //        toolStripLabel1.Text = "条数： " + dataGridView.RowCount;
            //}



            //销售 显示

            xiaoshou_sortabledinningsOrderList = new China_System.Common.clsCommHelp.SortableBindingList<cls_xiaoshou_info>(xiaoshouResults);
            this.bindingSource3.DataSource = this.xiaoshou_sortabledinningsOrderList;

            this.dataGridView2.AutoGenerateColumns = false;
            this.dataGridView2.DataSource = this.bindingSource3;
            //库存
            kucun_sortabledinningsOrderList = new China_System.Common.clsCommHelp.SortableBindingList<cls_kucun_info>(kucunResults);
            this.bindingSource4.DataSource = this.kucun_sortabledinningsOrderList;

            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = this.bindingSource4;




            dav3(qtyTable);

        }

        private void dav3(DataTable qtyTable)
        {
            this.bindingSource2.DataSource = null;

            this.bindingSource2.DataSource = allqtyTable;
            this.dataGridView3.DataSource = this.bindingSource2;

            for (int j = 0; j < dataGridView3.ColumnCount; j++)
            {
                string MU = dataGridView3.Columns[j].HeaderText;

              //  checkedListBox2.Items.Add(MU);
            }
            int s = this.tabControl1.SelectedIndex;
            if (s == 0)
            {

            }
            else if (s == 1)
            {

                if (dataGridView3.DataSource != null)
                    toolStripLabel1.Text = "条数： " + dataGridView3.RowCount;

            }


        }
        private void KEYFile(object sender, DoWorkEventArgs e)
        {
            // kucunResults = new List<cls_kucun_info>();

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            //BusinessHelp.pbStatus = pbStatus;
            //BusinessHelp.tsStatusLabel1 = toolStripLabel1;
            DateTime oldDate = DateTime.Now;
            Buiness_Bankcharge(ref this.bgWorker, "A", "", "");
            DateTime FinishTime = DateTime.Now;  //   
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);
        }


        public List<cls_kucun_info> Buiness_Bankcharge(ref BackgroundWorker bgWorker, string casetype, string Password, string USER)
        {
            shib.Columns.Add("类别", System.Type.GetType("System.String"));//0
            shib.Columns.Add("调出店名", System.Type.GetType("System.String"));//0
            shib.Columns.Add("入库店名", System.Type.GetType("System.String"));//0
            shib.Columns.Add("35(S)", System.Type.GetType("System.String"));//1
            shib.Columns.Add("36(M)", System.Type.GetType("System.String"));//2
            shib.Columns.Add("37(L)", System.Type.GetType("System.String"));//3
            shib.Columns.Add("38(XL)", System.Type.GetType("System.String"));//4
            shib.Columns.Add("39(2XL)", System.Type.GetType("System.String"));//5
            shib.Columns.Add("40(3XL)", System.Type.GetType("System.String"));//6
            shib.Columns.Add("41(4XL)", System.Type.GetType("System.String"));//7
            shib.Columns.Add("42(5XL)", System.Type.GetType("System.String"));//8
            shib.Columns.Add("43(6XL)", System.Type.GetType("System.String"));//9
            shib.Columns.Add("44(7XL)", System.Type.GetType("System.String"));//10
            shib.Columns.Add("45(8XL)", System.Type.GetType("System.String"));//11
            shib.Columns.Add("46(9XL)", System.Type.GetType("System.String"));//12
            shib.Columns.Add("47(10XL)", System.Type.GetType("System.String"));//13
            shib.Columns.Add("48(11XL)", System.Type.GetType("System.String"));//14
            shib.Rows.Add(shib.NewRow());





            string fin = "";
            xiaoshouResults = new List<cls_xiaoshou_info>();
            kucunResults = new List<cls_kucun_info>();
            ALllResults = new List<cls_xiaoshou_info>();
            ShowResults = new List<cls_diaobo_info>();

            xiaoshouResults = ReadxiaoshouFile(path2);
            kucunResults = ReadfindngFile(path);




            #region 显示1
            List<cls_xiaoshou_info> ALllResults1 = ALllResults.Where((x, ii) => ALllResults.FindIndex(z => z.kuanhao == x.kuanhao && z.kuanming == x.kuanming && z.cima == x.cima && z.yanse == x.yanse) == ii).ToList();//Lambda表达式去重  
            //查找店铺名字
            List<string> farenvalue = (from v in ALllResults select v.dianpumingcheng).Distinct().ToList();

            //for (int i = 0; i < farenvalue.Count; i++)
            {

                //List<cls_xiaoshou_info> findsapinfo = ALllResults1.FindAll(o => o.dianpumingcheng != null && farenvalue[i].Contains(o.dianpumingcheng));



                foreach (cls_xiaoshou_info temp in ALllResults1)
                {
                    if (temp.kuanhao == "E92008H" && temp.kuanming == "轻逸跑鞋" && temp.yanse == "黑色" && temp.cima == "38")
                    {

                    }
                    List<cls_xiaoshou_info> xiaoshouinfo = xiaoshouResults.FindAll(o => o.dianpumingcheng != null && temp.dianpumingcheng.Contains(o.dianpumingcheng) && o.kuanhao != null && temp.kuanhao.Contains(o.kuanhao) && o.cima != null && temp.cima.Contains(o.cima) && o.xingbie != null && temp.xingbie.Contains(o.xingbie) && o.danjuleixing != null && o.danjuleixing.Contains("销") && o.yanse != null && temp.yanse.Contains(o.yanse));
                    bool ist = false;

                    //销售
                    cls_diaobo_info item = new cls_diaobo_info();
                    if (xiaoshouinfo.Count > 0)
                    {
                        ist = true;

                        item.kuanhao = xiaoshouinfo[0].kuanhao;
                        item.xingbie = xiaoshouinfo[0].xingbie;
                        item.cima = xiaoshouinfo[0].cima;
                        item.fendianming = xiaoshouinfo[0].dianpumingcheng;
                        item.fendianming_xiaoshou = xiaoshouinfo[0].shuliang;


                        //new
                        item.kuanming = xiaoshouinfo[0].kuanming;
                        item.yanse = xiaoshouinfo[0].yanse;
                        item.diaopaijia = xiaoshouinfo[0].diaopaijia;
                        item.jijie = xiaoshouinfo[0].jijie;
                        item.leibie = xiaoshouinfo[0].leibie;
                        item.nianfen = xiaoshouinfo[0].nianfen;


                    }
                    //库存
                    List<cls_kucun_info> kucuninfo = kucunResults.FindAll(o => o.cangku != null && temp.dianpumingcheng.Contains(o.cangku) && o.kuanhao != null && temp.kuanhao.Contains(o.kuanhao) && o.cima != null && temp.cima.Contains(o.cima) && o.xingbie != null && temp.xingbie.Contains(o.xingbie) && o.danjuleixing != null && o.danjuleixing.Contains("库") && o.yanse != null && temp.yanse.Contains(o.yanse));
                    if (kucuninfo.Count > 0)
                    {
                        ist = true;
                        item.kucunming = kucuninfo[0].cangku;
                        item.kucun_shengyu = kucuninfo[0].kucunshuliang;

                        //new
                        item.kuanming = kucuninfo[0].kuanming;
                        item.yanse = kucuninfo[0].yanse;
                        item.diaopaijia = kucuninfo[0].diaopaijia;
                        item.jijie = kucuninfo[0].jijie;
                        item.leibie = kucuninfo[0].leibie;
                        item.nianfen = kucuninfo[0].nianfen;


                    }
                    if (item.kuanhao == null || item.kuanhao.Length < 1 || temp.kuanhao.Length < 1)
                    {


                    }

                    item.kuanhao = temp.kuanhao;
                    item.xingbie = temp.xingbie;
                    item.cima = temp.cima;
                    if (item.kuanhao == "E92008H" && item.kuanming == "轻逸跑鞋" && item.yanse == "黑色" && item.cima == "38")
                    {

                    }
                    if (ist == true)
                        ShowResults.Add(item);


                }
            #endregion
                #region 显示2
                qtyTable = new DataTable();

           //     qtyTable.Columns.Add("类型", System.Type.GetType("System.String"));//0
                qtyTable.Columns.Add("款号", System.Type.GetType("System.String"));//0
                qtyTable.Columns.Add("性别", System.Type.GetType("System.String"));//0
                qtyTable.Columns.Add("尺码", System.Type.GetType("System.String"));//0
                qtyTable.Columns.Add("颜色", System.Type.GetType("System.String"));//0
                qtyTable.Columns.Add("年份", System.Type.GetType("System.String"));//0
                //销
                for (int i = 0; i < farenvalue.Count; i++)
                    qtyTable.Columns.Add(farenvalue[i]+"(销)", System.Type.GetType("System.String"));//0

              //  f_count = farenvalue.Count;

                qtyTable.Columns.Add("汇总(销)", System.Type.GetType("System.String"));//0
            //    qtyTable.Columns.Add("剩余(销)", System.Type.GetType("System.String"));//0

                //库
                for (int i = 0; i < farenvalue.Count; i++)
                {
                    qtyTable.Columns.Add(farenvalue[i] + "(库)", System.Type.GetType("System.String"));//0
                    
                   // MessageBox.Show(mz[i]);
                }
                f_count = farenvalue.Count;
                mdzz = farenvalue;
                    
                qtyTable.Columns.Add("汇总(库)", System.Type.GetType("System.String"));//0
           //     qtyTable.Columns.Add("剩余", System.Type.GetType("System.String"));//0

                qtyTable.Columns.Add("调货信息", System.Type.GetType("System.String"));//0

                List<cls_xiaoshou_info> ALllResults2 = ALllResults.Where((x, ii) => ALllResults.FindIndex(z => z.kuanhao == x.kuanhao && z.kuanming == x.kuanming && z.cima == x.cima) == ii).ToList();//Lambda表达式去重  


                int jk = 0;
                foreach (cls_xiaoshou_info temp in ALllResults2)
                {
                    //返回 datable
                    if (temp.kuanhao == null || temp.kuanhao == "")
                    {


                    }

                    qtyTable.Rows.Add(qtyTable.NewRow());
                 //   qtyTable.Rows.Add(qtyTable.NewRow());

                   // qtyTable.Rows[jk][0] = "销";

                    qtyTable.Rows[jk][0] = temp.kuanhao;
                    qtyTable.Rows[jk][1] = temp.xingbie;
                    qtyTable.Rows[jk][2] = temp.cima;
                    qtyTable.Rows[jk][3] = temp.yanse;
                    qtyTable.Rows[jk][4] = temp.nianfen;
                   // qtyTable.Rows[jk][2] = temp.yanse;
                    //qtyTable.Rows[jk + 1][0] = "库";
                    //qtyTable.Rows[jk + 1][1] = temp.kuanhao;
                    //qtyTable.Rows[jk + 1][2] = temp.xingbie;
                    //qtyTable.Rows[jk + 1][3] = temp.cima;



                    //qtyTable.Rows[jk][9] = "库";
                    //qtyTable.Rows[jk][10] = temp.kuanhao;
                    //qtyTable.Rows[jk][11] = temp.xingbie;
                    //qtyTable.Rows[jk][12] = temp.cima;


                    double xiaoshouhuizong = 0;
                    double kucunhuizong = 0;


                    for (int i = 0; i < farenvalue.Count; i++)
                    {

                        #region 写入各种分店数据
                        //销售
                        List<cls_xiaoshou_info> xiaoshouinfo1 = xiaoshouResults.FindAll(o => o.dianpumingcheng != null && farenvalue[i].Contains(o.dianpumingcheng) && o.kuanhao != null && temp.kuanhao.Contains(o.kuanhao) && o.cima != null && temp.cima.Contains(o.cima) && o.xingbie != null && temp.xingbie.Contains(o.xingbie) && o.danjuleixing != null && o.danjuleixing.Contains("销"));
                        bool ist = false;



                        double nullableQty = (from s in xiaoshouinfo1
                                              where Convert.ToDouble(s.shuliang) > 0
                                              select Convert.ToDouble(s.shuliang)).Sum();


                        cls_diaobo_info item = new cls_diaobo_info();


                        item.kuanhao = temp.kuanhao;
                        item.xingbie = temp.xingbie;
                        item.cima = temp.cima;
                        item.fendianming = farenvalue[i];
                        //销售

                        if (xiaoshouinfo1.Count > 0)
                        {
                            ist = true;
                            item.fendianming_xiaoshou = Convert.ToString(nullableQty);

                        }

                        qtyTable.Rows[jk][5 + i] = item.fendianming_xiaoshou;
                        ///-----


                        //库存
                        List<cls_kucun_info> kucuninfo = kucunResults.FindAll(o => o.cangku != null && farenvalue[i].Contains(o.cangku) && o.kuanhao != null && temp.kuanhao.Contains(o.kuanhao) && o.cima != null && temp.cima.Contains(o.cima) && o.xingbie != null && temp.xingbie.Contains(o.xingbie) && o.danjuleixing != null && o.danjuleixing.Contains("库") && o.yanse != null && temp.yanse.Contains(o.yanse));
                        double nullableQty1 = (from s in kucuninfo
                                               where Convert.ToDouble(s.kucunshuliang) > 0
                                               select Convert.ToDouble(s.kucunshuliang)).Sum();



                        if (kucuninfo.Count > 0)
                        {
                            item.kucun_shengyu = Convert.ToString(nullableQty1);
                        }
                        //  qtyTable.Rows[jk+1][4 + i] = item.kucun_shengyu;
                        qtyTable.Rows[jk][10 + i] = item.kucun_shengyu;
                        //------
                        #endregion

                        xiaoshouhuizong = xiaoshouhuizong + nullableQty;
                        kucunhuizong = kucunhuizong + nullableQty1;

                    }


                    //写入汇总
                    qtyTable.Rows[jk][5 + farenvalue.Count] = Convert.ToString(xiaoshouhuizong);
                    //qtyTable.Rows[jk + 1][10 + farenvalue.Count] = Convert.ToString(kucunhuizong);
                    //qtyTable.Rows[jk + 1][10 + farenvalue.Count + 1] = Convert.ToString(kucunhuizong - xiaoshouhuizong);
                    qtyTable.Rows[jk][10 + farenvalue.Count] = Convert.ToString(kucunhuizong);
               //     qtyTable.Rows[jk][8 + farenvalue.Count + 1] = Convert.ToString(kucunhuizong - xiaoshouhuizong);
                    //-------

                    jk = jk + 1;


                }
                allqtyTable = new DataTable();
                allqtyTable = qtyTable;

                #endregion
            }

           


            foreach (cls_diaobo_info temp in ShowResults)
            {
                if (temp.kuanhao == "E92008H" && temp.kuanming == "轻逸跑鞋" && temp.yanse == "黑色" && temp.cima == "38")
                {

                }
            }

       //     MessageBox.Show(qtyTable.Rows[1][9].ToString());
       
                qtyTable.Rows.Add(qtyTable.NewRow());
                int hz_x = 0;
                int hz_k = 0;
                for (int i = 0; i < qtyTable.Rows.Count-1; i++)
                {
                    if (qtyTable.Rows[i][9] == null) {
                        hz_x += 0;
                        hz_k += 0;
                    }
                    hz_x += Convert.ToInt32(qtyTable.Rows[i][9]);
                    hz_k += Convert.ToInt32(qtyTable.Rows[i][14]);
                }
                qtyTable.Rows[(qtyTable.Rows.Count - 1)][9] = hz_x;
                qtyTable.Rows[(qtyTable.Rows.Count - 1)][14] = hz_k;
           

            return null;

           
        }
        public List<cls_kucun_info> ReadfindngFile(string instertext)
        {

            try
            {
                List<cls_kucun_info> Result = new List<cls_kucun_info>();


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
                    //      bgWorker.ReportProgress(0, "读入数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_kucun_info temp = new cls_kucun_info();

                    #region 基础信息

                    temp.danjuleixing = "";
                    if (o[i, 1] != null)
                        temp.danjuleixing = o[i, 1].ToString().Trim();


                    temp.cangku = "";
                    if (o[i, 2] != null)
                        temp.cangku = o[i, 2].ToString().Trim();

                    temp.kuanhao = "";
                    if (o[i, 3] != null)
                        temp.kuanhao = o[i, 3].ToString().Trim();

                    temp.kuanming = "";
                    if (o[i, 4] != null)
                        temp.kuanming = o[i, 4].ToString().Trim();
                    if (temp.kuanming == "" || temp.kuanming == null)
                        continue;

                    temp.yanse = "";
                    if (o[i, 5] != null)
                        temp.yanse = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp.cima = "";
                    if (o[i, 6] != null)
                        temp.cima = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                    temp.kucunshuliang = "";
                    if (o[i, 7] != null)
                        temp.kucunshuliang = o[i, 7].ToString().Trim();
                    temp.diaopaijia = "";
                    if (o[i, 8] != null)
                        temp.diaopaijia = o[i, 8].ToString().Trim();

                    temp.xingbie = "";
                    if (o[i, 9] != null)
                        temp.xingbie = o[i, 9].ToString().Trim();

                    temp.jijie = "";
                    if (o[i, 10] != null)
                        temp.jijie = o[i, 10].ToString().Trim();

                    temp.leibie = "";
                    if (o[i, 11] != null)
                        temp.leibie = o[i, 11].ToString().Trim();
                    temp.nianfen = "";
                    if (o[i, 12] != null)
                        temp.nianfen = o[i, 12].ToString().Trim();


                    #endregion

                    Result.Add(temp);



                    cls_xiaoshou_info temp1 = new cls_xiaoshou_info();

                    #region 基础信息

                    temp1.danjuleixing = "";
                    if (o[i, 1] != null)
                        temp1.danjuleixing = o[i, 1].ToString().Trim();


                    temp1.dianpumingcheng = "";
                    if (o[i, 2] != null)
                        temp1.dianpumingcheng = o[i, 2].ToString().Trim();

                    temp1.kuanhao = "";
                    if (o[i, 3] != null)
                        temp1.kuanhao = o[i, 3].ToString().Trim();

                    temp1.kuanming = "";
                    if (o[i, 4] != null)
                        temp1.kuanming = o[i, 4].ToString().Trim();
                    if (temp1.kuanming == "" || temp1.kuanming == null)
                        continue;

                    temp1.yanse = "";
                    if (o[i, 5] != null)
                        temp1.yanse = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp1.cima = "";
                    if (o[i, 6] != null)
                        temp1.cima = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                    temp1.shuliang = "";
                    if (o[i, 7] != null)
                        temp1.shuliang = o[i, 7].ToString().Trim();
                    temp1.diaopaijia = "";
                    if (o[i, 8] != null)
                        temp1.diaopaijia = o[i, 8].ToString().Trim();

                    temp1.xingbie = "";
                    if (o[i, 9] != null)
                        temp1.xingbie = o[i, 9].ToString().Trim();

                    temp1.jijie = "";
                    if (o[i, 10] != null)
                        temp1.jijie = o[i, 10].ToString().Trim();

                    temp1.leibie = "";
                    if (o[i, 11] != null)
                        temp1.leibie = o[i, 11].ToString().Trim();
                    temp1.nianfen = "";
                    if (o[i, 12] != null)
                        temp1.nianfen = o[i, 12].ToString().Trim();


                    temp1.nianfen = "库存表";


                    #endregion

                    ALllResults.Add(temp1);

                    //  xiaoshouResults.Add(temp1);
                }



                return Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show("表格存在异常,请参照原始表格格式修改:" + ex.Message);

                throw ex;
            }

        }
        public List<cls_xiaoshou_info> ReadxiaoshouFile(string instertext)
        {

            try
            {
                List<cls_xiaoshou_info> Result = new List<cls_xiaoshou_info>();


                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(instertext, Type.Missing, true, Type.Missing,
                    "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["Sheet1"];
                Microsoft.Office.Interop.Excel.Range rng;
                rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
                int rowCount = WS.UsedRange.Rows.Count - 1;
                object[,] o = new object[1, 1];
                o = (object[,])rng.Value2;
                clsCommHelp.CloseExcel(excelApp, analyWK);

                for (int i = 1; i <= rowCount; i++)
                {
                    //  bgWorker.ReportProgress(0, "读入数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                    cls_xiaoshou_info temp = new cls_xiaoshou_info();

                    #region 基础信息

                    temp.danjuleixing = "";
                    if (o[i, 1] != null)
                        temp.danjuleixing = o[i, 1].ToString().Trim();


                    temp.dianpumingcheng = "";
                    if (o[i, 2] != null)
                        temp.dianpumingcheng = o[i, 2].ToString().Trim();

                    temp.kuanhao = "";
                    if (o[i, 3] != null)
                        temp.kuanhao = o[i, 3].ToString().Trim();

                    temp.kuanming = "";
                    if (o[i, 4] != null)
                        temp.kuanming = o[i, 4].ToString().Trim();
                    if (temp.kuanming == "" || temp.kuanming == null)
                        continue;

                    temp.yanse = "";
                    if (o[i, 5] != null)
                        temp.yanse = o[i, 5].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 5]);
                    temp.cima = "";
                    if (o[i, 6] != null)
                        temp.cima = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                    temp.shuliang = "";
                    if (o[i, 7] != null)
                        temp.shuliang = o[i, 7].ToString().Trim();
                    temp.diaopaijia = "";
                    if (o[i, 8] != null)
                        temp.diaopaijia = o[i, 8].ToString().Trim();

                    temp.xingbie = "";
                    if (o[i, 9] != null)
                        temp.xingbie = o[i, 9].ToString().Trim();

                    temp.jijie = "";
                    if (o[i, 10] != null)
                        temp.jijie = o[i, 10].ToString().Trim();

                    temp.leibie = "";
                    if (o[i, 11] != null)
                        temp.leibie = o[i, 11].ToString().Trim();
                    temp.nianfen = "";
                    if (o[i, 12] != null)
                        temp.nianfen = o[i, 12].ToString().Trim();





                    #endregion

                    temp.beizhu1 = "销售表";

                    ALllResults.Add(temp);


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

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

            int s = this.tabControl1.SelectedIndex;

            if (s == 0)
            {
                downall(this.dataGridView3);

            }
            else if (s == 1)
            {

                downall(this.dataGridView4);

            }
            else if (s == 2)
            {

                downall(this.dataGridView2);

            }
            else if (s == 3)
            {

                downall(this.dataGridView1);

            }
        }

        private void downall(DataGridView dataGridView2)
        {
            if (dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\t";
            string strHeader = "";
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                strHeader += dataGridView2.Columns[i].HeaderText + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < dataGridView2.Rows.Count; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < dataGridView2.Columns.Count; k++)
                {
                    if (dataGridView2.Rows[j].Cells[k].Value != null)
                        strRowValue += dataGridView2.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ") + delimiter;
                    else
                        strRowValue += dataGridView2.Rows[j].Cells[k].Value + delimiter;
                }
                sw.WriteLine(strRowValue);
            }

            sw.Close();
            fa.Close();
            MessageBox.Show("下载完成！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            combox3change();
        }

        private void combox3change()
        {
            //if (comboBox3.Text.Length > 0 && comboBox3.Text.Contains("全部"))
            //{
            //    davshow(ShowResults);

            //}

            //else
            //{
            //    FilterShowResults = new List<cls_diaobo_info>();

            //    FilterShowResults = this.ShowResults.Where(s => s.kuanhao != null && s.kuanhao.ToString().StartsWith(comboBox3.Text.ToString())).ToList();

            //    davshow(FilterShowResults);

            //}
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            combox3change();
        }

        private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void checkedListBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            List<string> alist = new List<string>();
            //if (this.checkedListBox1.CheckedItems.Count > 0)
            //{
            //    foreach (string status in this.checkedListBox1.CheckedItems)
            //    {
            //        alist.Add(status);
            //    }
            //    //选择显示列
            //}

            //for (int j = 0; j < dataGridView.ColumnCount; j++)
            //{
            //    dataGridView.Columns[j].Visible = true;

            //}
            //for (int i = 0; i < alist.Count; i++)
            //{
            //    for (int j = 0; j < dataGridView.ColumnCount; j++)
            //    {
            //        string MU = dataGridView.Columns[j].HeaderText;
            //        if (MU == alist[i])
            //            dataGridView.Columns[j].Visible = false;
            //        else
            //        {

            //        }
            //    }
            //}
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            if (path2 == null || path2 == "" || path == null || path == "")
            {
                MessageBox.Show("请选择路径！");

                return;

            }
            kucunResults = new List<cls_kucun_info>();


            kucunResults = Buiness_Bankcharge(ref this.bgWorker, "A", "", "");


            InitializeDataGridView();

            davshow(ShowResults);
        }

        private void checkedListBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            List<string> alist = new List<string>();
            //if (this.checkedListBox2.CheckedItems.Count > 0)
            //{
            //    foreach (string status in this.checkedListBox2.CheckedItems)
            //    {
            //        alist.Add(status);
            //    }
            //    //选择显示列
            //}

            for (int j = 0; j < dataGridView3.ColumnCount; j++)
            {
                dataGridView3.Columns[j].Visible = true;

            }
            for (int i = 0; i < alist.Count; i++)
            {
                for (int j = 0; j < dataGridView3.ColumnCount; j++)
                {
                    string MU = dataGridView3.Columns[j].HeaderText;
                    if (MU == alist[i])
                        dataGridView3.Columns[j].Visible = false;
                    else
                    {

                    }
                }
            }
        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
         
            if (Convert.ToInt32(dataGridView3.CurrentCell.ColumnIndex.ToString()) >= 3 && Convert.ToInt32(dataGridView3.CurrentCell.ColumnIndex.ToString()) <= 3 + f_count - 1)
            {
               
          
           
      //      MessageBox.Show(dataGridView3.CurrentCell.ColumnIndex.ToString());
            string[] cf_dm = dataGridView3.Columns[dataGridView3.CurrentCell.ColumnIndex].HeaderText.Split(new char[] { '(' });         
            string chageva = "";
            string dian_pu = "";
            var form = new frmChangequilty();
            form.ayy = mdzz;
            if (form.ShowDialog() == DialogResult.OK)
            {
                
                chageva = form.txt;
                dian_pu = form.cob;
            }
          
            if (chageva == null || chageva == "")
                return;

            string sls = dataGridView3.Rows[RowRemark].Cells[cloumn].EditedFormattedValue.ToString();
            string MU = dataGridView3.Columns[cloumn].HeaderText;
            string value = dataGridView3.Rows[RowRemark].Cells["调货信息"].EditedFormattedValue.ToString();
            dataGridView3.Rows[RowRemark].Cells["调货信息"].Value = value + "-" + MU + "[" + chageva + "]";
            try
            {
             //   dataGridView3.Rows[RowRemark].Cells["剩余"].Value = Convert.ToInt32(sls) - Convert.ToInt32(chageva);
            }
            catch (FormatException)
            {
                MessageBox.Show("库存异常");
            }
           
         // shib = new DataTable();
           // int i = dataGridView3.CurrentCell.ColumnIndex;
           

          //  MessageBox.Show(dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[i].HeaderText].Value.ToString());
            
            for (int u = 0; u < cc; u++) 
            {
                shib.Rows.Add(shib.NewRow());
            }

            shib.Rows[cc][0] = dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[0].HeaderText].Value.ToString();

            shib.Rows[cc][1] = cf_dm[0];

            shib.Rows[cc][2] = dian_pu;
            if (Regex.Matches(dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[2].HeaderText].Value.ToString().ToString(), "[a-zA-Z]").Count <= 0 && HasChineseTest(dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[2].HeaderText].Value.ToString()) == false)
            {
                for (int cm = 35; cm <= 48; cm++)
                {
                    if (Convert.ToInt32(dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[2].HeaderText].Value.ToString()) >= 35 && Convert.ToInt32(dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[2].HeaderText].Value.ToString()) <= 48)
                    {


                        if (dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[2].HeaderText].Value.ToString() == cm.ToString())
                        {
                            shib.Rows[cc][cm - 32] = chageva;
                        }
                    }
                    else
                    {
                        MessageBox.Show("尺码异常!");
                        break;
                    }
                }

            }
            else 
            {
                string shibi_xh = dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[dataGridView3.Columns[2].HeaderText].Value.ToString().ToUpper();
                if (shibi_xh == "S")
                {
                    shib.Rows[cc][3] = chageva;
                }
                else if (shibi_xh == "M")
                {
                    shib.Rows[cc][4] = chageva;
                }
                else if (shibi_xh == "L")
                {
                    shib.Rows[cc][5] = chageva;
                }
                else if (shibi_xh == "XL")
                {
                    shib.Rows[cc][6] = chageva;
                }
                else if (shibi_xh == "2XL")
                {
                    shib.Rows[cc][7] = chageva;
                }
                else if (shibi_xh == "3XL")
                {
                    shib.Rows[cc][8] = chageva;
                }
                else if (shibi_xh == "4XL")
                {
                    shib.Rows[cc][9] = chageva;
                }
                else if (shibi_xh == "5XL")
                {
                    shib.Rows[cc][10] = chageva;
                }
                else if (shibi_xh == "6XL")
                {
                    shib.Rows[cc][11] = chageva;
                }
                else if (shibi_xh == "7XL")
                {
                    shib.Rows[cc][12] = chageva;
                }
                else if (shibi_xh == "8XL")
                {
                    shib.Rows[cc][13] = chageva;
                }
                else if (shibi_xh == "9XL")
                {
                    shib.Rows[cc][14] = chageva;
                }
                else if (shibi_xh == "10XL")
                {
                    shib.Rows[cc][15] = chageva;
                }
                else if (shibi_xh == "11XL")
                {
                    shib.Rows[cc][16] = chageva;
                }
                else 
                {
                    MessageBox.Show("尺码异常!");
                  
                }
            }
           

           
            cc += 1;
           
            dataGridView4.DataSource = shib;
            ///////////////////////////////////////////////////////
            }
            else
            {
                MessageBox.Show("请选择对应凋拨店铺");
            }
        }

        public bool HasChineseTest(string text)
        {
            //string text = "是不是汉字，ABC,keleyi.com";
            char[] c = text.ToCharArray();
            bool ischina = false;

            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                {
                    ischina = true;

                }
                else
                {
                    //  ischina = false;
                }
            }
            return ischina;

        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            RowRemark = e.RowIndex;
            cloumn = e.ColumnIndex;
           
        }

        private void 清空单元格ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView3.Rows[RowRemark].Cells[cloumn].Value = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            combox4change();
        }

        private void filterButton_Click(object sender, EventArgs e)
        {
            combox3change();
        }

        private void combox4change()
        {
            if (comboBox4.Text.Length > 0 && comboBox4.Text.Contains("全部") && comboBox5.Text.Length > 0 && comboBox5.Text.Contains("全部") && comboBox6.Text.Length > 0 && comboBox6.Text.Contains("全部"))
            {
                //dav3(allqtyTable);
                this.bindingSource2.Filter = "";
                comboBox4.SelectedIndex = 0;
                comboBox5.SelectedIndex = 0;
                comboBox6.SelectedIndex = 0;
                this.dataGridView3.DataSource = this.bindingSource2;
                this.dataGridView3.Refresh();
            }

            else
            {
                ApplyFilter();


            }
        }
        private void ApplyFilter()
        {
            string filter = "";
            if (this.comboBox4.Text.Length > 0 && !comboBox4.Text.Contains("全部"))
            {
                filter += "(款号='" + this.comboBox4.Text + "')";
            }
            if (this.comboBox5.Text.Length > 0 && !comboBox5.Text.Contains("全部"))
            {
                if (filter.Length > 0)
                {
                    filter += " and ";
                }
                filter += "(尺码='" + this.comboBox5.Text + "')";
            }

            if (this.comboBox6.Text.Length > 0 && !comboBox6.Text.Contains("全部"))
            {
                if (filter.Length > 0)
                {
                    filter += " and ";
                }
                filter += "(性别='" + this.comboBox6.Text + "')";
            }


            this.bindingSource2.Filter = filter;
            this.dataGridView3.DataSource = this.bindingSource2;

        }

        private void combox1change()
        {
            if (comboBox1.Text.Length > 0 && comboBox1.Text.Contains("全部"))
            {

                dav2show(xiaoshouResults);
            }

            else
            {
                shaixuan_xiaoshouResults = new List<cls_xiaoshou_info>();


                shaixuan_xiaoshouResults = this.xiaoshouResults.Where(s => s.kuanhao != null && s.kuanhao.ToString().StartsWith(comboBox1.Text.ToString())).ToList();

                dav2show(shaixuan_xiaoshouResults);

            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            combox1change();

        }
        private void dav2show(List<cls_xiaoshou_info> ShowResults)
        {


            //销售 显示

            xiaoshou_sortabledinningsOrderList = new China_System.Common.clsCommHelp.SortableBindingList<cls_xiaoshou_info>(ShowResults);
            this.bindingSource3.DataSource = this.xiaoshou_sortabledinningsOrderList;

            this.dataGridView2.AutoGenerateColumns = false;
            this.dataGridView2.DataSource = this.bindingSource3;

            if (dataGridView2.DataSource != null)
                toolStripLabel1.Text = "条数： " + dataGridView2.RowCount;
        }

        private void combox2change()
        {
            if (comboBox2.Text.Length > 0 && comboBox2.Text.Contains("全部"))
            {

                dav3show(kucunResults);
            }

            else
            {
                kuncun_xiaoshouResults = new List<cls_kucun_info>();


                kuncun_xiaoshouResults = this.kucunResults.Where(s => s.kuanhao != null && s.kuanhao.ToString().StartsWith(comboBox2.Text.ToString())).ToList();

                dav3show(kuncun_xiaoshouResults);

            }
        }
        private void dav3show(List<cls_kucun_info> ShowResults)
        {
            //销售 显示

            kucun_sortabledinningsOrderList = new China_System.Common.clsCommHelp.SortableBindingList<cls_kucun_info>(ShowResults);
            this.bindingSource4.DataSource = this.kucun_sortabledinningsOrderList;

            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = this.bindingSource4;

            if (dataGridView1.DataSource != null)
                toolStripLabel1.Text = "条数： " + dataGridView1.RowCount;
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            combox1change();

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            combox2change();


        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            combox2change();

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int s = this.tabControl1.SelectedIndex;
            if (s == 0)
            {
                //if (dataGridView.DataSource != null)
                //    toolStripLabel1.Text = "条数： " + dataGridView.RowCount;

            }
            else if (s == 1)
            {
                if (dataGridView3.DataSource != null)
                    toolStripLabel1.Text = "条数： " + dataGridView3.RowCount;


            }
            else if (s == 2)
            {
                if (dataGridView2.DataSource != null)
                    toolStripLabel1.Text = "条数： " + dataGridView2.RowCount;


            }
            else if (s == 3)
            {
                if (dataGridView1.DataSource != null)
                    toolStripLabel1.Text = "条数： " + dataGridView1.RowCount;


            }

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void 隐藏ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            try
            {
                int i = dataGridView3.CurrentCell.ColumnIndex;
                List<string> alist = new List<string>();
                dataGridView3.Columns[i].Visible = false;
            }
            catch (NullReferenceException)
            {

                MessageBox.Show("请选中列");
            }
                          
         }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int j = 0; j < dataGridView3.ColumnCount; j++)
            {
                dataGridView3.Columns[j].Visible = true;

            }
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        
    }
}
