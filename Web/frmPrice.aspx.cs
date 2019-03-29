using China_System.Common;
using clsBuiness;
using SDZdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.SessionState;
using System.Web.UI.HtmlControls;

namespace Web
{
    public partial class frmPrice : System.Web.UI.Page
    {
        public string alterinfo;
        List<cls_order_info> readCards;
        public string user;
        public string pass;
        private SortableBindingList<cls_order_info> sortablePendingOrderList;
        object ddd;
        public string Show_infomation;
        bool ischeck_zhengjianhaoma = true;
        clsAllnew BusinessHelp;
        private string servename;
        List<cls_order_info> t_Item_info;

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

                bind();
            }
            if (IsPostBack)
                bind();
        }
        public class SortableBindingList<T> : BindingList<T>
        {
            private bool isSortedCore = true;
            private ListSortDirection sortDirectionCore = ListSortDirection.Ascending;
            private PropertyDescriptor sortPropertyCore = null;
            private string defaultSortItem;

            public SortableBindingList() : base() { }

            public SortableBindingList(IList<T> list) : base(list) { }

            protected override bool SupportsSortingCore
            {
                get { return true; }
            }

            protected override bool SupportsSearchingCore
            {
                get { return true; }
            }

            protected override bool IsSortedCore
            {
                get { return isSortedCore; }
            }

            protected override ListSortDirection SortDirectionCore
            {
                get { return sortDirectionCore; }
            }

            protected override PropertyDescriptor SortPropertyCore
            {
                get { return sortPropertyCore; }
            }

            protected override int FindCore(PropertyDescriptor prop, object key)
            {
                for (int i = 0; i < this.Count; i++)
                {
                    if (Equals(prop.GetValue(this[i]), key)) return i;
                }
                return -1;
            }

            protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
            {
                isSortedCore = true;
                sortPropertyCore = prop;
                sortDirectionCore = direction;
                Sort();
            }

            protected override void RemoveSortCore()
            {
                if (isSortedCore)
                {
                    isSortedCore = false;
                    sortPropertyCore = null;
                    sortDirectionCore = ListSortDirection.Ascending;
                    Sort();
                }
            }

            public string DefaultSortItem
            {
                get { return defaultSortItem; }
                set
                {
                    if (defaultSortItem != value)
                    {
                        defaultSortItem = value;
                        Sort();
                    }
                }
            }

            private void Sort()
            {
                List<T> list = (this.Items as List<T>);
                list.Sort(CompareCore);
                ResetBindings();
            }

            private int CompareCore(T o1, T o2)
            {
                int ret = 0;
                if (SortPropertyCore != null)
                {
                    ret = CompareValue(SortPropertyCore.GetValue(o1), SortPropertyCore.GetValue(o2), SortPropertyCore.PropertyType);
                }
                if (ret == 0 && DefaultSortItem != null)
                {
                    PropertyInfo property = typeof(T).GetProperty(DefaultSortItem, BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.IgnoreCase, null, null, new Type[0], null);
                    if (property != null)
                    {
                        ret = CompareValue(property.GetValue(o1, null), property.GetValue(o2, null), property.PropertyType);
                    }
                }
                if (SortDirectionCore == ListSortDirection.Descending) ret = -ret;
                return ret;
            }

            private static int CompareValue(object o1, object o2, Type type)
            {
                if (o1 == null) return o2 == null ? 0 : -1;
                else if (o2 == null) return 1;
                else if (type.IsPrimitive || type.IsEnum) return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
                else if (type == typeof(DateTime)) return Convert.ToDateTime(o1).CompareTo(o2);
                else return String.Compare(o1.ToString().Trim(), o2.ToString().Trim());
            }
        }

        public void bind()
        {

            btReadcard_server_Click(null, EventArgs.Empty);
            InitialSystemInfo();
        }
        public string clear_bind()
        {

            //gvList.DataSource = readCards;

            //gvList.DataBind();
            readCards = new List<cls_order_info>();

            InitialSystemInfo();

            this.txtCompletionTime.Text = "";
            this.txrearchNAME.Text = "";

            return "ok";

        }
        public static string ajaxclear_bind()
        {
            List<cls_order_info> readCards = new List<cls_order_info>();

            //this.gvList.AutoGenerateColumns = false;
            //sortablePendingOrderList = new SortableBindingList<cls_order_info>(readCards);
            ////this.bindingSource1.DataSource = sortablePendingOrderList;
            //this.gvList.DataSource = sortablePendingOrderList;
            //gvList.DataKeyNames = new string[] { "Order_id" };//主键
            // this.gvList.DataBind();
            //Show_infomation = "共计 " + sortablePendingOrderList.Count() + " 条";



            return "ok";


        }

        protected void GridView_OnRowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "Btn_Operation")
            {
                int RowRemark = Convert.ToInt32(e.CommandArgument);
                if (RowRemark >= 0)
                {
                    ////string QiHao = gvList.Rows[RowRemark].Cells[1].Text.ToString();
                    //BusinessHelp = new clsAllnew();
                    //BusinessHelp.rev_servename = servename;
                    //gohome();

                    //string QiHao = gvList.DataKeys[RowRemark].Value.ToString();
                    //string sql2 = "delete from t_Item_3002 where   FItemID='" + QiHao + "'";

                    ////BusinessHelp.Readt_PICServer(sql2);
                    //BusinessHelp.deleteCard(sql2);

                    ////删 t_Item

                    //sql2 = "delete from t_Item where   FItemID='" + QiHao + "'";

                    //BusinessHelp.deleteCard(sql2);



                    //Page.ClientScript.RegisterStartupScript(this.GetType(), "", "<script>alert('删除成功')</script>");
                    //btReadcard_server_Click(null, EventArgs.Empty);

                    //InitialSystemInfo();


                }
            }
            else if (e.CommandName == "Btn_View")
            {
                int RowRemark = Convert.ToInt32(e.CommandArgument);
                if (RowRemark >= 0)
                {

                }

            }
        }

        private void gohome()
        {
            string ex = "网络不畅通，请检查数据库连接配置是否正确";

            if (BusinessHelp.ConStr == null || BusinessHelp.ConStr == null)
                //Response.Redirect("~/Myadmin/login.aspx");
                HttpContext.Current.Response.Redirect("~/ErrorPage/ErrorPage.aspx?Error=" + ex.ToString());

        }
        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {
            this.gvList.EditIndex = e.NewEditIndex;
            //使编辑的行是当前操作的行   editIndex:编辑行的索引  newEditIndex:所编辑的行的索引
            bind();

        }

        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            gvList.EditIndex = -1;
            bind();
        }
        protected void GridView_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            //e.Row.Attributes.Add("style", "height:43px Width:43px");
            if (e.Row.RowState == DataControlRowState.Edit || e.Row.RowState == DataControlRowState.Alternate || e.Row.RowState == DataControlRowState.Edit)
            {
                TextBox curText;
                for (int i = 0; i < e.Row.Cells.Count; i++)
                {
                    if (e.Row.Cells[i].Controls.Count != 0)
                    {
                        curText = e.Row.Cells[i].Controls[0] as TextBox;
                        if (curText != null)
                        {
                            curText.Width = Unit.Pixel(10);
                        }
                    }
                }
            }
            if (e.Row.RowType == DataControlRowType.DataRow)
            {

                //鼠标经过时，行背景色变

                e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='#ffd800'");

                //鼠标移出时，行背景色变

                e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='#FFFFFF'");

            }
        }
        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            // if (!IsPostBack)
            {
                BusinessHelp = new clsAllnew();


                List<cls_order_info> AddreadCards = new List<cls_order_info>();
                cls_order_info item = new cls_order_info();

                item.PotNo = ((TextBox)(gvList.Rows[e.RowIndex].Cells[0].Controls[0])).Text.ToString().Trim();
                item.DDate = ((TextBox)(gvList.Rows[e.RowIndex].Cells[1].Controls[0])).Text.ToString().Trim();
                item.AlCnt = ((TextBox)(gvList.Rows[e.RowIndex].Cells[2].Controls[0])).Text.ToString().Trim();
                item.Lsp = ((TextBox)(gvList.Rows[e.RowIndex].Cells[3].Controls[0])).Text.ToString().Trim();
                item.Djzsp = ((TextBox)(gvList.Rows[e.RowIndex].Cells[4].Controls[0])).Text.ToString().Trim();
                item.Djwd = ((TextBox)(gvList.Rows[e.RowIndex].Cells[5].Controls[0])).Text.ToString().Trim();
                item.Fzb = ((TextBox)(gvList.Rows[e.RowIndex].Cells[6].Controls[0])).Text.ToString().Trim();
                item.FeCnt = ((TextBox)(gvList.Rows[e.RowIndex].Cells[7].Controls[0])).Text.ToString().Trim();
                item.SiCnt = ((TextBox)(gvList.Rows[e.RowIndex].Cells[8].Controls[0])).Text.ToString().Trim();
                item.AlOCnt = ((TextBox)(gvList.Rows[e.RowIndex].Cells[9].Controls[0])).Text.ToString().Trim();
                item.CaFCnt = ((TextBox)(gvList.Rows[e.RowIndex].Cells[10].Controls[0])).Text.ToString().Trim();
                item.MgCnt = ((TextBox)(gvList.Rows[e.RowIndex].Cells[11].Controls[0])).Text.ToString().Trim();
                item.LDYJ = ((TextBox)(gvList.Rows[e.RowIndex].Cells[12].Controls[0])).Text.ToString().Trim();
                item.MLsp = ((TextBox)(gvList.Rows[e.RowIndex].Cells[13].Controls[0])).Text.ToString().Trim();
                item.LPW = ((TextBox)(gvList.Rows[e.RowIndex].Cells[14].Controls[0])).Text.ToString().Trim();

                item.Order_id = gvList.DataKeys[e.RowIndex].Values[0].ToString();
                AddreadCards.Add(item);

                BusinessHelp.changeCardServer(AddreadCards);

                gvList.EditIndex = -1;
                bind();
            }
        }



        protected void GridView1_DataBound(object sender, EventArgs e)
        {
            // 演示ToolTip，使用GridView自带的ToolTip
            for (int i = 0; i < gvList.Rows.Count; i++)
            {
                if (gvList.Rows[i].Cells[3].Text == "990112")
                    gvList.Rows[i].Cells[3].Text = "男";
                else if (gvList.Rows[i].Cells[3].Text == "990113")
                    gvList.Rows[i].Cells[3].Text = "女";

                gvList.Rows[i].Cells[8].ToolTip = gvList.Rows[i].Cells[8].Text;
                if (gvList.Rows[i].Cells[8].Text.Length > 4 && gvList.Rows[i].Cells[8].Text != "&nbsp;")
                    gvList.Rows[i].Cells[8].Text = gvList.Rows[i].Cells[8].Text.Substring(0, 4) + "...";
            }
        }


        protected void Button1_Click(object sender, EventArgs e)
        {

        }

        private void InitialSystemInfo()
        {
            if (readCards == null)
                readCards = new List<cls_order_info>();

            this.gvList.AutoGenerateColumns = false;
            sortablePendingOrderList = new SortableBindingList<cls_order_info>(readCards);

            this.gvList.DataSource = sortablePendingOrderList;
            gvList.DataKeyNames = new string[] { "Order_id" };//主键
            this.gvList.DataBind();
            Show_infomation = "共计 " + sortablePendingOrderList.Count() + " 条";
        }

        protected void btwrite_Click(object sender, EventArgs e)
        {
            BusinessHelp = new clsAllnew();
            BusinessHelp.rev_servename = servename;
            gohome();
            //if (readCards != null && readCards.Count > 0)
            //{
            //    string sql2 = "delete from NewMeasueDataTable where   PotNo='" + readCards[0].PotNo + "'";

            //    BusinessHelp.deleteCard(sql2);
            //}
            BusinessHelp.createICcard_info_Server(readCards);


            // alterinfo = "添加用户成功！";
        }


        private void MyGo()
        {
            Button1_Click(null, EventArgs.Empty);
            this.hidden1.Value = "2";

        }
        private void CreateITEM_CARD_Server(clsAllnew BusinessHelp)
        {

        }
        public static void ShowConfirm(string strMsg, string strUrl_Yes, string strUrl_No)
        {
            System.Web.HttpContext.Current.Response.Write("<Script Language='JavaScript'>if ( window.confirm('" + strMsg + "')) { window.location.href='" + strUrl_Yes +
            "' } else {window.location.href='" + strUrl_No + "' };</script>");
        }
        protected void btReadcard_server_Click(object sender, EventArgs e)
        {
            BusinessHelp = new clsAllnew();
            BusinessHelp.rev_servename = servename;
            gohome();
            readCards = new List<cls_order_info>();
            //string conditions = "select * from NewMeasueDataTable";//成功

            int xuliehao = 0;

            xuliehao = Select_xuhaolei(xuliehao);

            string conditions = "select * from NewMeasueDataTable where ";//成功

            conditions = SelectSql(xuliehao, conditions);

            if (!conditions.Contains("DDate"))
                return;

            readCards = BusinessHelp.Readt_ItemServer(conditions);

            InitialSystemInfo();
        }

        protected void button2_Click(object sender, EventArgs e)
        {
            clear_bind();
        }

        void toExcel(GridView gv)
        {
            Response.Charset = "GB2312";
            Response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");

            //string fileName = "export.xls";
            string fileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            string style = @"<style> .text { mso-number-format:\@; } </script> ";
            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=" + fileName);
            Response.ContentType = "application/excel";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            gv.RenderControl(htw);
            Response.Write(style);
            Response.Write(sw.ToString());
            Response.End();

        }
        public override void VerifyRenderingInServerForm(Control control)
        {

        }


        protected void toExcel(object sender, EventArgs e)
        {
            //bind();
            //toExcel(gvList);
            Response.Clear();
            string fileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

            Response.AddHeader("content-disposition",

            "attachment;filename=" + fileName);

            Response.Charset = "";

            // If you want the option to open the Excel file without saving than

            // comment out the line below

            // Response.Cache.SetCacheability(HttpCacheability.NoCache);

            Response.ContentType = "application/vnd.xlsx";

            System.IO.StringWriter stringWrite = new System.IO.StringWriter();

            System.Web.UI.HtmlTextWriter htmlWrite =
            new HtmlTextWriter(stringWrite);

            // turn off paging 
            gvList.AllowPaging = false;
            bind();


            gvList.RenderControl(htmlWrite);

            Response.Write(stringWrite.ToString());

            Response.End();

            // turn the paging on again 
            gvList.AllowPaging = true;
            bind();

        }
        [System.Web.Services.WebMethod()]

        private static void inputlog(string aainput)
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "bin\\log.txt";
            if (File.Exists(A_Path))
            {
                StreamWriter sw = new StreamWriter(A_Path);
                sw.WriteLine(aainput);
                sw.Flush();
                sw.Close();
            }
        }
        private static void inputlog1(string aainput)
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "bin\\log2.txt";
            if (File.Exists(A_Path))
            {
                StreamWriter sw = new StreamWriter(A_Path);
                sw.WriteLine(aainput);
                sw.Flush();
                sw.Close();
            }
        }
        private static void gohome1()
        {
            clsAllnew BusinessHelp = new clsAllnew();

            //if (BusinessHelp.ConStr == null || BusinessHelp.ConStr == null)
            //    Response.Redirect("~/Myadmin/login.aspx");
        }

        protected void btsearch_Click1(object sender, EventArgs e)
        {
            if (this.txrearchNAME.Text != "" && this.txtCompletionTime.Text != "")
            {
                BusinessHelp = new clsAllnew();
                BusinessHelp.rev_servename = servename;
                gohome();
                readCards = new List<cls_order_info>();

                int xuliehao = 0;

                xuliehao = Select_xuhaolei(xuliehao);

                string conditions = "select * from NewMeasueDataTable where ";//成功

                conditions = SelectSql(xuliehao, conditions);

                readCards = BusinessHelp.Readt_ItemServer(conditions);
                if (readCards != null && readCards.Count > 1)
                {
                    alterinfo = "本区信息已经数据库存在，无法添加！";
                    //readCards = new List<cls_order_info>();

                }
                else
                {


                    xuliehao = Select_xuhaolei(xuliehao);
                    if (xuliehao > 0)
                    {
                        for (int i = 0; i < 38; i++)
                        {
                            cls_order_info item = new cls_order_info();
                            xuliehao++;
                            item.PotNo = xuliehao.ToString();
                            item.DDate = txtCompletionTime.Text.ToString();

                            readCards.Add(item);
                        }
                        btwrite_Click(null, EventArgs.Empty);
                    }

                }

                InitialSystemInfo();
            }
        }

        private string SelectSql(int xuliehao, string conditions)
        {
            int DSD = 0;

            if (txrearchNAME.Text.Length > 0)
            {
                DSD++;
                conditions += "PotNo like '" + xuliehao.ToString().Substring(0, 2) + "%'";
            }
            if (txtCompletionTime.Text.Length > 0 && DSD > 0)
            {
                //  conditions += " AND DDate like '%" + txtCompletionTime.Text + "%'";
                conditions += " AND  convert(varchar(10),[DDate],120) = '" + txtCompletionTime.Text + "'";

            }
            if (txtCompletionTime.Text.Length > 0 && DSD == 0)
            {
                // conditions += "DDate like '%" + txtCompletionTime.Text + "%'";
            }
            return conditions;
        }

        private int Select_xuhaolei(int xuliehao)
        {
            string changindex = txrearchNAME.Text;
            if (changindex.Contains("一"))
                xuliehao = 1100;
            else if (changindex.Contains("二"))
                xuliehao = 1200;
            else if (changindex.Contains("三"))
                xuliehao = 1300;
            else if (changindex.Contains("四"))
                xuliehao = 1400;
            else if (changindex.Contains("五"))
                xuliehao = 1500;
            else if (changindex.Contains("六"))
                xuliehao = 1600;
            return xuliehao;
        }


        protected void GridView1_RowCreated(object sender, GridViewRowEventArgs e)
        {
            //for (int i = 0; i < e.Row.Cells.Count; i++)
            //{
            //    e.Row.Cells[i].Attributes.Add("style", "word-break :keep-all ; word-wrap:keep-all");

            //}

        }


    }
}