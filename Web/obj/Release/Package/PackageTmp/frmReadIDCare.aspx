<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeBehind="frmReadIDCare.aspx.cs" Inherits="Web.frmReadIDCare" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" class="trbackcolor">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>数据单号录入</title>
    <link href="../../Myadmin/css/common.css" rel="stylesheet" type="text/css" />

    <%--    <a href="./Myadmin/login.aspx">
        <input type="text" size="26" style="font-size: 16pt; border-style: none" value="      首页>数据录入" /></a>--%>


    <script src="/Scripts/My97DatePicker/WdatePicker.js" type="text/javascript"></script>

    <script src="/Myadmin/js/jquery-1.7.1.min.js" type="text/javascript"></script>
    <script src="/Myadmin/js/json2.js" type="text/javascript"></script>
    <script type="text/javascript">
        function MyConfirm() {
            if (confirm("号已存在,确定要继续吗?") == true) {
                document.getElementById("hidden1").value = "1";
            }
            else {
                document.getElementById("hidden1").value = "0";
            }
            form1.submit();
        }
        function btsearchcheck() {
            if (document.form1.txrearchNAME.value == "" || document.form1.txtCompletionTime.value == "") {
                alert("请输入完整信息！");
                document.form1.txrearchNAME.focus();

                return false
            }
        }
        function ClearData() {

        }
        function nvhome() {

            window.location.href = "frmReadIDCare.aspx";
        }



        function ReadCard() {
            ClearData();
            CertCtl.IsRepeat(false);

            var result = CertCtl.ReadCard();
            var imagel = CertCtl.ExportJPGCardB();
            //var imagel1 = CertCtl.ExportJPGCardF();
            var imagelall = CertCtl.ExportJPGCardV();

            var errosinfo = '';
            var idResultDesc1 = '';
            if (result == "0") {

                errosinfo = "成功";
                idResultDesc1 = "读卡成功";
                //  alert("读卡成功+");

            }
            else {

                errosinfo = "失败";
                idResultDesc1 = "读卡失败";
                //    alert("读卡失败+");
            }


            var postData = { mingcheng: CertCtl.Name, minzu: CertCtl.Nation, xingbie: CertCtl.Sex, chushengriqi: CertCtl.Born, jiatingzhuzhi: CertCtl.Address, zhengjianhaoma: CertCtl.CardNo, zhengjianyouxiao: CertCtl.ExpiredDate, FData: CertCtl.GetJPGCardVBase64(), FDataF: CertCtl.GetJPGCardBBase64(), idResult: errosinfo };//, idResultDesc: idResultDesc1   CertCtl.EffectedDate + "-" + 

            $.ajax({
                type: "post", //要用post方式                 
                url: "frmReadIDCare.aspx/GetRankedUserDept",//方法所在页面和方法名
                contentType: "application/json; charset=utf-8",
                async: false,
                dataType: "json",
                data: JSON.stringify(postData),
                success: function (data) {

                    alert(data.d);//返回的数据用data.d获取内容
                    //  window.location.reload();
                },
                error: function (err) {
                    alert(err);
                }
            });
            //  __doPostBack('bind', '');
        }
        function cleardata() {

            var date = new Date();
            var seperator1 = "-";
            var year = date.getFullYear();
            var month = date.getMonth() + 1;
            var strDate = date.getDate();
            if (month >= 1 && month <= 9) {
                month = "0" + month;
            }
            if (strDate >= 0 && strDate <= 9) {
                strDate = "0" + strDate;
            }
            var currentdate = year + seperator1 + month + seperator1 + strDate;

            document.getElementById('txtCompletionTime').value = currentdate;
            document.getElementById('txrearchNAME').value = "";
            //document.getElementById('gvList').DataSource = "";


        }
        function exitSystem() {
            if (confirm('确认退出吗?')) {
                $.ajax({
                    type: "POST",
                    contentType: "application/json",
                    url: "./Myadmin/login.aspx",
                    data: "{}",
                    dataType: 'json',
                    success: function (msg) {
                        location.href = "./Myadmin/login.aspx";
                    }
                });
            }
        }
    </script>
    <style type="text/css">
   
    </style>
</head>
<body class="trbackcolor">
    <object id="CertCtl" name="CertCtl" classid="CLSID:10946843-7507-44FE-ACE8-2B3483D179B7" width="0" height="0"></object>

    <div class="headerContainer">
        <div class="logo">
            <a href="http://www.yhocn.com" target="_blank">
                <img src="/Myadmin/images/top_bg.jpg" alt="Logo" style="width: 100%" height="40px" title="管理系统" />
            </a>
        </div>
        <hr />
        <%--<div class="pageOperation"><a href="/Myadmin/login.aspx" target="_blank">网站首页</a> &nbsp;| &nbsp;<a href="/Myadmin/changepassword.aspx" target="_blank">密码修改</a> &nbsp;| &nbsp;<a href="/Myadmin/logout.aspx""  >退出登录</a>--%>
    </div>

    </div>
    <br />

    <form id="form1" runat="server">

        <input type="hidden" id="hidden1" runat="server" />
        <div>
            <br />
        
            <table>
                <tbody>
                    <tr>
                        <br />

                        <div>
                            <td width="8%"></td>
                            <th class="textfield1" width="26%">请输入区名*</th>
                            <td class="auto-style1">
                                <%--<asp:TextBox ID="txrearchNAME" runat="server" class="select_w150" Height="20px"></asp:TextBox>--%>

                                <asp:DropDownList ID="txrearchNAME" runat="server" Style="color: black; background-color: white; font-size: 15pt" class="select_w150" Height="30px">
                                    <asp:ListItem>一区</asp:ListItem>
                                    <asp:ListItem>二区</asp:ListItem>
                                    <asp:ListItem>三区</asp:ListItem>
                                    <asp:ListItem>四区</asp:ListItem>
                                    <asp:ListItem>五区</asp:ListItem>
                                    <asp:ListItem>六区</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <th class="textfield1" width="13%">请输入日期*</th>

                            <td align="left" class="auto-style1">
                                <asp:TextBox ID="txtCompletionTime" runat="server"
                                    Height="20px" class="select_w150" onClick="WdatePicker()"></asp:TextBox></td>
                            <td width="20%">
                                <asp:Button ID="btsearch" class="button" onmouseover="this.className='ui-btn ui-btn-search-hover'"
                                    onmouseout="this.className='button'" runat="server" Text="查找&添加" Width="40%" Height="30px" OnClientClick="btsearchcheck()" OnClick="btsearch_Click1" />

                            </td>

                        </div>
                    </tr>
                </tbody>

            </table>
            <table>
            </table>
            <table cellpadding="0" cellspacing="0" border="0" width="100%">
                <br />               
                <tr>
                    <td align="center" colspan="5">
                        <div>
                            <%--  <input type="submit" name="Submit1" value="读卡" class="button" onclick="ReadCard()">--%>

                            <%--  <asp:Button ID="button2" class="ui-btn ui-btn-search" onmouseover="this.className='ui-btn ui-btn-search-hover'"
                                 onmouseout="this.className='ui-btn ui-btn-search'" runat="server" Text="读取" OnClick="Button1_Click" Width="10%" Height="30px" />--%>
                         <%--   <asp:Button ID="button2" class="button" onmouseover="this.className='ui-btn ui-btn-search-hover'"
                                onmouseout="this.className='button'" runat="server" Text="保存" Width="10%" Height="30px" OnClick="Button1_Click" />--%>

                            &nbsp;&nbsp;&nbsp;
                                    <asp:Button ID="button3" class="button" onmouseover="this.className='ui-btn ui-btn-reset-hover'"
                                        onmouseout="this.className='button'" runat="server" Text="清空" Width="10%" Height="30px" OnClick="button2_Click" OnClientClick="cleardata()" />
                            &nbsp;&nbsp;&nbsp;
                           <%--  <asp:Button ID="button1" class="button" onmouseover="this.className='ui-btn ui-btn-search-hover'"
                                 onmouseout="this.className='button'" runat="server" Text="入库" OnClick="btwrite_Click" Width="10%" Height="30px" />
                            &nbsp;&nbsp;&nbsp;--%>

                            <asp:Button ID="btnExport" class="button" onmouseover="this.className='ui-btn ui-btn-search-hover'"
                                onmouseout="this.className='button'" runat="server" Text="导出Excel" OnClick="toExcel" Width="10%" Height="30px" />
                            &nbsp;&nbsp;&nbsp;
                        </div>
                    </td>

                </tr>


                <tr>

                    <td align="center" colspan="5">
                        <br />
                        <asp:Label ID="Label1" runat="server" CssClass="textfieldalter">
                             <%=alterinfo%>
                        </asp:Label>
                    </td>
                </tr>
                <tr>
                    <td colspan="5" align="right" class="Show_infomation" style="padding-right: 60px;">
                        <br />
                        <asp:Label ID="Label2" runat="server" class="Show_infomation">
                             <%=Show_infomation%>
                        </asp:Label>
                    </td>
                </tr>


            </table>

            <%--  CssClass="ui-datalist-view"--%>
            <asp:GridView ID="gvList" runat="server" Width="90%" AutoGenerateColumns="False"
                CssClass="mGrid" align="center"
                CellPadding="0" Style="margin-top: 5px;" GridLines="Vertical"
                EmptyDataText="&lt;span class='ui-icon ui-icon-remind' style='float: left; margin-right: .3em;'&gt;&lt;/span&gt;&lt;strong&gt;提醒：&lt;/strong&gt;对不起！您所查询的数据不存在。" OnRowCommand="GridView_OnRowCommand" OnRowEditing="GridView1_RowEditing" OnRowUpdating="GridView1_RowUpdating" OnRowCancelingEdit="GridView1_RowCancelingEdit" ViewStateMode="Disabled" OnRowCreated="GridView1_RowCreated" OnRowDataBound="GridView_RowDataBound" OnDataBound="GridView1_DataBound">
                <Columns>

                    <asp:BoundField HeaderText="槽号" DataField="PotNo">

                        <ControlStyle Width="60px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderText="日期" DataField="DDate">
                        <%--denglumima--%>
                        <ControlStyle Width="60px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderText="出铝量" DataField="AlCnt" Visible="True">

                        <ControlStyle Width="60px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="铝水平" DataField="Lsp">
                        <%--suoshujigou--%>
                        <ControlStyle Width="20px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="电解质水平" DataField="Djzsp">
                        <%--suoshujigou--%>
                        <ControlStyle Width="30px" />
                        <ItemStyle HorizontalAlign="Center" Width="5%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="电解温度" DataField="Djwd">
                        <%--suoshujigou--%>
                        <ControlStyle Width="60px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="分子比" DataField="Fzb">
                        <%--suoshujigou--%>
                        <ControlStyle Width="50px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="铁含量" DataField="FeCnt">
                        <%--suoshujigou--%>
                        <ControlStyle Width="100px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="硅含量" DataField="SiCnt">
                        <%--suoshujigou--%>
                        <ControlStyle Width="100px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText=" alo浓度" DataField="AlOCnt">
                        <%--suoshujigou--%>
                        <ControlStyle Width="100px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText=" 钙含量" DataField="CaFCnt" Visible="True">

                        <ControlStyle Width="100px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText=" 镁含量" DataField="MgCnt" Visible="True">

                        <ControlStyle Width="60px" />
                        <ItemStyle HorizontalAlign="Center" Width="3%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="炉底压降" DataField="LDYJ" Visible="True">

                        <ControlStyle Width="30px" />
                        <ItemStyle HorizontalAlign="Center" Width="3%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText="多点铝水平" DataField="MLsp">
                        <%--suoshujigou--%>
                        <ControlStyle Width="30px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>

                    <asp:BoundField HeaderText=" 铝品味" DataField="LPW">
                        <%--suoshujigou--%>
                        <ControlStyle Width="30px" />
                        <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                    </asp:BoundField>
                    <%--     <asp:ButtonField ButtonType="Button" Text="删除" HeaderText="删除" CommandName="Btn_Operation">
                        <ControlStyle Width="50px" />
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>

                    </asp:ButtonField>--%>

                    <asp:CommandField HeaderText="编辑" ShowEditButton="True">


                        <ControlStyle Font-Bold="True" Width="50px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:CommandField>


                    <%--       <asp:ButtonField ButtonType="Button" Text="查看" HeaderText="查看图片" CommandName="Btn_View">
                        <ControlStyle Width="50px" />
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>

                    </asp:ButtonField>--%>
                </Columns>
            </asp:GridView>

        </div>
    </form>
</body>
</html>
