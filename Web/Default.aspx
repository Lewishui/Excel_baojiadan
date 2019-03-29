<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Web._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="../Scripts/Common.js" type="text/javascript"></script>

    <script type="text/javascript">
        //$(document).ready(function () {

        //    //$.Tipmsg.r = null;
        //    //var showmsg = function (msg) {

        //    //        alert("所填信息没有经过验证，请稍后…");
        //    //};          

        //})
        //确定删除
        function ConfirmDel() {
            if (confirm("确定删除吗？")) {
                return true;
            }
            else {
                return false;
            }
        }

    </script>

    <style type="text/css">
        * {
            margin: 0;
            padding: 0;
        }

        .form_input_fill {
        }
        .ui-btn-restore {}
    </style>

</head>
<body style="height: 140px">
    <form id="form1" runat="server">
        <div>
        </div>
        <div id="query_form">
            <table border="1px" width="100%" class="ui-table" cellpadding="0" cellspacing="0"
                border="1px">
                <tr align="center">
                    <td colspan="6">
                        <asp:Button ID="btnSave" runat="server" class="ui-btn ui-btn-save" onmouseover="this.className='ui-btn ui-btn-save-hover'"
                            onmouseout="this.className='ui-btn ui-btn-save'" runat="server" Text="读取" OnClick="btnSave_Click" Height="37px" Width="109px" />&nbsp&nbsp
                                <asp:Button ID="btnCopyPreDocVersion" runat="server" CssClass="ui-btn ui-btn-restore" Text="复制前版"
                                    OnClientClick="ConfirmDel(); return false;" Height="37px" Width="110px" />
                        <asp:Button ID="btnSave0" runat="server" class="ui-btn ui-btn-save" onmouseover="this.className='ui-btn ui-btn-save-hover'"
                            onmouseout="this.className='ui-btn ui-btn-save'" runat="server" Text="生成" OnClick="btnCopyPreDocVersion_Click" Height="37px" Width="109px" />
                    </td>

                </tr>
            </table>
        </div>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="ui-table">
            <tr>
                <th align="right" style="width: 10%;">
                    <asp:Label ID="lblModelName" runat="server" Text="模板名称"></asp:Label>
                </th>
                <td align="left" valign="middle" class="main_form_td_bg2" style="padding-left: 7px; width: 60%;">
                    <span class="font_accentuate">
                        <asp:TextBox ID="txbModelName" runat="server" CssClass="form_input_fill" Height="58px" Width="99%"></asp:TextBox>
                    </span>
                </td>
            </tr>
        </table>
    </form>
    <li><a href='/WebForm2.aspx' ><span>转换功能</span></a></li>
      	


</body>
</html>
