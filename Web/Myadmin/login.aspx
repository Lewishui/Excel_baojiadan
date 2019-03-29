<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="login.aspx.cs" Inherits="Web.login" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

    <link href="inc/logincss.css" rel="stylesheet" type="text/css" />
    <link href="../../Myadmin/css/common.css" rel="stylesheet" type="text/css" />
    <script language="JavaScript">
        function CheckLogin() {
            if (document.MyForm.username.value == "") {
                alert("请输入用户名再提交！");
                document.MyForm.username.focus();
                return false
            }
            if (document.MyForm.password.value == "") {
                alert("请输入密码再提交！");
                document.MyForm.password.focus()
                return false

            }

        }

        //$(document).ready(function () {

        //    $('select').addClass("form-control");

        //});
        //$(document).ready(function () {

        //});

    </script>
    <style type="text/css">
        .auto-style1 {
            width: 227px;
        }

        .select {
            position: absolute;
            width: 297px;
            height: 58px;
            padding: 0 24px 0 8px;
            color: #fff;
            font: 12px/21px arial,sans-serif;
            background: url(select.png) no-repeat; /*span背景图片，其实就是dropdownlist图片*/
            overflow: hidden;
            top: 0px;
            left: 0px;
        }

        .auto-style2 {
            width: 404px;
            height: 53px;
        }
    </style>
</head>

<body>
    <div class="mains">
        <div class="inners">
            <p>
                <span style="font-size: 12pt"  align="left" class="font_gray">建议使用IE浏览器</span>

            </p>
            <div class="lefts">

                <p align="left" class="pname">
                    <b><span style="font-size: 43pt">六钻价格表</span></b>
                </p>
            </div>

            <div class="login">
                <form name="MyForm" id="MyForm" runat="server">
                    <%--<form action="admin_login.asp" method="post" name="MyForm" id="Form1" runat="server" >--%>
                    <div class="dbweizhi">
                        <table id="Table1">
                        </table>
                    </div>
                    <input type="hidden" value="chklogin" name="reaction">

                    <div class="center">
                        <div class="inner">
                            <%--     &nbsp;<p>&nbsp;</p>--%>
                            <%--         <p>&nbsp;</p>
                            <p>&nbsp;</p>
                            <p>&nbsp;</p>
                            <p>&nbsp;</p>
                            <p>&nbsp;</p>
                            <p>&nbsp;</p>--%>
                            <table cellpadding="0" cellspacing="0" id="innnertalbe">
                             
                                <tr>
                                    <td height="60">
                                        <p align="right" style="width: 81px">
                                            <b><span style="font-size: 13pt">用 户 名：</span></b>
                                        </p>
                                    </td>
                                    <td class="auto-style1">
                                        <input name="username" type="text" class="select_w150" id="username" size="16" maxlength="100" value="<%=user%>" /></td>
                                </tr>
                                <tr>
                                    <td height="60">
                                        <p align="right">
                                            <b><span style="font-size: 13pt">密&nbsp; 码：</span></b>
                                        </p>
                                    </td>
                                    <td class="auto-style1">
                                        <input name="password" type="password" class="select_w150" id="password" size="16" maxlength="100" value="<%=pass%>" /></td>
                                </tr>

                                <tr>
                                    <td height="49"></td>
                                    <td class="auto-style1">

                                        <%--<asp:Button  name="image" runat="server" type="submit" class="LoginSub" OnClientClick="CheckLogin()" onclick="CheckLogin()" value=" 登 录 " />--%>
                                        <asp:Button ID="image" runat="server" Text=" 登 录 " class="LoginSub" OnClick="HtmlBtn_Click" OnClientClick="CheckLogin()" />

                                          <br>
                                        <br>

                                        <span style="color: #4cff00;"></span>

                                    </td>

                                </tr>

                                <tr>
                                    <br>
                                    <td height="49"></td>
                               


                                </tr>
                                <tr>
                                    <td align="center" colspan="5">
                                    
                                        <asp:Label ID="Label1" runat="server">
                                       <%=alterinfo1%>
                                        </asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <div class="clearfix"></div>

                    </div>

                </form>
            </div>

        </div>
        <div class="ui-login-footer">
            <p>
                <%--<span class="font_gray">云合未来计算机技术有限公司  © Copyright 2018-2020  技术支持：信息技术中心   联系电话：16619776280</span>--%>

            </p>
        </div>
    </div>

</body>
</html>
