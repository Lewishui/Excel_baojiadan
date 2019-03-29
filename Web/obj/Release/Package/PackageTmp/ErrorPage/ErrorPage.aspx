<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ErrorPage.aspx.cs" Inherits="Web.ErrorPage.ErrorPage" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body style="text-align: center">
    <form id="form1" runat="server">
        <div>
            <image src="/Myadmin/images/ui-bg_error.gif">

            </image>

            <br />
             <br />
            <asp:Label ID="ErrorMessageLabel" runat="server" Text=""></asp:Label>

        </div>
    </form>
</body>
</html>
