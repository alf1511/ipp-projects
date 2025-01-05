<%@ Page Language="VB" AutoEventWireup="false" CodeBehind="Preview.aspx.vb" Inherits="ReportWeb.Preview"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>PREVIEW</title>
    <style type="text/css">
        body {
            font-family:verdana,helvetica,arial,sans-serif;
            font-size:0.75em;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table>
        <tr>
            <td>
                <table id="tableTitle" runat="server">
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table id="tableHeader" runat="server">
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Panel ID="PanelContent" runat="server"></asp:Panel>
                <table border="1" style="border-collapse:collapse" id="tableContent" runat="server">
                </table>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
