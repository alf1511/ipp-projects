<%@ Page Language="VB" AutoEventWireup="false" CodeBehind="QueryTools.aspx.vb" Inherits="ReportWeb.QueryTools" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
	<title>QUERY TOOLS - WEB REPORT</title>
	<link rel="stylesheet" href="css/Default.css" />
	<%--
	<script src="js/bootstrap-datetimepicker.min.js" type="text/javascript"></script>
	<script src="js/bootstrap-datetimepicker.id.js" type="text/javascript"></script>
	--%>
	<link href="js/jquery-ui.css" rel="stylesheet" />
	<script src="js/external/jquery/jquery.js" type="text/javascript"></script>
	<script src="js/jquery-ui.js" type="text/javascript"></script>
</head>

<body>
	<form id="form1" runat="server">
		<asp:ScriptManager ID="ScriptManager1" runat="server" />
		<TABLE style="width:100%;">
			<TBODY>
				<TR>
					<TD colspan="4" style="text-align:center;">
						<div class="MiniTitle"> <a href="Home.aspx">HOME</a> /
							<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="QUERY TOOLS" Font-Bold="True"></asp:Label>
						</div>
					</TD>
					<TD></TD>
				</TR>
				<TR>
					<TD colspan="5">
						<asp:Label id="SelectDB" runat="server" Text="Database : "></asp:Label>
						<asp:DropDownList id="ddlDatabase" runat="server"></asp:DropDownList>
					</TD>
				</TR>
				<TR>
					<TD colspan="4">
						<asp:Label id="lblError" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
					</TD>
					<TD></TD>
				</TR>
				<TR>
					<TD style="width:45px;vertical-align:top;text-align:right;">
						<asp:Label id="lblQuery" runat="server" Text="Query"></asp:Label>
					</TD>
					<TD style="width:5px;text-align:center;vertical-align:top;">:</TD>
					<TD style="width:720px;text-align:center;vertical-align:top;">
						<asp:TextBox id="txtQuery" runat="server" Width="720px" Height="200px" TextMode="MultiLine" style="resize:none;"></asp:TextBox>
					</TD>
					<TD style="text-align:left;vertical-align:top;width:90px;">
						<asp:Button id="BtnExecute" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Execute"></asp:Button>
						<br />
						<button id="BtnCopy" class="button button4 buttonWidth buttonTouch" type="button">Copy</button>
						<br />
						<asp:Button id="BtnNewTab" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="New Tab"></asp:Button>
					</TD>
					<TD></TD>
				</TR>
				<tr>
					<td></td>
					<td></td>
					<td style="text-align:left;vertical-align:top;">
						<asp:Button ID="BtnSelect" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Select" UseSubmitBehavior="False" />
						<asp:Button ID="BtnConcat" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Concat" UseSubmitBehavior="False" />
						<asp:Button ID="BtnGroupConcat" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Grp_Conc" UseSubmitBehavior="False" />
						<asp:Button ID="BtnReplace" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="replace" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAddTime" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="addtime" UseSubmitBehavior="False" />
						<asp:Button ID="BtnUpdTime" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="updtime" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnTrackNum" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="tracknum" UseSubmitBehavior="False" />
						<asp:Button ID="BtnOrderNo" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="orderno" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAddInfo" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="addinfo" UseSubmitBehavior="False" />
						<asp:Button ID="BtnCode" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="code" UseSubmitBehavior="False" />
						<asp:Button ID="BtnName" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="name" UseSubmitBehavior="False" />
						<asp:Button ID="BtnLog" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="log" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnFrom" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="From" UseSubmitBehavior="False" />
						<asp:Button ID="BtnInnerJoin" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Inner Join" UseSubmitBehavior="False" />
						<asp:Button ID="BtnLeftJoin" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Left Join" UseSubmitBehavior="False" />
						<asp:Button ID="BtnTransaction" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Transaction" UseSubmitBehavior="False" />
						<asp:Button ID="BtnTracking" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Tracking" UseSubmitBehavior="False" />
						<asp:Button ID="BtnTrcHist" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="TrcHist" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnTrcDlvrInfo" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="TrcDlvrInfo" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAutoOrder" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="AutoOrder" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAutoOrdTrc" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="AutoOrdTrc" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAutoOrdTrcHis" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="AutoOrdTrcHis" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAutoOrdCallback" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="AutoOrdCallback" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnTracelog" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="tracelog" UseSubmitBehavior="False" />
						<asp:Button ID="BtnReqResLog" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="reqreslog" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAccount" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="account" UseSubmitBehavior="False" />
						<asp:Button ID="BtnIdmStore" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="idmstore" UseSubmitBehavior="False" />
						<asp:Button ID="BtnMstLogin" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="mstlogin" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnWhere" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Where" UseSubmitBehavior="False" />
						<asp:Button ID="BtnAnd" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="And" UseSubmitBehavior="False" />
						<asp:Button ID="BtnOr" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Or" UseSubmitBehavior="False" />
						<asp:Button ID="BtnCurdateActInAct" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="curdate act-inact" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnEquals" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="=" UseSubmitBehavior="False" />
						<asp:Button ID="BtnLike" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Like" UseSubmitBehavior="False" />
						<asp:Button ID="BtnParenthesis" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="( )" UseSubmitBehavior="False" />
						<asp:Button ID="BtnDoubleQuote" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text='" "' UseSubmitBehavior="False" />
						<asp:Button ID="BtnCurdate" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text='curdate()' UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnGroupBy" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Group By" UseSubmitBehavior="False" />
						<asp:Button ID="BtnOrderBy" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Order By" UseSubmitBehavior="False" />
						<asp:Button ID="BtnDesc" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Desc" UseSubmitBehavior="False" />
						<asp:Button ID="BtnLimit" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Limit" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnShowTbl" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Show Tbl" UseSubmitBehavior="False" />
						<asp:Button ID="BtnCreateTbl" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Create Tbl" UseSubmitBehavior="False" />
						<asp:Button ID="BtnShowProc" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Show Proc" UseSubmitBehavior="False" />
						<asp:Button ID="BtnCreateProc" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Create Proc" UseSubmitBehavior="False" />
						<asp:Button ID="BtnDBDev" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="DB Dev" UseSubmitBehavior="False" />
						<br />
						<asp:Button ID="BtnTrcBalikanStaInWhoIn" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="StaIn-WhoIn" UseSubmitBehavior="False" />
						<asp:Button ID="BtnConsServiceEdit" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="ConsServiceEdit" UseSubmitBehavior="False" />
					</td>
					<td style="text-align:left;vertical-align:top;">
						<asp:Button ID="BtnClear" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Clear" UseSubmitBehavior="False" />
						<br />						
						<asp:Button id="BtnDownload" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Download"></asp:Button>
						<br />
						<asp:Button ID="BtnShowSlave" runat="server" CssClass="button button4 buttonWidth buttonTouch" Text="Show Slave Status" UseSubmitBehavior="False" />
					</td>
					<td></td>
				</tr>
				<TR>
					<TD style="text-align:right;vertical-align:top;">
						<asp:Label id="lblResult" runat="server" Text="Result"></asp:Label>
					</TD>
					<TD style="text-align:left;vertical-align:top;">:</TD>
					<TD style="text-align:left;vertical-align:top;">
						<asp:TextBox id="txtResult" runat="server" Width="700px" Height="200px" TextMode="MultiLine"></asp:TextBox>
					</TD>
					<TD colspan="2"></TD>
				</TR>
				<TR id="TrRowsCount" runat="server" visible="false">
				    <TD colspan="2"></TD>
				    <TD colspan="3">
				        <asp:Label id="lblCtr" runat="server" Text="... Rows"></asp:Label>
				    </TD>
				</TR>
				<TR id="TrData" runat="server" visible="false">
					<TD style="text-align:right;vertical-align:top;">
						<asp:Label id="lblData" runat="server" Text="Data"></asp:Label>
					</TD>
					<TD style="text-align:left;vertical-align:top;">:</TD>
					<TD colspan="3" style="text-align:left;vertical-align:top;">
						<asp:GridView id="gvData" runat="server" CssClass="TblGridView"></asp:GridView>
					</TD>
				</TR>
			</TBODY>
		</TABLE>
	</form>
	<script type="text/javascript">
		$(document).ready(function(){
		        var copyBtn = document.querySelector("#BtnCopy");

		        copyBtn.addEventListener("click", function(event) {
		            var copyTextarea = document.querySelector('#<%=txtQuery.ClientID%>');
		            copyTextarea.focus();
		            copyTextarea.select();

		            try {
		                var successful = document.execCommand("copy");
		                var msg = successful ? "successful" : "unsuccessful";
		                console.log("Fallback: Copying text command was " + msg);
		            } catch (err) {
		                console.error("Fallback: Oops, unable to copy", err);
		            }
		        });
		    });
    </script>
</body>

</html>