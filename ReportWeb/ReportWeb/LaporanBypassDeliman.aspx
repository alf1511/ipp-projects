<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="LaporanBypassDeliman.aspx.vb" Inherits="ReportWeb.LaporanBypassDeliman" title="Laporan Bypass Deliman" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=4>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="Laporan Bypass Deliman" Font-Bold="True"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD colSpan=4></TD>
			</TR>
			<TR>
				<TD colSpan=4>
					<asp:Label id="lblError" runat="server" Text="" Font-Italic="true" ForeColor="red"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD style="WIDTH: 150px">
					<asp:Label id="lblBaypassType" runat="server" Text="Bypass Type : "></asp:Label>
				</TD>
				<TD style="WIDTH: 5px">:</TD>
				<TD style="WIDTH: 250px">
					<asp:DropDownList id="ddlBypassType" runat="server" Width="150px"></asp:DropDownList>
				</TD>
			</TR>
			<%--<TR>
				<TD colSpan=3>
					<asp:TextBox style="DISPLAY: inline; WIDTH: 0px; opacity: 0" id="TxtError" runat="server"></asp:TextBox>
					<asp:Label id="Label1" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
				</TD>
			</TR>--%>
			<TR>
				<TD vAlign=top>
					<asp:Label id="lblAwb" runat="server" Text="Order No / AWB"></asp:Label>
					<BR />
					<asp:Label id="lblNoAWBNotes2" runat="server" Text="*maks. 20 AWB" Font-Italic="True"></asp:Label>
					<BR />
					<asp:Label id="lblNoAWBNotes3" runat="server" Text="*Satu AWB per Baris" Font-Italic="True"></asp:Label>
				</TD>
				<TD vAlign=top>:</TD>
				<TD vAlign=top>
					<asp:TextBox style="resize: none" id="TxtAwb" runat="server" Width="200px" TextMode="MultiLine" Height="150px"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD colSpan=4></TD>
			</TR>
			<TR>
				<TD></TD>
				<TD></TD>
				<TD>
					<asp:Button id="btnProses" runat="server" CssClass="button button4 buttonWidth" Text="Proses"></asp:Button>
					<asp:Button id="btnPreview" runat="server" CssClass="button button4 buttonWidth" Text="Preview"></asp:Button>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="BtnProses"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>