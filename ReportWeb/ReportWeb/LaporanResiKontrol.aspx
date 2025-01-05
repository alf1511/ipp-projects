<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="LaporanResiKontrol.aspx.vb" Inherits="ReportWeb.LaporanResiKontrol" title="Laporan Resi Kontrol" %>
<asp:Content ID="CPH" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=3>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="Laporan Nomor Resi Konsol" Font-Bold="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3></TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:Label id="lblError" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblNomorResi" runat="server" Text="Nomor Resi"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td" runat="server" visible="true" style="width: 200px;">
					<asp:TextBox style="resize: none" id="txtNomorResi" runat="server" TextMode="MultiLine" Height="21px" Width="200px"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD style="width: 130px;">
					<asp:Label id="lblHub" runat="server" Text="Pilih Hub"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="lblHub2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD style="width: 200px;">
					<asp:DropDownList id="ddlHub" runat="server" Width="200px"></asp:DropDownList>
				</TD>
				<TD></TD>
			</TR>
				<TD class="Center" colSpan=3>
					<TABLE width="100%">
						<TBODY>
							<TR>
								<TD width="50%">
									<asp:Button id="btnProses" runat="server" CssClass="button button4 buttonWidth" Text="Proses"></asp:Button>
								</TD>
								<TD>
									<asp:Button id="btnPreview" runat="server" CssClass="button button4 buttonWidth" Text="Preview"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
		</TBODY>
	</TABLE>
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="BtnProses"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>