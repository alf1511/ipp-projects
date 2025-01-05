<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="MstHubExpeditionReport.aspx.vb" Inherits="ReportWeb.MstHubExpeditionReport" title="Mst Hub Expedition" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=4>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="Mst Hub Expedition Report" Font-Bold="True"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD colSpan=4>
					<asp:Label id="lblError" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblAsal" runat="server" Text="Hub"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td1" runat="server" visible="true">
					<asp:DropDownList id="ddlHub" runat="server" width="200px"></asp:DropDownList>
					<asp:Button id="btnMulti1" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
				<TD></TD>
			</TR>
			<TR id="tr1" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll1" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon1" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList1" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal1" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih1" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblProses" runat="server" Text="Proses"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td2" runat="server" visible="true">
					<asp:DropDownList id="ddlProses" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti2" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr2" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll2" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon2" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList2" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal2" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih2" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblEkspedisi" runat="server" Text="Ekspedisi"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td3" runat="server" visible="true">
					<asp:DropDownList id="ddlEkspedisi" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti3" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr3" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll3" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon3" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList3" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal3" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih3" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
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
			</TR>
			<TR id="TrSearch" runat="server" visible="false">
				<TD align=left colSpan=4>
					<asp:Label id="lblSearch" runat="server" Text=""></asp:Label>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="btnPreview"></asp:PostBackTrigger>
	<asp:PostBackTrigger ControlID="btnProses"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>