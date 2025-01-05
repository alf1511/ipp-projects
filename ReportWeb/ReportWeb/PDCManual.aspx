<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="PDCManual.aspx.vb" Inherits="ReportWeb.PDCManual" title="PDC MANUAL" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=4>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="PDC MANUAL" Font-Bold="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD colSpan=4></TD>
			</TR>
			<TR>
				<TD colSpan=4>
					<asp:Label id="lblError" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD style="width: 130px;">
					<asp:Label id="lblDC" runat="server" Text="Pilih DC"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="lblDC2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD style="width: 200px;">
					<asp:DropDownList id="ddlDC" runat="server" Width="200px"></asp:DropDownList>
				</TD>
				<TD></TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblKodeToko" runat="server" Text="Daftar Kode Toko"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="lblKodeToko2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD>
					<asp:TextBox id="TxtKodeToko" runat="server" Width="200px" Height="150px" TextMode="MultiLine"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="Label1" runat="server" Text="AWB"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="Label2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD>
					<asp:TextBox id="TxtTrackNum" runat="server" Width="200px" Height="150px" TextMode="MultiLine"></asp:TextBox>
				</TD>
				<TD>
					<asp:Button id="BtnCari" runat="server" CssClass="button button4 buttonWidth" Text="Cari" Width="65px"></asp:Button>
				</TD>
			</TR>
		<%--	<TR>
				<TD colSpan=4>
					<asp:Label id="lblNotes" runat="server" Text="* bila lebih dari 1, gunakan tanda koma sebagai pemisah"></asp:Label>
				</TD>
			</TR>--%>
			<TR>
				<TD colSpan=4></TD>
			</TR>
		</TBODY>
	</TABLE>
	<TABLE id="trData" runat="server" visible="false">
		<TBODY>
			<TR>
				<TD>
					<asp:GridView id="gvData" runat="server" CssClass="TblGridView" AutoGenerateColumns="False">
						<Columns>
							<asp:BoundField DataField="TrackNum" HeaderText="AWB" />
							<asp:BoundField DataField="Store" HeaderText="Toko" />
							<asp:BoundField DataField="Zone" HeaderText="Zona" />
						</Columns>
					</asp:GridView>
				</TD>
				<TD vAlign=bottom>
					<asp:Button id="BtnProses" runat="server" CssClass="button button4 buttonWidth" Text="Proses" Width="65px"></asp:Button>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
	<TABLE id="trNoPicking" runat="server" visible="false">
		<TBODY>
			<TR>
				<TD style="width: 130px;">
					<asp:Label id="lblNoPicking" runat="server" Text="Nomor Picking"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="lblNoPicking2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD style="width: 200px;">
					<asp:TextBox id="TxtNoPicking" runat="server" Width="200px" ReadOnly="true"></asp:TextBox>
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