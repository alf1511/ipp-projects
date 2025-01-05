<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="FlagHRB.aspx.vb" Inherits="ReportWeb.FlagHRB" title="FLAG HRB" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=4>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="UPDATE STATUS HILANG/RUSAK/BENCANA" Font-Bold="True"></asp:Label>
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
				<TD style="WIDTH: 155px" vAlign=top>
					<asp:Label id="lblNoAWB" runat="server" Text="No AWB"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:Label id="lblNoAWB2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:TextBox id="txtNoAWB" runat="server" TextMode="MultiLine" Height="150px" Width="280px" style="resize:none;"></asp:TextBox>
				</TD>
				<TD vAlign=top>
					<asp:Button id="BtnTambah" runat="server" CssClass="button button4 buttonWidth" Text="Tambah" Width="75px" UseSubmitBehavior="False"></asp:Button>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
	<TABLE id="TrData" runat="server" visible="false">
		<TBODY>
			<TR>
				<TD>
					<asp:Label id="LblTotalBaris" runat="server"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:GridView id="gvData" runat="server" CssClass="TblGridView" AutoGenerateColumns="False" OnRowCommand="gvData_RowCommand">
						<Columns>
							<asp:TemplateField HeaderText="Hapus" ShowHeader="False">
								<ItemTemplate>
									<asp:LinkButton ID="LinkHapus" runat="server" CausesValidation="false" CommandArgument='<%# Container.DataItemIndex %>' CommandName="Hapus" Text="[X]"></asp:LinkButton>
								</ItemTemplate>
								<ItemStyle CssClass="TblGridViewLink" Width="45px"></ItemStyle>
							</asp:TemplateField>
							<asp:BoundField DataField="TrackNum" HeaderText="AWB"></asp:BoundField>
							<asp:BoundField DataField="Informasi" HeaderText="Hasil Cek" HtmlEncode="False"></asp:BoundField>
						</Columns>
						<EmptyDataTemplate>Data tidak ada !</EmptyDataTemplate>
					</asp:GridView>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="LblTotalBerhasil" runat="server"></asp:Label>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
	<TABLE>
		<TBODY>
			<TR>
				<TD style="WIDTH: 155px">
					<asp:Label id="LblPenyebab" runat="server" Text="Status"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="LblPenyebab2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD colSpan=2>
					<asp:DropDownList id="ddlPenyebab" runat="server"></asp:DropDownList>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="LblInformasi" runat="server" Text="Informasi"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="LblInformasi2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD colSpan=2>
					<asp:Button id="BtnCek" runat="server" CssClass="button button4 buttonWidth" Text="Cek Daftar" Width="85px" UseSubmitBehavior="False"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD style="WIDTH: 155px" vAlign=top>
					<asp:Label id="LblKeterangan" runat="server" Text="Keterangan Lengkap"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:Label id="LblKeterangan2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:TextBox id="TxtKeterangan" runat="server" TextMode="MultiLine" Height="150px" Width="280px" style="resize:none;"></asp:TextBox>
				</TD>
				<TD vAlign=top>
					<asp:Button id="BtnProses" runat="server" CssClass="button button4 buttonWidth" Text="Proses" Width="65px"></asp:Button>
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