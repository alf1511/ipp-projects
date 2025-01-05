<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="BypassCodDeliman.aspx.vb" Inherits="ReportWeb.BypassCodDeliman" title="Bypass Kode Bayar COD Deliman" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=4>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="BYPASS KODE BAYAR COD DELIMAN" Font-Bold="True"></asp:Label>
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
					<asp:DropDownList id="ddlPaymentType" runat="server" Width="150px"></asp:DropDownList>
				</TD>
				<TD style="WIDTH: 5px">:</TD>
				<TD style="WIDTH: 250px">
					<asp:TextBox id="TxtPaymentCode" runat="server" Width="250px"></asp:TextBox>
				</TD>
				<TD>
					<asp:Button id="BtnCari" onclick="btnCari_Click" runat="server" CssClass="button button4 buttonWidth" Text="Cari" Width="65px"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD colSpan=4></TD>
			</TR>
		</TBODY>
	</TABLE>
    <TABLE id="trData" runat="server" visible="False">
		<TBODY>
			<TR id="TrReason" visible="true">
				<TD style="WIDTH: 150px">
					<asp:Label id="lblReason" runat="server" Text="Alasan"></asp:Label>
				</TD>
				<TD style="WIDTH: 5px">:</TD>
				<TD>
					<asp:DropDownList id="ddlReason" runat="server" Width="150px"></asp:DropDownList>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:GridView id="gvData" runat="server" CssClass="TblGridView" AutoGenerateColumns="False" OnRowCommand="gvData_RowCommand" OnRowCancelingEdit="gvData_RowCancelingEdit">
						<Columns>
							<asp:TemplateField HeaderText="AWB">
								<ItemTemplate>
									<asp:Label id="LblDtlAWB" runat="server" Text='<%#"AWB : " & Eval("TrackNum").ToString() & "<br/>Order No : " & Eval("OrderNo").ToString()%>'></asp:Label>
								</ItemTemplate>
								<ItemStyle HorizontalAlign="Left" VerticalAlign="Top"></ItemStyle>
							</asp:TemplateField>
							<asp:TemplateField HeaderText="Informasi COD">
								<ItemTemplate>
									<asp:Label id="LblDtlCOD" runat="server" Text='<%# Eval("CodPaymentBiller").ToString() & " - " & Eval("CodPaymentCode").ToString()%>'></asp:Label>
									<asp:Label id="LblDtlCODValue" runat="server" Text='<%# Eval("CodValue").ToString() %>'></asp:Label>
									<asp:Button id="BtnDtlBypassCOD" runat="server" CommandArgument='<%# Container.DataItemIndex %>' CommandName="Bypass" CssClass="button button4 buttonWidth" Text="Bypass" Width="95px" visible="true"></asp:Button>
								</ItemTemplate>
							</asp:TemplateField>
						</Columns>
					</asp:GridView>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="BtnCari"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>
