<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="SettingMstHubExpedition.aspx.vb" Inherits="ReportWeb.SettingMstHubExpedition" title="Setting Daftar Ekspedisi Hub" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<script type="text/javascript">
    function CheckAll(oCheckbox) {
        var gvData = document.getElementById("<%=gvData.ClientID%>");
        var isChecked = 0;
        for(i = 1;i < gvData.rows.length; i++) {
            gvData.rows[i].cells[2].getElementsByTagName("INPUT")[0].checked = oCheckbox.checked;
            if(gvData.rows[i].cells[2].getElementsByTagName("INPUT")[0].checked == true) {
                isChecked += 1;
            }
        }
    }
</script>
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=4>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="PENGATURAN DAFTAR EKSPEDISI UNTUK PROSES HUB" Font-Bold="True"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD style="WIDTH: 130px"></TD>
				<TD></TD>
				<TD style="WIDTH: 200px"></TD>
				<TD></TD>
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
				<TD>
					<asp:DropDownList id="ddlAsal" runat="server" width="200px"></asp:DropDownList>
				</TD>
				<TD></TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblProses" runat="server" Text="Proses"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD>
					<asp:DropDownList id="ddlProses" runat="server"></asp:DropDownList>
				</TD>
				<TD>
					<asp:Button id="btnCari" onclick="btnCari_Click" runat="server" CssClass="button button4 buttonWidth" Text="Cari"></asp:Button>
				</TD>
			</TR>
			<TR id="TrSearch" runat="server" visible="false">
				<TD align=left colSpan=4>
					<asp:Label id="lblSearch" runat="server" Text=""></asp:Label>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
	<TABLE id="TrData" runat="server" visible="false">
		<TBODY>
			<TR>
				<TD style="WIDTH: 130px" vAlign=top>
					<asp:Label id="lblDaftarEkspedisi" runat="server" Text="Daftar Ekspedisi"></asp:Label>
				</TD>
				<TD vAlign=top>:</TD>
				<TD vAlign=top>
					<asp:GridView id="gvData" runat="server" CssClass="TblGridView" AutoGenerateColumns="False">
						<Columns>
							<asp:TemplateField HeaderText="Nama">
								<ItemTemplate>
									<asp:Label ID="lblDetailNama" runat="server" Text='<%#Eval("Alias").ToString%>' />
								</ItemTemplate>
							</asp:TemplateField>
							<asp:TemplateField HeaderText="Account">
								<ItemTemplate>
									<asp:Label ID="lblDetailAccount" runat="server" Text='<%#Eval("Account").ToString%>' />
								</ItemTemplate>
							</asp:TemplateField>
							<asp:TemplateField>
								<HeaderTemplate>
									<div class="ScaleUp">
										<asp:CheckBox ID="CBStatusAll" runat="server" onclick="CheckAll(this)" />
									</div>
								</HeaderTemplate>
								<ItemTemplate>
									<div class="ScaleUp">
										<asp:CheckBox ID="CBStatus" runat="server" Checked='<%#IIf(Eval("Status").ToString() = "0", "False", "True")%>' />
									</div>
								</ItemTemplate>
								<ItemStyle Height="50px" Width="50px" VerticalAlign="Middle" />
								<HeaderStyle Height="50px" Width="50px" VerticalAlign="Middle" />
							</asp:TemplateField>
						</Columns>
						<EmptyDataTemplate>Data tidak ada !</EmptyDataTemplate>
					</asp:GridView>
				</TD>
				<TD vAlign=top>
					<asp:Button id="btnSimpan" onclick="btnSimpan_Click" runat="server" CssClass="button button4 buttonWidth" Text="Simpan"></asp:Button>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
</ContentTemplate>
</asp:UpdatePanel>
</asp:Content>