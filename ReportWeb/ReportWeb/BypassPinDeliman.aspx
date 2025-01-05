<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="BypassPinDeliman.aspx.vb" Inherits="ReportWeb.BypassPinDeliman" title="Bypass PIN Deliman" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=3>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="BYPASS PIN DELIMAN" Font-Bold="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3></TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:TextBox style="DISPLAY: inline; WIDTH: 0px; opacity: 0" id="TxtError" runat="server"></asp:TextBox>
					<asp:Label id="lblError" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD vAlign=top>
					<asp:Label id="LbAwb" runat="server" Text="Order No / AWB"></asp:Label>
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
				<TD></TD>
				<TD></TD>
				<TD>
					<asp:Button id="btnCari" onclick="btnCari_Click" runat="server" CssClass="button button4 buttonWidth" Text="Cari"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:UpdateProgress id="PageUpdateProgress" runat="server">
						<ProgressTemplate>
							<div class="globallabel" style="text-align:left;">Loading...</div>
						</ProgressTemplate>
					</asp:UpdateProgress>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
	<TABLE id="trData" runat="server" visible="false">
		<TBODY>
			<TR>
				<TD style="WIDTH: 150px">
					<asp:Label id="lblReason" runat="server" Text="Alasan"></asp:Label>
				</TD>
				<TD style="WIDTH: 5px">:</TD>
				<TD>
					<asp:DropDownList id="ddlReason" runat="server" Width="200px"></asp:DropDownList>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:GridView id="gvData" runat="server" CssClass="TblGridView" AutoGenerateColumns="False" OnRowCommand="gvData_RowCommand" OnRowCancelingEdit="gvData_RowCancelingEdit">
						<Columns>
							<asp:TemplateField HeaderText="Detail">
								<ItemTemplate>
									<asp:Label id="LblDtlAWB" runat="server" Text='<%#"AWB : " & Eval("TrackNum").ToString() & "<br/>Order No : " & Eval("OrderNo").ToString()%>'></asp:Label>
								</ItemTemplate>
								<ItemStyle HorizontalAlign="Left" VerticalAlign="Top"></ItemStyle>
							</asp:TemplateField>
							<asp:TemplateField HeaderText="Pin">
								<ItemTemplate>
									<asp:Label id="LblDtlPickup" runat="server" Text='<%#"Pin Ambil - "%>'></asp:Label>
									<asp:Label id="LblDtlPickupSudahSelesai" runat="server" Text='<%#"Transaksi sudah selesai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlPickupTidakAda" runat="server" Text='<%#"Tidak ada PIN Ambil"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINPickupDraft").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlPickupSudahDipakai" runat="server" Text='<%#"Pin sudah dipakai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINPickupDraft").ToString() <> "" And Eval("PINPickup").ToString() <> Eval("PINPickupDraft").ToString(),true,false)%>'></asp:Label>
									<asp:Button id="BtnDtlPickupPakai" runat="server" CommandArgument='<%# Container.DataItemIndex %>' CommandName="Pickup" CssClass="button button4 buttonWidth" Text="Pakai" Width="95px" visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINPickupDraft").ToString() <> "" And Eval("PINPickup").ToString() = Eval("PINPickupDraft").ToString(),true,false)%>'></asp:Button>
									<asp:Label id="LblDtlCancel" runat="server" Text='<%#"<br/>Pin Batal - "%>'></asp:Label>
									<asp:Label id="LblDtlCancelSudahSelesai" runat="server" Text='<%#"Transaksi sudah selesai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlCancelTidakAda" runat="server" Text='<%#"Tidak ada PIN Batal"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINCancelDraft").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlCancelSudahDipakai" runat="server" Text='<%#"Pin sudah dipakai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINCancelDraft").ToString() <> "" And Eval("PINCancel").ToString() <> Eval("PINCancelDraft").ToString(),true,false)%>'></asp:Label>
									<asp:Button id="BtnDtlCancelPakai" runat="server" CommandArgument='<%# Container.DataItemIndex %>' CommandName="Cancel" CssClass="button button4 buttonWidth" Text="Pakai" Width="95px" visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINCancelDraft").ToString() <> "" And Eval("PINCancel").ToString() = Eval("PINCancelDraft").ToString(),true,false)%>'></asp:Button>
									<asp:Label id="LblDtlKeep" runat="server" Text='<%#"<br/>Pin Titip - "%>'></asp:Label>
									<asp:Label id="LblDtlKeepSudahSelesai" runat="server" Text='<%#"Transaksi sudah selesai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlKeepTidakAda" runat="server" Text='<%#"Tidak ada PIN Titip"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINKeepDraft").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlKeepSudahDipakai" runat="server" Text='<%#"Pin sudah dipakai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINKeepDraft").ToString() <> "" And Eval("PINKeep").ToString() <> Eval("PINKeepDraft").ToString(),true,false)%>'></asp:Label>
									<asp:Button id="BtnDtlKeepPakai" runat="server" CommandArgument='<%# Container.DataItemIndex %>' CommandName="Keep" CssClass="button button4 buttonWidth" Text="Pakai" Width="95px" visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINKeepDraft").ToString() <> "" And Eval("PINKeep").ToString() = Eval("PINKeepDraft").ToString(),true,false)%>'></asp:Button>
									<asp:Label id="LblDtlReturn" runat="server" Text='<%#"<br/>Pin Kembali - "%>'></asp:Label>
									<asp:Label id="LblDtlReturnSudahSelesai" runat="server" Text='<%#"Transaksi sudah selesai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlReturnTidakAda" runat="server" Text='<%#"Tidak ada PIN Kembali"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINReturnDraft").ToString() = "",true,false)%>'></asp:Label>
									<asp:Label id="LblDtlReturnSudahDipakai" runat="server" Text='<%#"Pin sudah dipakai"%>' Visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINReturnDraft").ToString() <> "" And Eval("PINReturn").ToString() <> Eval("PINReturnDraft").ToString(),true,false)%>'></asp:Label>
									<asp:Button id="BtnDtlReturnPakai" runat="server" CommandArgument='<%# Container.DataItemIndex %>' CommandName="Return" CssClass="button button4 buttonWidth" Text="Pakai" Width="95px" visible='<%#IIF(Eval("BasePIN").ToString() <> "" And Eval("PINReturnDraft").ToString() <> "" And Eval("PINReturn").ToString() = Eval("PINReturnDraft").ToString(),true,false)%>'></asp:Button>
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
	<asp:PostBackTrigger ControlID="btnCari"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>