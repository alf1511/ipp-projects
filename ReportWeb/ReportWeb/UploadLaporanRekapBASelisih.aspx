<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="UploadLaporanRekapBASelisih.aspx.vb" Inherits="ReportWeb.UploadLaporanRekapBASelisih" title="UploadLaporanRekapBASelisih" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=3>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="UPLOAD LAPORAN REKAP BA SELISIH" Font-Bold="True"></asp:Label>
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
					<asp:Label id="lblFilePath" runat="server" Text="Upload File"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD>
					<asp:FileUpload id="FileUpload" runat="server"></asp:FileUpload>
				</TD>
			</TR>
			<%--<TR>
				<TD colSpan=2></TD>
				<TD>
					<B>* Catatan :</B>
					<OL type=1>
						<LI>File Excel 2007 - ke atas (.xlsx)</LI>
					</OL>
				</TD>
			</TR>--%>
			<TR>
				<TD colSpan=2></TD>
				<TD>
					<asp:Button id="BtnProses" runat="server" CssClass="button button4 buttonWidth" Text="Proses" Width="65px"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD colSpan=2></TD>
				<TD>
					<asp:Button id="BtnDownloadTemplate" runat="server" CssClass="button button4 buttonWidth" Text="Download Template" Width="165px"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3></TD>
			</TR>

		</TBODY>
	</TABLE>
	<%--<TABLE class="SampleTable">
		<TBODY>
			<TR>
				<TH>No</TH>
				<TH>Cluster</TH>
				<TH>CutOff</TH>
				<TH>StoreCode</TH>
				<TH>StoreName</TH>
				<TH>DcCode</TH>
				<TH>DcName</TH>
			</TR>
			<TR>
				<TD>1</TD>
				<TD>BGR A</TD>
				<TD>12</TD>
				<TD>T019</TD>
				<TD>SHOLEH ISKANDAR KM 9</TD>
				<TD>G113</TD>
				<TD>BGR</TD>
			</TR>
			<TR>
				<TD>2</TD>
				<TD>BGR A</TD>
				<TD>12</TD>
				<TD>T1PJ</TD>
				<TD>GRIYA ALAM SENTUL</TD>
				<TD>G113</TD>
				<TD>BGR</TD>
			</TR>
			<TR>
				<TD>3</TD>
				<TD>BGR A</TD>
				<TD>12</TD>
				<TD>T6B8</TD>
				<TD>LINGKAR UTARA</TD>
				<TD>G113</TD>
				<TD>BGR</TD>
			</TR>
			<TR>
				<TD>4</TD>
				<TD>BGR A</TD>
				<TD>12</TD>
				<TD>T82O</TD>
				<TD>PANDU ADMAWINATA</TD>
				<TD>G113</TD>
				<TD>BGR</TD>
			</TR>
			<TR>
				<TD>5</TD>
				<TD>BGR A</TD>
				<TD>12</TD>
				<TD>T8VM</TD>
				<TD>GUNUNG PANCAR</TD>
				<TD>G113</TD>
				<TD>BGR</TD>
			</TR>
		</TBODY>
	</TABLE>--%>
	<br />
	<TABLE id="TrData" runat="server" visible="false">
		<TBODY>
		    <TR>
		        <TD>Hasil Proses :</TD>
		    </TR>
			<TR>
				<TD>
					<asp:GridView id="gvData" runat="server" CssClass="TblGridView" AutoGenerateColumns="False">
					    <Columns>
					        <asp:TemplateField HeaderText="Keterangan" ShowHeader="False">
						        <ItemTemplate>
							        <asp:Label ID="LblKeterangan" runat="server" Text='<%#Eval("Result").ToString%>' />
						        </ItemTemplate>
					        </asp:TemplateField>
							 <asp:TemplateField HeaderText="Detail" ShowHeader="False">
						        <ItemTemplate>
							        <asp:Label ID="LblDetail" runat="server" Text='<%#Eval("Detail").ToString%>' />
						        </ItemTemplate>
					        </asp:TemplateField>
					        <asp:TemplateField HeaderText="Jumlah" ShowHeader="False">
						        <ItemTemplate>
							        <asp:Label ID="LblJumlah" runat="server" Text='<%#Eval("Qty").ToString%>' />
						        </ItemTemplate>
					        </asp:TemplateField>
					    </Columns>
						<EmptyDataTemplate>Data tidak ada !</EmptyDataTemplate>
					</asp:GridView>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="BtnProses"></asp:PostBackTrigger>
	<asp:PostBackTrigger ControlID="BtnDownloadTemplate"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>