<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="SettingLastMileExpedition.aspx.vb" Inherits="ReportWeb.SettingLastMileExpedition" title="Setting Arahan Ekspedisi LastMile" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=3>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="SETTING ARAHAN EKSPEDISI LASTMILE" Font-Bold="True"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3></TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:Label id="lblError" runat="server" Text="" Font-Italic="true" ForeColor="red"></asp:Label>
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
			<TR>
				<TD colSpan=2></TD>
				<TD>
					<B>* Catatan :</B>
					<OL type=1>
						<LI>File Excel 2007 - ke atas (.xlsx)</LI>
						<LI>
							<H2>SEMUA DATA LAMA HARUS IKUT DIUPLOAD</H2>
						</LI>
					</OL>
				</TD>
			</TR>
			<TR>
				<TD colSpan=2></TD>
				<TD>
					<asp:Button id="BtnProses" runat="server" CssClass="button button4 buttonWidth" Text="Proses" Width="65px"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD colSpan=2></TD>
				<TD>
					<asp:Button id="BtnDownload" runat="server" CssClass="button button4 buttonWidth" Text="Download Data Terakhir" Width="165px"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3></TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:Label id="lblContoh" runat="server" Text="Contoh Format File Excel (Kolom Judul tidak harus diberi warna)"></asp:Label>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
	<TABLE class="SampleTable">
		<TBODY>
			<TR>
				<TH>KodePropinsiTujuan</TH>
				<TH>NamaPropinsiTujuan</TH>
				<TH>KodeCurrentHub</TH>
				<TH>NamaCurrentHub</TH>
				<TH>KodeLayanan</TH>
				<TH>NamaLayanan</TH>
				<TH>KodeEkspedisi</TH>
				<TH>NamaEkspedisi</TH>
				<TH>Prioritas</TH>
			</TR>						
			<TR>
			    <TD>1</TD>
				<TD>DKI Jakarta</TD>
				<TD>*</TD>
				<TD>ALL</TD>
				<TD>*</TD>
				<TD>ALL</TD>
				<TD>3000009</TD>
				<TD>JNE</TD>
				<TD>1</TD>
			</TR>
			<TR>
			    <TD>2</TD>
				<TD>Jawa Barat</TD>
				<TD>*</TD>
				<TD>ALL</TD>
				<TD>*</TD>
				<TD>ALL</TD>
				<TD>3000070</TD>
				<TD>POS LASTMILE</TD>
				<TD>2</TD>
			</TR>
			<TR>
			    <TD>2</TD>
				<TD>Jawa Barat</TD>
				<TD>H117</TD>
				<TD>BGR2</TD>
				<TD>*</TD>
				<TD>ALL</TD>
				<TD>3000070</TD>
				<TD>POS LASTMILE</TD>
				<TD>1</TD>
			</TR>
			<TR>
			    <TD>2</TD>
				<TD>Jawa Barat</TD>
				<TD>*</TD>
				<TD>ALL</TD>
				<TD>21</TD>
				<TD>PRIORITY</TD>
				<TD>3000064</TD>
				<TD>JNE YES</TD>
				<TD>2</TD>
			</TR>
			<TR>
			    <TD>2</TD>
				<TD>Jawa Barat</TD>
				<TD>H033</TD>
				<TD>TGR2</TD>
				<TD>9</TD>
				<TD>PICKUP TO CUSTOMER</TD>
				<TD>3000009</TD>
				<TD>JNE</TD>
				<TD>2</TD>
			</TR>
		</TBODY>
	</TABLE>
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="BtnDownload"></asp:PostBackTrigger>
	<asp:PostBackTrigger ControlID="BtnProses"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>
