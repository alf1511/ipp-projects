<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="TindakLanjutSGAReport.aspx.vb" Inherits="ReportWeb.TindakLanjutSGAReport" title="LAPORAN TINDAK LANJUT SGA" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <script type="text/javascript">
        $(function () {
            $( "#<%= txtTgl1.ClientID %>").datepicker({
                dateFormat: "yy-mm-dd",
                maxDate: '0',
                showOn: "button",
                buttonImage: "images/calendar-128_20x20.png",
                buttonImageOnly: true,
                onClose: function (selectedDate) {
                    $("#<%= txtTgl2.ClientID %>").datepicker("option", "minDate", selectedDate);
                    $(".ui-datepicker-trigger").css("vertical-align", "middle");
                }
            });
            $( "#<%= txtTgl2.ClientID %>").datepicker({
                dateFormat: "yy-mm-dd",
                maxDate: '0',
                showOn: "button",
                buttonImage: "images/calendar-128_20x20.png",
                buttonImageOnly: true
            });
            $(".ui-datepicker-trigger").css("vertical-align", "middle");
        });
        Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);
        function EndRequestHandler(sender, args) {
            //Binding Code Again
            $( "#<%= txtTgl1.ClientID %>").datepicker({
                dateFormat: "yy-mm-dd",
                maxDate: '0',
                showOn: "button",
                buttonImage: "images/calendar-128_20x20.png",
                buttonImageOnly: true,
                onClose: function (selectedDate) {
                    $("#<%= txtTgl2.ClientID %>").datepicker("option", "minDate", selectedDate);
                    $(".ui-datepicker-trigger").css("vertical-align", "middle");
                }
            });
            $( "#<%= txtTgl2.ClientID %>").datepicker({
                dateFormat: "yy-mm-dd",
                maxDate: '0',
                showOn: "button",
                buttonImage: "images/calendar-128_20x20.png",
                buttonImageOnly: true
            });
            $(".ui-datepicker-trigger").css("vertical-align", "middle");
        }
    </script>
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=4>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="Laporan Tindak Lanjut SGA" Font-Bold="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<%--<TD style="WIDTH: 140px"></TD>
				<TD></TD>
				<TD style="WIDTH: 200px"></TD>
				<TD></TD>--%>
			</TR>
			<TR>
				<TD colSpan=4>
					<asp:Label id="lblError" runat="server" Text="" Font-Italic="true" ForeColor="red"></asp:Label>
					<asp:TextBox style="WIDTH: 0px; HEIGHT: 0px; opacity: 0" id="TxtError" runat="server"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD vAlign=top>
					<asp:Label id="lblTanggal" runat="server" Text="Periode"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:Label id="lblTanggal2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD colSpan=2>
					<asp:TextBox id="txtTgl1" runat="server" Width="120px" ReadOnly="true"></asp:TextBox> s/d 
					<asp:TextBox id="txtTgl2" runat="server" Width="120px" ReadOnly="true"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="LblTipeSeller" runat="server" Text="Kategori Seller"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td61" runat="server" visible="true">
					<asp:DropDownList id="ddlTipeSeller" runat="server" autopostback="true"></asp:DropDownList>
					<asp:Button id="btnMulti61" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr61" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll61" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon61" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList61" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal61" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih61" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblSubKategoriSeller" runat="server" Text="Sub Kategori Seller"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td7" runat="server" visible="true">
					<asp:DropDownList id="ddlSubKategoriSeller" runat="server" autopostback="true"></asp:DropDownList>
					<asp:Button id="btnMulti7" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr7" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll7" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon7" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList7" runat="server" RepeatDirection="Horizontal" RepeatColumns="2"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal7" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih7" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="LblNamaEcom" runat="server" Text="Nama e-Commerce"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td51" runat="server" visible="true">
					<asp:DropDownList id="ddlECom" runat="server" autopostback="true"></asp:DropDownList>
					<asp:Button id="btnMulti51" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr51" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll51" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon51" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList51" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal51" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih51" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="Label6" runat="server" Text="Jenis Layanan"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td40" runat="server" visible="true">
					<asp:DropDownList id="ddlLayanan" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti40" runat="server" CssClass="button button4" Text="..." visible="true"></asp:Button>
				</TD>
			</TR>
			<TR id="tr40" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll40" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon40" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList40" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal40" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih40" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
			<TR>
				<TD>
					<asp:Label id="lblStatus" runat="server" Text="Status"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="lblStatus2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD id="tdStatus" runat="server" visible="true"  colspan="2">
					<asp:DropDownList id="ddlStatus" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlStatus_SelectedIndexChanged"></asp:DropDownList>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lbLTindakLanjut" runat="server" Text="Tindak Lanjut"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="lbTindakLanjut2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD id="td3" runat="server" visible="true" colspan="2">
					<asp:DropDownList id="ddlTindakLanjut" runat="server" Enabled="false"></asp:DropDownList>
					<asp:Button id="btnMulti3" runat="server" CssClass="button button4" Enabled="false" Text="..."></asp:Button>
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
				<TD vAlign=top>
					<asp:Label id="lblNoTGF" runat="server" Text="No. Resi"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:Label id="lblNoTGF2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD colSpan=2>
					<asp:TextBox style="resize: none" id="txtNoTGF" runat="server" TextMode="MultiLine" Height="112px" Width="225px"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD style="WIDTH: 140px"></TD>
				<TD></TD>
				<TD>
					<asp:Button id="btnProses" runat="server" CssClass="button button4 buttonWidth" Text="Proses"></asp:Button>
					<asp:Button id="btnPreview" runat="server" CssClass="button button4 buttonWidth" Text="Preview"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD colSpan=4>
					<BR />
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
