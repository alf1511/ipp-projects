<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="PerformanceAWB.aspx.vb" Inherits="ReportWeb.PerformanceAWB" title="Performance AWB" %>
<asp:Content ID="CPH" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
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
				<TD align=center colSpan=3>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="PERFORMANCE AWB" Font-Bold="true"></asp:Label>
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
					<asp:Label id="lblRegionAsal" runat="server" Text="Region (HUB ASAL)"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td3" runat="server" visible="true">
					<asp:DropDownList id="ddlRegionAsal" runat="server"></asp:DropDownList>
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
				<TD>
					<asp:Label id="lblAsal" runat="server" Text="Hub Asal"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td4" runat="server" visible="true">
					<asp:DropDownList id="ddlAsal" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti4" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr4" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll4" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon4" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList4" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal4" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih4" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="Label1" runat="server" Text="Region (HUB TUJUAN)"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td1" runat="server" visible="true">
					<asp:DropDownList id="ddlRegion" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti1" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
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
					<asp:Label id="lblTujuan" runat="server" Text="Hub Tujuan"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td2" runat="server" visible="true">
					<asp:DropDownList id="ddlTujuan" runat="server"></asp:DropDownList>
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
					<asp:Label id="lblKategoriSeller" runat="server" Text="KategoriSeller"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td6" runat="server" visible="true">
					<asp:DropDownList id="ddlKategoriSeller" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti6" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr6" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll6" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon6" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList6" runat="server" RepeatDirection="Horizontal" RepeatColumns="2"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal6" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih6" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
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
					<asp:DropDownList id="ddlSubKategoriSeller" runat="server"></asp:DropDownList>
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
					<asp:Label id="lblEcom" runat="server" Text="Nama e-Commerce"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td30" runat="server" visible="true">
					<asp:DropDownList id="ddlECom" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti30" runat="server" CssClass="button button4" Text="..."></asp:Button>
				</TD>
			</TR>
			<TR id="tr30" runat="server" visible="false">
				<TD colSpan=3>
					<TABLE border="block">
						<TBODY>
							<TR>
								<TD>
									<asp:Button id="btnAll30" runat="server" CssClass="button button4 FloatLeft" Text="Check All"></asp:Button>
									<asp:Button id="btnNon30" runat="server" CssClass="button button4 FloatRight" Text="Un-Check All"></asp:Button>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:CheckBoxList id="chkList30" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Button id="btnBatal30" runat="server" CssClass="button button4 FloatLeft" Text="Batal"></asp:Button>
									<asp:Button id="btnPilih30" runat="server" CssClass="button button4 FloatRight" Text="Pilih"></asp:Button>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="Label3" runat="server" Text="Periode"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD>
					<asp:TextBox id="txtTgl1" runat="server" Width="120px" ReadOnly="true"></asp:TextBox>
					<asp:Label id="Label4" runat="server" Text="s/d"></asp:Label>
					<asp:TextBox id="txtTgl2" runat="server" Width="120px" ReadOnly="true"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD colSpan=2></TD>
				<TD>
					<asp:CheckBox id="ChkAsFinalDate" runat="server" Text="Sebagai Tanggal Final"></asp:CheckBox>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="Label7" runat="server" Text="Jenis Layanan"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td40" runat="server" visible="true">
					<asp:DropDownList id="ddlLayanan" runat="server"></asp:DropDownList>
					<asp:Button id="btnMulti40" runat="server" CssClass="button button4" Text="..." visible="true"></asp:Button>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="Label9" runat="server" Text="Kategori"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td5" runat="server" visible="true">
					<asp:DropDownList id="ddlKategori" runat="server"></asp:DropDownList>
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
				<TD colSpan=3></TD>
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
		</TBODY>
	</TABLE>
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="btnPreview"></asp:PostBackTrigger>
	<asp:PostBackTrigger ControlID="btnProses"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>