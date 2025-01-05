<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="RekapSuratJalan.aspx.vb" Inherits="ReportWeb.RekapSuratJalan" title="REKAP SURAT JALAN" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<script type="text/javascript">
    $(function() {

        $( "#<%= txtTgl1.ClientID %>" ).datepicker({
            dateFormat: "yy-mm-dd",
            maxDate: '0',
            showOn: "button",
            buttonImage: "images/calendar-128_20x20.png",
            buttonImageOnly: true,
            onClose: function(selectedDate) {
                $("#<%= txtTgl2.ClientID %>").datepicker("option", "minDate", selectedDate);
                $(".ui-datepicker-trigger").css("vertical-align","middle");
            }
        });

        $( "#<%= txtTgl2.ClientID %>" ).datepicker({
            dateFormat: "yy-mm-dd",
            maxDate: '0',
            showOn: "button",
            buttonImage: "images/calendar-128_20x20.png",
            buttonImageOnly: true
        });

        $(".ui-datepicker-trigger").css("vertical-align","middle");
    });

    Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);
    function EndRequestHandler(sender, args) {
        //Binding Code Again
        $( "#<%= txtTgl1.ClientID %>" ).datepicker({
            dateFormat: "yy-mm-dd",
            maxDate: '0',
            showOn: "button",
            buttonImage: "images/calendar-128_20x20.png",
            buttonImageOnly: true,
            onClose: function(selectedDate) {
                $("#<%= txtTgl2.ClientID %>").datepicker("option", "minDate", selectedDate);
                $(".ui-datepicker-trigger").css("vertical-align","middle");
            }
        });

        $( "#<%= txtTgl2.ClientID %>" ).datepicker({
            dateFormat: "yy-mm-dd",
            maxDate: '0',
            showOn: "button",
            buttonImage: "images/calendar-128_20x20.png",
            buttonImageOnly: true
        });

        $(".ui-datepicker-trigger").css("vertical-align","middle");
    }
</script>
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center colSpan=3>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="REKAP SURAT JALAN" Font-Bold="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD colSpan=3></TD>
			</TR>
			<TR>
				<TD colSpan=3>
					<asp:TextBox style="width:0px;opacity:0;display:inline;border:none;" id="TxtError" runat="server"></asp:TextBox>
					<asp:Label id="lblError" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
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
				<TD>
					<asp:Label id="lblEkspedisi" runat="server" Text="Ekspedisi 3PL"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td1" runat="server" visible="true">
					<asp:DropDownList id="ddlEkspedisi" runat="server"></asp:DropDownList>
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