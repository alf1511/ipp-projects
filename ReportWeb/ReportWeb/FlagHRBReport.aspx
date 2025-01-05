<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="FlagHRBReport.aspx.vb" Inherits="ReportWeb.FlagHRBReport" title="LAPORAN FLAG GAGAL / HRB" %>
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
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="LAPORAN FLAG GAGAL / HRB" Font-Bold="True"></asp:Label>
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
				<TD vAlign=top>
					<asp:Label id="lblNoAWB" runat="server" Text="No AWB"></asp:Label>
					<BR />
					<asp:Label id="lblNoAWBNotes1" runat="server" Text="*boleh dikosongkan" Font-Italic="True"></asp:Label>
					<BR />
					<asp:Label id="lblNoAWBNotes2" runat="server" Text="*maks. 10 AWB" Font-Italic="True"></asp:Label>
					<BR />
					<asp:Label id="lblNoAWBNotes3" runat="server" Text="*Satu AWB per Baris" Font-Italic="True"></asp:Label>
				</TD>
				<TD vAlign=top>:</TD>
				<TD vAlign=top>
					<asp:TextBox id="txtNoAWB" runat="server" Width="280px" Height="150px" TextMode="MultiLine"></asp:TextBox>
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
</ContentTemplate>
<triggers>
	<asp:PostBackTrigger ControlID="btnPreview"></asp:PostBackTrigger>
	<asp:PostBackTrigger ControlID="btnProses"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>

