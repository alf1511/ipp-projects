<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="ProcessPickingDCReport.aspx.vb" Inherits="ReportWeb.ProcessPickingDCReport" title="Laporan Asuransi Seller" %>
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
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="LAPORAN PROSES DC PICKING" Font-Bold="True"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD align=center colSpan=3>
					<asp:Label id="lblError" runat="server" Text="" Font-Italic="true" ForeColor="red"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblPeriode" runat="server" Text="Periode"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD>
					<asp:TextBox id="txtTgl1" runat="server" ReadOnly="true" Width="120px"></asp:TextBox>
					<asp:Label id="Label4" runat="server" Text="s/d"></asp:Label>
					<asp:TextBox id="txtTgl2" runat="server" ReadOnly="true" Width="120px"></asp:TextBox>
				</TD>
			</TR>
			<TR id="TrHub" runat="server">
				<TD>
					<asp:Label id="LblHub" runat="server" Text="Hub"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="LblHub2" runat="server" Text=":"></asp:Label></TD>
				<TD>
					<asp:DropDownList id="ddlHub" runat="server"></asp:DropDownList>
				</TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblKodeToko" runat="server" Text="Kode Toko"></asp:Label>
				</TD>
				<TD>
					<asp:Label id="lblKodeToko2" runat="server" Text=":"></asp:Label>
				</TD>
				<TD>
					<asp:TextBox id="TxtKodeToko" runat="server" Width="200px" Height="150px" TextMode="MultiLine"></asp:TextBox>
				</TD>
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
	<asp:PostBackTrigger ControlID="BtnProses"></asp:PostBackTrigger>
	<asp:PostBackTrigger ControlID="btnPreview"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>