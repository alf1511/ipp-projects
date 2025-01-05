<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="CalculateHubMustReport.aspx.vb" Inherits="ReportWeb.CalculateHubMustReport" title="LAPORAN TIMBANG UKUR" %>
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
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="Laporan Timbang Ukur Ulang" Font-Bold="true"></asp:Label>
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
					<asp:Label id="LblNamaEcom" runat="server" Text="Partner"></asp:Label>
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
					<asp:Label id="lblHub" runat="server" Text="Hub Asal"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td4" runat="server" visible="true">
					<asp:DropDownList id="ddlHub" runat="server"></asp:DropDownList>
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
					<asp:Label id="lblCabangDC" runat="server" Text="DC Asal"></asp:Label>
				</TD>
				<TD>:</TD>
				<TD id="td2" runat="server" visible="true">
					<asp:DropDownList id="ddlCabangDc" runat="server"></asp:DropDownList>
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
				<TD vAlign=top>
					<asp:Label id="lblKodeToko" runat="server" Text="Toko Asal"></asp:Label>
					<BR />
					<BR />
					<asp:Label id="lblKodeTokoNotes" runat="server" Text="*maks. 20 Toko" Font-Italic="True"></asp:Label>
					<BR />
					<asp:Label id="lblKodeTokoNotes2" runat="server" Text="*Satu Toko per Baris" Font-Italic="True"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:Label id="lblNoteToko" runat="server" Text=":"></asp:Label>
				</TD>
				<TD colSpan=2>
					<asp:TextBox style="resize: none" id="txtKodeToko" runat="server" TextMode="MultiLine" Height="112px" Width="225px"></asp:TextBox>
				</TD>
			</TR>
			<TR>
				<TD vAlign=top>
					<asp:Label id="lblNoAwb" runat="server" Text="No. Resi"></asp:Label>
					<BR />
					<BR />
					<asp:Label id="lblNoAwbNotes" runat="server" Text="*maks. 20 Resi" Font-Italic="True"></asp:Label>
					<BR />
					<asp:Label id="lblNoAwbNotes2" runat="server" Text="*Satu Resi per Baris" Font-Italic="True"></asp:Label>
				</TD>
				<TD vAlign=top>
					<asp:Label id="lblNoteAwb" runat="server" Text=":"></asp:Label>
				</TD>
				<TD colSpan=2>
					<asp:TextBox style="resize: none" id="txtNoAwb" runat="server" TextMode="MultiLine" Height="112px" Width="225px"></asp:TextBox>
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
