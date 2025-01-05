<%@ Page Language="VB" MasterPageFile="~/MasterPage_Hub.master" AutoEventWireup="false" CodeBehind="AWB3PLDetil.aspx.vb" Inherits="ReportWeb.AWB3PLDetil" title="3PL LIST AWB DAN KONS DETIL" Culture="en-US" %>
<asp:Content ID="CPH" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<script type="text/javascript">
function DisableScreen() {
	var x = document.getElementById('<%=PageUpdateProgress.ClientID%>');
	var release = document.getElementById('DisableScreenTimeout').value;
	x.style.display = 'block';
	setTimeout(function() {
		x.style.display = 'none';
	}, release);
}

function tabE(obj, e) {
	var e = (typeof event != 'undefined') ? window.event : e; // IE : Moz
	if(e.keyCode == 13) {
		var ele = document.forms[0].elements;
		for(var i = 0; i < ele.length; i++) {
			var q = (i == ele.length - 1) ? 0 : i + 1; // if last element : if any other
			if(obj == ele[i]) {
				if(ele[q].type.toLowerCase() == 'text') {
					ele[q].focus();
					break;
				} else if(ele[q].type.toLowerCase() != 'submit') {
					ele[q + 1].focus();
					break;
				}
			}
		}
		return false;
	}
}

function ValidatingWeight(MyThis) {
	var MyControl = MyThis;
	var MyControlValue = parseFloat(MyControl.value);
	if(isNaN(MyControlValue)) {
		MyControl.value = "0";
	} else {
		MyControl.value = MyControlValue;
	}
}

function ValidatingInt(MyThis) {
	var MyControl = MyThis;
	var MyControlValue = parseInt(MyControl.value);
	if(isNaN(MyControlValue)) {
		MyControl.value = "0";
	} else {
		MyControl.value = MyControlValue;
	}
}

function ConfirmChangeProses() {
	try {
		document.getElementById("<%=ConfirmChange.ClientID%>").value = "no";
		if(confirm("Cetak Bukti Proses AWB 3PL ?\r\n\r\nOK untuk Cetak dan Proses\r\nCancel untuk Proses saja")) {
			document.getElementById("<%=ConfirmChange.ClientID%>").value = "yes";
			return true;
		}
		return false;
	} catch(err) {
		return false;
	}
}
</script>
<div class="MiniTitle"><a onclick="BackConfirm('AWB3PL','3PL List AWB dan Kons');" href="javascript:void(0);">3PL LIST AWB DAN KONS</a> / DETIL</div>
<asp:UpdatePanel runat="server" ID="UpdatePanel1">
	<ContentTemplate>
		<asp:HiddenField id="ConfirmChange" runat="server"></asp:HiddenField>
		<TABLE class="MainTable">
			<TBODY>
				<TR id="TrOtherExpedition" runat="server" visible="TRUE">
					<TD>
						<asp:Label id="lblOtherExpedition" runat="server" Text="3PL" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblOtherExpedition2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD colSpan=3>
						<asp:DropDownList id="ddlOtherExpedition" runat="server" CssClass="globalddl" OnSelectedIndexChanged="ddlOtherExpedition_SelectedIndexChanged" AutoPostBack="True"></asp:DropDownList>
					</TD>
				</TR>
				<TR id="TrTglAWB3PL" runat="server" visible="false">
					<TD>
						<asp:Label id="lblTglAWB" runat="server" Text="Tgl AWB 3PL" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblTglAWB2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD>
						<asp:TextBox id="TxtTglAWB" runat="server" CssClass="globalinput" ReadOnly="True" autocomplete="off" Width="120px"></asp:TextBox>
					</TD>
					<TD>
						<asp:Button id="BtnPickDate" onclick="BtnPickDate_Click" runat="server" CssClass="ButtonPickDate" UseSubmitBehavior="False"></asp:Button>
					</TD>
					<TD></TD>
				</TR>
				<TR id="trCldr" runat="server" visible="false">
					<TD colSpan=2></TD>
					<TD colSpan=3>
						<asp:Calendar id="CldrTglUpload" runat="server" Width="200px" OnSelectionChanged="CldrTglUpload_SelectionChanged" Visible="False" Height="180px" ForeColor="Black" Font-Size="8pt" Font-Names="Verdana" DayNameFormat="Shortest" CellPadding="4" BorderColor="#999999" BackColor="White">
							<SelectedDayStyle BackColor="#666666" Font-Bold="True" ForeColor="White" />
							<SelectorStyle BackColor="#CCCCCC" />
							<WeekendDayStyle BackColor="#FFFFCC" />
							<TodayDayStyle BackColor="#CCCCCC" ForeColor="Black" />
							<OtherMonthDayStyle ForeColor="Gray" />
							<NextPrevStyle VerticalAlign="Bottom" />
							<DayHeaderStyle BackColor="#CCCCCC" Font-Bold="True" Font-Size="7pt" />
							<TitleStyle BackColor="#999999" BorderColor="Black" Font-Bold="True" />
						</asp:Calendar>
					</TD>
				</TR>
				<TR id="TrAsal" runat="server" visible="false">
					<TD>
						<asp:Label id="lblAsal" runat="server" Text="Asal" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblAsal2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD colSpan=3>
						<asp:DropDownList id="ddlAsal" runat="server" CssClass="globalddl" AutoPostBack="False"></asp:DropDownList>
					</TD>
				</TR>
				<TR id="TrTujuan" runat="server" visible="false">
					<TD>
						<asp:Label id="lblTujuan" runat="server" Text="Tujuan" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblTujuan2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD colSpan=3>
						<asp:DropDownList id="ddlTujuan" runat="server" CssClass="globalddl" Visible="False" AutoPostBack="False"></asp:DropDownList>
						<asp:TextBox id="txtTujuan" runat="server" CssClass="globalinput" Visible="False" ReadOnly="True"></asp:TextBox>
					</TD>
				</TR>
				<TR id="TrBerat" runat="server" visible="false">
					<TD>
						<asp:Label id="lblBerat" runat="server" Text="Berat" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblBerat2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD colSpan=3>
						<asp:TextBox onblur="ValidatingWeight(this);" id="TxtBerat" onfocus="RemoveIfZero(this);" runat="server" Text="0" CssClass="globalinput" Width="120px"></asp:TextBox>
						<asp:Label id="LblBeratUnit" runat="server" Text="KG"></asp:Label>
					</TD>
				</TR>
				<TR id="TrQtyVehicle" runat="server" visible="false">
					<TD>
						<asp:Label id="lblQtyVehicle" runat="server" Text="Jml. Kendaraan" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblQtyVehicle2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD colSpan=3>
						<asp:TextBox onblur="ValidatingInt(this);" id="TxtQtyVehicle" onfocus="RemoveIfZero(this);" runat="server" Text="0" CssClass="globalinput" Width="120px"></asp:TextBox>
					</TD>
				</TR>
				<TR id="TrQtyPickup" runat="server" visible="false">
					<TD>
						<asp:Label id="lblQtyPickup" runat="server" Text="Jml. Pickup" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblQtyPickup2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD colSpan=3>
						<asp:TextBox onblur="ValidatingInt(this);" id="TxtQtyPickup" onfocus="RemoveIfZero(this);" runat="server" Text="0" CssClass="globalinput" Width="120px"></asp:TextBox>
					</TD>
				</TR>
				<TR>
					<TD>
						<asp:Label id="lblOtherExpeditionAWB" runat="server" Text="AWB 3PL" CssClass="globallabel"></asp:Label>
					</TD>
					<TD align=center>
						<asp:Label id="lblOtherExpeditionAWB2" runat="server" Text=":" CssClass="globallabel"></asp:Label>
					</TD>
					<TD colSpan=2>
						<asp:TextBox id="txtOtherExpeditionAWB" runat="server" CssClass="globalinput" AutoCompleteType="Disabled"></asp:TextBox>
					</TD>
					<TD style="PADDING-LEFT: 5px">
						<asp:CheckBox id="CBAutoAWB3PL" runat="server" Text="Auto" CssClass="globalcheckbox" AutoPostBack="True" Visible="False" OnCheckedChanged="CBAutoAWB3PL_CheckedChanged"></asp:CheckBox>
					</TD>
				</TR>
				<TR style="HEIGHT: 1px">
					<TD style="WIDTH: 25%"></TD>
					<TD style="WIDTH: 5px"></TD>
					<TD style="WIDTH: 125px"></TD>
					<TD></TD>
					<TD style="WIDTH: 60px"></TD>
				</TR>
				<TR>
					<TD colSpan=5>
						<DIV class="BoxLblError">
							<asp:UpdateProgress id="PageUpdateProgress" runat="server">
								<ProgressTemplate>
									<div class="globallabel" style="text-align:left;">Loading...</div>
									<div class="graybackground-div" style="display:block;"></div>
								</ProgressTemplate>
							</asp:UpdateProgress>
							<asp:Label id="lblError" runat="server" CssClass="globallabel" ForeColor="Red" Font-Italic="True"></asp:Label>
						</DIV>
					</TD>
				</TR>
				<TR id="TrKonsCtr" runat="server" visible="false">
					<TD align=left colSpan=5>
						<asp:Label id="lblCounterKons" runat="server" CssClass="globallabel"></asp:Label>
					</TD>
				</TR>
				<TR id="TrKons" runat="server" visible="false">
					<TD colSpan=5>
						<asp:GridView id="gvData" runat="server" CssClass="TblGridView" AutoGenerateColumns="False">
							<Columns>
								<asp:TemplateField HeaderText="Detail Kons" ShowHeader="False">
									<ItemTemplate>
										<asp:Label ID="LblDetail" runat="server" Text='<%#"<b>Tanggal : </b>" & Eval("SjDate").ToString & "<br/><b>Kons : </b>" & Eval("ConsNum").ToString & "<br/><b>SJ No : </b>" & Eval("SjNum").ToString & "<br/><b>3PL : </b>" & Eval("OtherExpeditionName").ToString & "<br/><b>Tujuan : </b>" & Eval("DestinationName").ToString & "<br/><b>AWB 3PL : </b>" & Eval("OtherExpeditionAWB").ToString%>' /> </ItemTemplate>
									<ItemStyle HorizontalAlign="Left"></ItemStyle>
								</asp:TemplateField>
								<asp:TemplateField HeaderText="Berat(KG)" ShowHeader="False">
									<ItemTemplate>
										<asp:TextBox ID="txtOtherExpeditionWeight" Text='<%#Eval("OtherExpeditionWeight").ToString%>' runat="server" CssClass="globalinputlink" style="width:100px;" autocomplete="off" onfocus="RemoveIfZero(this);" onblur="ValidatingWeight(this);" onkeydown="return tabE(this,event);" /> </ItemTemplate>
									<ItemStyle Width="100px"></ItemStyle>
								</asp:TemplateField>
							</Columns>
							<EmptyDataTemplate>Data tidak ada !</EmptyDataTemplate>
						</asp:GridView>
					</TD>
				</TR>
				<TR id="TrAWBCtr" runat="server" visible="false">
					<TD align=left colSpan=5>
						<asp:Label id="lblCounterAWB" runat="server" CssClass="globallabel"></asp:Label>
					</TD>
				</TR>
				<TR id="TrAWB" runat="server" visible="false">
					<TD colSpan=5>
						<asp:GridView id="gvDataAWB" runat="server" CssClass="TblGridView" AutoGenerateColumns="False">
							<Columns>
								<asp:TemplateField HeaderText="Detail AWB" ShowHeader="False">
									<ItemTemplate>
										<asp:Label ID="LblDetailAWB" runat="server" Text='<%#"<b>Tanggal : </b>" & Eval("SjDate").ToString & "<br/><b>AWB : </b>" & Eval("TrackNum").ToString & "<br/><b>SJ No : </b>" & Eval("SjNum").ToString & "<br/><b>3PL : </b>" & Eval("OtherExpeditionName").ToString & "<br/><b>Tujuan : </b>" & Eval("DestinationName").ToString & "<br/><b>AWB 3PL : </b>" & Eval("OtherExpeditionAWB").ToString%>' /> </ItemTemplate>
									<ItemStyle HorizontalAlign="Left"></ItemStyle>
								</asp:TemplateField>
								<asp:TemplateField HeaderText="Berat(KG)" ShowHeader="False">
									<ItemTemplate>
										<asp:TextBox ID="txtOtherExpeditionWeightAWB" Text='<%#Eval("OtherExpeditionWeight").ToString()%>' runat="server" CssClass="globalinputlink" style="width:100px;" autocomplete="off" onfocus="RemoveIfZero(this);" onblur="ValidatingWeight(this);" onkeydown="return tabE(this,event);" /> </ItemTemplate>
									<ItemStyle Width="100px"></ItemStyle>
								</asp:TemplateField>
							</Columns>
							<EmptyDataTemplate>Data tidak ada !</EmptyDataTemplate>
						</asp:GridView>
					</TD>
				</TR>
				<TR id="TrPUPCtr" runat="server" visible="false">
					<TD align=left colSpan=5>
						<asp:Label id="lblCounterPUP" runat="server" CssClass="globallabel"></asp:Label>
					</TD>
				</TR>
				<TR id="TrPUP" runat="server" visible="false">
					<TD colSpan=5>
						<asp:GridView id="gvDataPUP" runat="server" CssClass="TblGridView" AutoGenerateColumns="False">
							<Columns>
								<asp:TemplateField HeaderText="Detail Jemput Rekanan" ShowHeader="False">
									<ItemTemplate>
										<asp:Label ID="LblDetailPUP" runat="server" Text='<%#"<b>Tanggal : </b>" & Eval("SjDate").ToString & "<br/><b>S.Tugas : </b>" & Eval("TrackNum").ToString%>' /> </ItemTemplate>
									<ItemStyle HorizontalAlign="Left"></ItemStyle>
								</asp:TemplateField>
								<asp:TemplateField HeaderText="Berat(KG)" ShowHeader="False" Visible="False">
									<ItemTemplate>
										<asp:TextBox ID="txtOtherExpeditionWeightPUP" Text='<%#Eval("OtherExpeditionWeight").ToString()%>' runat="server" CssClass="globalinputlink" style="width:100px;" autocomplete="off" onfocus="RemoveIfZero(this);" onblur="ValidatingWeight(this);" onkeydown="return tabE(this,event);" /> </ItemTemplate>
									<ItemStyle Width="100px"></ItemStyle>
								</asp:TemplateField>
							</Columns>
							<EmptyDataTemplate>Data tidak ada !</EmptyDataTemplate>
						</asp:GridView>
					</TD>
				</TR>
				<TR>
					<TD align=right colSpan=5>
						<asp:Button id="btnSimpan" runat="server" Text="Simpan" CssClass="buttonCek" Width="65px" onclientclick="ConfirmChangeProses();DisableScreen();DisableMe(this);" UseSubmitBehavior="False"></asp:Button>
					</TD>
				</TR>
			</TBODY>
		</TABLE>
	</ContentTemplate>
</asp:UpdatePanel>
</asp:Content>