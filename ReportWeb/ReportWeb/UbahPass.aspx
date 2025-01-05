<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeBehind="UbahPass.aspx.vb" Inherits="ReportWeb.UbahPass" title="UBAH PASSWORD" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<script type="text/javascript">
    function ValidasiKonfirmasi() {
        try {
            var NewPassword = document.getElementById('<%=txtNewPassword.ClientID%>');
            var Konfirmasi = document.getElementById('<%=txtKonfirmasi.ClientID%>');

            if(NewPassword.value == Konfirmasi.value) {

                var string = NewPassword.value

                if(NewPassword.value.trim() == "") {
                    document.getElementById('<%=lblError.ClientID%>').innerHTML = "Password Baru tidak boleh kosong atau hanya spasi !";
                    return false;
                } else if(Konfirmasi.value.trim() == "") {
                    document.getElementById('<%=lblError.ClientID%>').innerHTML = "Konfirmasi tidak boleh kosong atau hanya spasi !";
                    return false;
                } else if(string.indexOf("\\") !== -1) {
                    document.getElementById('<%=lblError.ClientID%>').innerHTML = "Password tidak boleh mengandung karakter \\ !";
                    return false;
                } else if(string.indexOf("'") !== -1) {
                    document.getElementById('<%=lblError.ClientID%>').innerHTML = "Password tidak boleh mengandung karakter ' !";
                    return false;
                }

                return true;
            }
            else {
                document.getElementById('<%=lblError.ClientID%>').innerHTML = "Password Baru dan Konfirmasi tidak sesuai !";
                return false;
            }

        }
        catch(err) {}
    }

    function showHidePass(txtPassID, btnShowHideID) {
        var txtPass = document.getElementById(txtPassID);
        if (txtPass.type === "password") {
            txtPass.type = "text";
            document.getElementById(btnShowHideID).src = "images/eye-slash-solid.svg";
        } else {
            txtPass.type = "password";
            document.getElementById(btnShowHideID).src = "images/eye-solid.svg";
        }
    }
</script>
<asp:UpdatePanel runat="server" ID="UpdatePanelContent">
<ContentTemplate>
	<TABLE>
		<TBODY>
			<TR>
				<TD align=center>
					<asp:Label id="lblJudul" runat="server" CssClass="Judul" Text="UBAH PASSWORD" Font-Bold="True"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD></TD>
			</TR>
			<TR>
				<TD>
					<asp:Label id="lblError" runat="server" Text="" ForeColor="red" Font-Italic="true"></asp:Label>
				</TD>
			</TR>
			<TR>
				<TD>
					<TABLE class="MainTable">
						<TBODY>
							<TR>
								<TD>
									<asp:Label id="LblOldPassword" runat="server" CssClass="globallabel" Text="Password Lama"></asp:Label>
								</TD>
								<TD>
								    <asp:Label id="LblOldPassword2" runat="server" CssClass="globallabel" Text=":"></asp:Label>
								</TD>
								<TD>
									<asp:TextBox id="txtOldPassword" runat="server" CssClass="globalinput" autocomplete="off" TextMode="Password"></asp:TextBox>
									<img id="btnShowHide1" src="images/eye-solid.svg" onclick="showHidePass('<%= txtOldPassword.ClientID %>',this.id)" height="20px" width="20px" style="margin-bottom:-6px;"</img>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Label id="LblNewPassword" runat="server" CssClass="globallabel" Text="Password Baru"></asp:Label>
								</TD>
								<TD>
									<asp:Label id="LblNewPassword2" runat="server" CssClass="globallabel" Text=":"></asp:Label>
								</TD>
								<TD>
									<asp:TextBox id="txtNewPassword" runat="server" CssClass="globalinput" autocomplete="off" TextMode="Password"></asp:TextBox>
									<img id="btnShowHide2" src="images/eye-solid.svg" onclick="showHidePass('<%= txtNewPassword.ClientID %>',this.id)" height="20px" width="20px" style="margin-bottom:-6px;"</img>
								</TD>
							</TR>
							<TR>
								<TD>
									<asp:Label id="LblKonfirmasi" runat="server" CssClass="globallabel" Text="Konfirmasi"></asp:Label>
								</TD>
								<TD>
									<asp:Label id="LblKonfirmasi2" runat="server" CssClass="globallabel" Text=":"></asp:Label>
								</TD>
								<TD>
									<asp:TextBox id="txtKonfirmasi" runat="server" CssClass="globalinput" autocomplete="off" TextMode="Password"></asp:TextBox>
									<img id="btnShowHide3" src="images/eye-solid.svg" onclick="showHidePass('<%= txtKonfirmasi.ClientID %>',this.id)" height="20px" width="20px" style="margin-bottom:-6px;"</img>
								</TD>
							</TR>
							<TR>
								<TD align=right colSpan=3>
									<asp:Button id="btnSimpan" runat="server" CssClass="button button4 buttonWidth" Text=">>>" OnClientClick="return ValidasiKonfirmasi();"></asp:Button>
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
	<asp:PostBackTrigger ControlID="btnSimpan"></asp:PostBackTrigger>
</triggers>
</asp:UpdatePanel>
</asp:Content>