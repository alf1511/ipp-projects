Imports System.Data
Imports ReportWeb.ClsWebVer
Imports Microsoft.VisualBasic.ApplicationServices
Imports ReportWeb

Partial Class UbahPass
    Inherits System.Web.UI.Page

    Dim MyPage As String
    Dim ObjFungsi As New ClsFungsi
    Dim ObjCon As New ClsConnection
    Dim ObjSQL As New ClsSQL
    Dim ObjService As New ClsService
    Dim serv As New LocalCore
    Dim param() As Object
    Dim respon() As Object
    Dim ScriptManager1 As New ScriptManager

    Dim PageTimeout As Integer = 3600000 'In miliseconds

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("NIK") = "2015548000"
        Session("NameOfUser") = "Ryu"

        ScriptManager1 = ScriptManager.GetCurrent(Me.Page)
        'PageTimeout in miliseconds
        ScriptManager1.AsyncPostBackTimeout = PageTimeout / 1000 'In Seconds

        MyPage = System.Web.VirtualPathUtility.GetFileName(Request.RawUrl).Replace(".aspx", "")
        Dim MyLi As String = "Link" & MyPage
        Dim AllowedMenu As String = "" & Session("AllowedMenu")
        If AllowedMenu.Contains(MyLi) Then

            Dim MyLink As String = "Link" & MyPage
            Session("CurrentMenu") = MyLink

            If Not IsPostBack Then

                If Not IsNothing(Session("ResultMessage" & MyPage)) Then
                    lblError.Text = Session("ResultMessage" & MyPage)
                    Session("ResultMessage" & MyPage) = Nothing
                End If

            End If
        Else
            Response.Redirect("NotAuthorized.aspx", False)
        End If

    End Sub

    Private Function Validasi() As Boolean

        Dim newpass As String = txtNewPassword.Text
        Dim newpassconf As String = txtKonfirmasi.Text
        Try

            If newpass <> newpassconf Then
                lblError.Text = "Konfirmasi Password tidak sesuai !!"
                Return False
            End If

            If newpass.Length < 6 Then
                lblError.Text = "Password minimal 6 karakter !!"
                Return False
            End If

            If newpass.Contains("'") Or newpass.Contains("\") Then
                lblError.Text = "Password tidak boleh mengandung karakter ' dan \ !!"
                Return False
            End If

            If newpassconf.Contains("'") Or newpassconf.Contains("\") Then
                lblError.Text = "Password tidak boleh mengandung karakter ' dan \ !!"
                Return False
            End If

            If Not (ObjFungsi.isAlphaAndNumeric(newpass) And ObjFungsi.isAlphaAndNumeric(newpassconf)) Then
                lblError.Text = "Password harus kombinasi huruf dan angka !!"
                Return False
            End If

            Return True

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

            Return False

        End Try

    End Function

    Protected Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSimpan.Click

        If Validasi() Then

            If LoginPasswordUpdate() Then

                Session("ResultMessage" & MyPage) = "Berhasil !"
                Response.Redirect(Request.RawUrl, False)

            End If

        End If

    End Sub

#Region "webservice"

    Private Function LoginPasswordUpdate() As Boolean

        Try

            ReDim param(5)
            param(0) = UserWS
            param(1) = PassWS
            param(2) = Session("NIK")
            param(3) = "RPT"
            param(4) = txtOldPassword.Text
            param(5) = txtNewPassword.Text

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")

            Dim i As Integer = 1
            Dim success As Boolean = False
            While i <= maxTryWS And success = False

                Try

                    respon = serv.LoginPasswordUpdateV2(param)
                    success = True

                Catch ex As Exception

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
                    ObjFungsi.WriteTracelogTxt(Pesan)

                    i = i + 1

                End Try

            End While

            If success Then

                lblError.Text = respon(1)

                If respon(0) = "0" Then

                    Return True

                Else

                    Return False

                End If

            Else

                lblError.Text = "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
                ObjFungsi.WriteTracelogTxt(Pesan)

                Return Nothing

            End If

        Catch ex As Exception

            lblError.Text = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function

#End Region

End Class
