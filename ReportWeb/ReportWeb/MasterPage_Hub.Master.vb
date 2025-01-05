Imports ReportWeb.ClsWebVerHub

Partial Class MasterPage_Hub
    Inherits System.Web.UI.MasterPage

    Dim ObjFungsi As New ClsFungsi

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim Name As String = ObjFungsi.ReadWebConfig("AppName").ToString
        If Name <> "" Then
            Page.Title = Page.Title & " - " & Name
            lblAppName.Text = Name
        Else
            Page.Title = Page.Title & " - WEB HUB"
        End If

        lblAppVer.Text = "v. " & ClsWebVer.AppVersion


        If Session("NameOfUser") = "" Or Session("NIK") = "" Then
            Response.Redirect("Login.aspx", False)
        Else

            lblUserName.Text = Session("NameOfUser") & " - " & Session("CurrentWebName") & " (" & Session("CurrentWebCode") & ")"

            Dim NotAllowedHUB As String = ""
            Try
                NotAllowedHUB = ObjFungsi.ReadWebConfig("NotAllowedHUB", False)
            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

            End Try

            'If NotAllowedHUB.Contains(Session("CurrentWebCode")) Then
            'Session.Clear()
            'Response.Redirect("Login.aspx", False)
            'End If

        End If

    End Sub

    Protected Sub btnLogout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogout.Click

        Session.Clear()
        Response.Redirect("Login.aspx", False)

    End Sub

End Class

