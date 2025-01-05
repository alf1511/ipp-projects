Public Class TestEmail
    Inherits System.Web.UI.Page
    Dim serv As New LocalCore
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
    Protected Sub BtnExecute_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnExecute.Click

        serv.UpdateStatusDeliveryInfo()

    End Sub

End Class