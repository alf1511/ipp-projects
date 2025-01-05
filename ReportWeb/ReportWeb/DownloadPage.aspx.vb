
Imports System.IO

Partial Class DownloadPage
    Inherits System.Web.UI.Page

    Dim ObjFungsi As New ClsFungsi

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim Name As String = ObjFungsi.ReadWebConfig("AppName")
        If Name <> "" Then
            Page.Title = Page.Title & " - " & Name
        Else
            Page.Title = Page.Title & " - WEB REPORT"
        End If

        If Not IsNothing(Session("DownloadPage_FileToDownload")) Then

            Dim FileToDownload As String = "" & Session("DownloadPage_FileToDownload")
            Dim DLFileName As String = "" & Session("DownloadPage_DLFileName")

            Dim DLFileExt As String = "" & Session("DownloadPage_DLFileExt")
            DLFileExt = DLFileExt.Trim
            If DLFileExt <> "" Then
                While DLFileExt.StartsWith(".")
                    DLFileExt = DLFileExt.Substring(1, DLFileExt.Length - 1)
                End While
                DLFileExt = "." & DLFileExt
            End If
            DLFileExt = DLFileExt.ToLower

            Session("DownloadPage_FileToDownload") = Nothing
            Session("DownloadPage_DLFileName") = Nothing
            Session("DownloadPage_DLFileExt") = Nothing

            If FileToDownload <> "" Then

                Dim fileInfo As FileInfo = New FileInfo(FileToDownload)

                If fileInfo.Exists Then

                    Response.Clear()
                    Response.AddHeader("Content-Disposition", "attachment; filename=""" + DLFileName & DLFileExt & """")
                    Response.AddHeader("Content-Length", fileInfo.Length.ToString())

                    If DLFileExt = ".jpg" Or DLFileExt = ".jpeg" Then
                        Response.ContentType = "image/jpeg"
                    ElseIf DLFileExt = ".png" Then
                        Response.ContentType = "image/png"
                    ElseIf DLFileExt = ".pdf" Then
                        Response.ContentType = "application/pdf"
                    Else
                        Response.ContentType = "application/octet-stream"
                    End If

                    Response.TransmitFile(FileToDownload)

                    HttpContext.Current.Response.Flush() 'Sends all currently buffered output to the client.
                    HttpContext.Current.Response.SuppressContent = True 'Gets or sets a value indicating whether to send HTTP content to the client.
                    HttpContext.Current.ApplicationInstance.CompleteRequest() 'Causes ASP.NET to bypass all events and filtering in the HTTP pipeline chain of execution and directly execute the EndRequest event

                End If

            End If

        End If

    End Sub

End Class

