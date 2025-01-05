Imports System.IO
Imports System.Collections.Generic
Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Drawing
Imports Org.BouncyCastle.Asn1.Cmp
Imports MySql.Data.MySqlClient

Public Class LaporanResiKontrol
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
    Dim PageTimeout As Integer = 7200000 'In miliseconds
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("NIK") = "2015548000"
        Session("NameOfUser") = "Ryu"

        ScriptManager1 = ScriptManager.GetCurrent(Me.Page)
        'PageTimeout in miliseconds
        ScriptManager1.AsyncPostBackTimeout = PageTimeout / 1000 'In Seconds

        MyPage = System.Web.VirtualPathUtility.GetFileName(Request.RawUrl).Replace(".aspx", "")
        Dim MyLi As String = "li" & MyPage

        'Dim AllowedMenu As String = "" & Session("AllowedMenu")
        'If AllowedMenu.Contains(MyLi) Then

        Dim MyLink As String = "Link" & MyPage
        Session("CurrentMenu") = MyLink

        If Not IsPostBack Then

            LoadData()

            If Not IsNothing(Session("ResultMessage" & MyPage)) Then
                lblError.Text = Session("ResultMessage" & MyPage)
                Session("ResultMessage" & MyPage) = Nothing
            End If

        End If
        'Else
        'Response.Redirect("NotAuthorized.aspx", False)
        'End If
    End Sub

    Private Sub LoadData()

        Dim dt As DataTable = GetHubList()
        ObjFungsi.ddlBound(ddlHub, dt, "Display", "Value")

        ViewState("HubCode") = ""
        ViewState("Nomor") = ""

    End Sub

    Private Function ValidasiProses(ByRef Nomor As String) As Boolean

        Try
            Nomor = txtNomorResi.Text.Trim.Replace(vbCrLf, "/")
            If TxtNomorResi.Text = "" Then
                lblError.Text = lblNomorResi.Text & " masih kosong!"
                Return False
            End If
            Dim KodeResi() As String = Nomor.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)
            ViewState("HubCode") = ddlHub.SelectedValue

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

    Private Function GetHubList() As DataTable

        Dim dt As DataTable = Nothing

        dt = ObjFungsi.GetHubList(Me.GetType().Name, Session("UserCode"))

        If dt.TableName = "ERROR" Then
            lblError.Text = dt.Rows(0).Item("RESPON")
        End If

        Return dt

    End Function

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        ProsesPreview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        ProsesPreview(True)

    End Sub

    Private Sub ProsesPreview(ByVal PreviewOnly As Boolean)
        Dim NomorResi As String = ""
        Try

            If ValidasiProses(NomorResi) Then

                Dim Hub As String = "" & ViewState("HubCode")

                Proses(Hub, NomorResi, PreviewOnly)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub Proses(ByVal Hub As String, ByVal NomorResi As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim ds As DataSet = Nothing

            Try

                SqlQuery = "call `resikonsol_datainfo`(@NomorResi,@Hub,@PreviewOnly);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@Hub", Hub)
                SqlParam.Add("@NomorResi", NomorResi)
                SqlParam.Add("@PreviewOnly", PreviewOnly)

                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                If PreviewOnly Then

                    Dim HeaderRow As Integer = 1
                    Dim HeaderTitle(HeaderRow) As String
                    Dim HeaderContent(HeaderRow) As String

                    HeaderTitle(0) = "Hub"
                    HeaderTitle(1) = "Nomor Resi"

                    HeaderContent(0) = Hub
                    HeaderContent(1) = NomorResi


                    ds = ObjSQL.SQLInsertIntoDataset(MCon, SqlQuery, SqlParam)

                    Session("TitleReport") = "Laporan Nomor Resi Konsol"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReportDs") = ds

                    'Response.Write("<script language=javascript>child=window.open('Preview.aspx');</script>")
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)

                Else

                    Dim PesanError As String = ""

                    Dim FileName As String = "LaporanResiKonsol_" & AutoNum
                    Dim FileExt As String = ".csv"
                    Dim FileNameExt As String = FileName & FileExt

                    'Dim hasil As String = ObjFungsi.DataSetToFile(ds, "|", False, FileName, FileExt, "BIGDATA", PesanError)
                    Dim hasil As String = ObjFungsi.DataReaderToFile(MCon, SqlQuery, SqlParam, "|", False, FileName, FileExt, "BIGDATA", PesanError)

                    Dim Result() As String = hasil.Split("|")
                    If Result(0) = "1" Then

                        Dim ZipFileName As String = FileName
                        Dim ZipFileExt As String = ".zip"
                        Dim ZipFileNameExt As String = ZipFileName & ZipFileExt

                        Dim FileToZip As String = "" & Result(1)
                        Dim FilesToZip As String() = FileToZip.Split("|")
                        Dim FileToDownload As String = ObjFungsi.ProsesZip(FilesToZip, "CompressedReport", ZipFileNameExt, PesanError)

                        If FileToDownload <> "" Then

                            Dim fileInfo As FileInfo = New FileInfo(FileToDownload)

                            If fileInfo.Exists Then

                                Session("DownloadPage_FileToDownload") = FileToDownload
                                Session("DownloadPage_DLFileName") = ZipFileName
                                Session("DownloadPage_DLFileExt") = ZipFileExt

                                ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Download_" & AutoNum, "child=window.open('DownloadPage.aspx');", True)

                            Else

                                ObjFungsi.WriteTracelogTxt("File " & FileToDownload & " not exists")
                                lblError.Text = "File " & ZipFileNameExt & " not exists"
                                ScriptManager1.SetFocus(lblError)

                            End If

                        Else

                            ObjFungsi.WriteTracelogTxt("File " & FileNameExt & " fail to create!")
                            lblError.Text = "File " & FileNameExt & " fail to create!"
                            ScriptManager1.SetFocus(lblError)

                        End If

                    Else
                        lblError.Text = Result(1)
                        ScriptManager1.SetFocus(lblError)
                    End If

                End If

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

                lblError.Text = ex.Message
                ScriptManager1.SetFocus(lblError)

            Finally

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                Try
                    ds = Nothing
                Catch ex As Exception
                End Try

            End Try

        Else

            Response.Redirect("Login.aspx", False)

            Dim VirtualPath As String = Request.CurrentExecutionFilePath
            Dim FileName As String = System.Web.VirtualPathUtility.GetFileName(VirtualPath)
            Session("RedirectMessage") = FileName & ", User Kosong"

        End If

    End Sub

End Class