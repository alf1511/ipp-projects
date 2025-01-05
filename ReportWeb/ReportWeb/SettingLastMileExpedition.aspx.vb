Imports System.Data
Imports System.IO
Imports System.Collections.Generic
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient
Imports Microsoft.VisualBasic.ApplicationServices

Partial Class SettingLastMileExpedition
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
        Dim MyLi As String = "li" & MyPage

        Dim AllowedMenu As String = "" & Session("AllowedMenu")
        'If AllowedMenu.Contains(MyLi) Then

        'Dim MyLink As String = "Link" & MyPage
        'Session("CurrentMenu") = MyLink

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

    End Sub

    Private Function ValidasiProses() As Boolean

        lblError.Text = ""

        Try

            If Not FileUpload.HasFile Then
                lblError.Text = "Belum ada file yang dipilih!"
                Return False
            End If

            Return True

        Catch ex As Exception

            lblError.Text = ex.Message
            Return False

        End Try

    End Function

    Protected Sub BtnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnProses.Click

        If ValidasiProses() Then

            Dim ErrorMessage As String = ""
            Dim UploadedPath As String = ""
            If ObjFungsi.UploadFile(UploadedPath, FileUpload, Session("NIK"), MyPage, ErrorMessage) Then

                ErrorMessage = ""

                Dim Columns(4) As String
                Columns(0) = "F1" 'ProvinceCode
                'Columns(1) = "F3" 'ExpeditionCode
                Columns(1) = "F7" 'Expedition
                Columns(2) = "F3" 'DstHub
                Columns(3) = "F5" 'Service
                Columns(4) = "F9" 'Priority

                Dim DataAll As String = ObjFungsi.ReadFileExcel(UploadedPath, Columns, ErrorMessage)
                If DataAll <> "" Then
                    UpdateData(DataAll)
                Else
                    lblError.Text = ErrorMessage
                End If

            Else
                lblError.Text = ErrorMessage
            End If

        End If

    End Sub

    Private Sub UpdateData(ByVal DataAll As String)
        Try

            ReDim param(3)
            param(0) = UserWS
            param(1) = PassWS
            param(2) = DataAll 'Province,Expedition,DstHub,Service|...
            param(3) = "" & Session("NIK") 'NIK

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.UpdateLastMileExpeditionList(AppName, AppVersion, param)

            If respon(0) = "0" Then

                Session("ResultMessage" & MyPage) = "Berhasil !!"
                Response.Redirect(Request.RawUrl, False)

            Else
                lblError.Text = respon(1)
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub FileUpload_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileUpload.DataBinding
        lblError.Text = ""
    End Sub

    Protected Sub BtnDownload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDownload.Click

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim dt As DataTable = Nothing

            Try

                SqlQuery = "call `report_LastMileExpedition`();"

                Dim SqlParam As New Dictionary(Of String, String)

                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If
                dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                Dim DLFileName As String = "LastMileExpedition_" & AutoNum
                Dim DLFileExt As String = ".xlsx"
                Dim DLFileNameExt As String = DLFileName & DLFileExt

                Dim PesanError As String = ""

                Dim Results As String = ObjFungsi.WriteFileExcel(dt, "LastMileExpedition", DLFileNameExt, PesanError)
                Dim Result() As String = Results.Split("|")
                If Result(0) = "1" Then

                    Dim FileToDownload As String = Result(1)

                    If FileToDownload <> "" Then

                        Dim fileInfo As FileInfo = New FileInfo(FileToDownload)

                        If fileInfo.Exists Then

                            Session("DownloadPage_FileToDownload") = FileToDownload
                            Session("DownloadPage_DLFileName") = DLFileName
                            Session("DownloadPage_DLFileExt") = DLFileExt

                            ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Download_" & AutoNum, "child=window.open('DownloadPage.aspx');", True)

                        Else

                            ObjFungsi.WriteTracelogTxt("File " & FileToDownload & " not exists")
                            lblError.Text = "File " & DLFileNameExt & " not exists"

                        End If

                    Else

                        ObjFungsi.WriteTracelogTxt("File " & DLFileNameExt & " fail to create!")
                        lblError.Text = "File " & DLFileNameExt & " fail to create!"

                    End If

                Else
                    lblError.Text = PesanError
                End If

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

                lblError.Text = ex.Message

            Finally

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

            End Try

        Else

            Response.Redirect("Login.aspx", False)

            Dim VirtualPath As String = Request.CurrentExecutionFilePath
            Dim FileName As String = System.Web.VirtualPathUtility.GetFileName(VirtualPath)
            Session("RedirectMessage") = FileName & ", User Kosong"

        End If

    End Sub

End Class
