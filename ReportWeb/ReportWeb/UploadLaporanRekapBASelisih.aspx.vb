Imports System.Data
Imports ReportWeb.ClsWebVer
Imports System.IO
Imports System.Collections.Generic
Imports MySql.Data.MySqlClient
Imports Org.BouncyCastle.Crypto.Engines

Partial Class UploadLaporanRekapBASelisih
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

        'Dim AllowedMenu As String = "" & Session("AllowedMenu")
        'If AllowedMenu.Contains(MyLi) Then

        Dim MyLink As String = "Link" & MyPage
        Session("CurrentMenu") = MyLink

        If Not IsPostBack Then

            If Not IsNothing(Session("ResultMessage" & MyPage)) Then
                lblError.Text = Session("ResultMessage" & MyPage)
                Session("ResultMessage" & MyPage) = Nothing
            End If

        End If

        'Else
        'Response.Redirect("NotAuthorized.aspx", False)
        'End If

    End Sub


    Protected Sub BtnDownloadTemplate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDownloadTemplate.Click

        ProsesDownloadTemplate()

    End Sub

    Private Sub ProsesDownloadTemplate()

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim dt As DataTable = Nothing

            Try
                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                SqlQuery = "call `template_laporanBAselisih`()"

                Dim SqlParam As New Dictionary(Of String, String)

                dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                Dim FileName As String = "TemplateLaporanBASelisih_" & AutoNum
                Dim FileExt As String = ".xlsx"
                Dim FileNameExt As String = FileName & FileExt

                Dim PesanError As String = ""

                Dim hasil As String = ObjFungsi.WriteFileExcel(dt, "", FileName & FileExt, PesanError)

                Dim Result() As String = hasil.Split("|")
                If Result(0) = "1" Then

                    Dim FileToDownload As String = "" & Result(1)
                    If FileToDownload <> "" Then

                        Dim fileInfo As FileInfo = New FileInfo(FileToDownload)

                        If fileInfo.Exists Then

                            Session("DownloadPage_FileToDownload") = FileToDownload
                            Session("DownloadPage_DLFileName") = FileName
                            Session("DownloadPage_DLFileExt") = FileExt

                            ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Download_" & AutoNum, "child=window.open('DownloadPage.aspx');", True)

                        Else

                            ObjFungsi.WriteTracelogTxt("File " & FileToDownload & " not exists")
                            lblError.Text = "File " & FileToDownload & " not exists"

                        End If

                    Else
                        ObjFungsi.WriteTracelogTxt("File " & FileName & " fail to create!")
                        lblError.Text = "File " & FileName & " fail to create!"
                    End If
                Else
                    lblError.Text = Result(1)
                End If

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

                lblError.Text = ex.Message

            Finally

                Try
                    MCon.Close()
                Catch ex As Exception
                End Try
                Try
                    MCon.Dispose()
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

                Dim Columns(9) As String
                Columns(0) = "F1"
                Columns(1) = "F2"
                Columns(2) = "F3"
                Columns(3) = "F4"
                Columns(4) = "F5"
                Columns(5) = "F6"
                Columns(6) = "F7"
                Columns(7) = "F8"
                Columns(8) = "F9"
                Columns(9) = "F10"

                Dim DataAll As String = ObjFungsi.ReadFileExcel(UploadedPath, Columns, ErrorMessage)
                If DataAll <> "" Then
                    UpdateData(DataAll, UploadedPath)
                Else
                    lblError.Text = ErrorMessage
                End If

            Else
                lblError.Text = ErrorMessage
            End If

        End If

    End Sub

    Private Sub UpdateData(ByVal DataAll As String, ByVal UploadedPath As String)

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)

        Dim MCon As MySqlConnection = Nothing

        Try
            lblError.Text = ""
            TrData.Visible = False

            Dim SBDataList As New StringBuilder
            SBDataList.Append("NO,TIPE,EKSPEDISI,NAMA,ORI,DST,RESI,TANGGAL,PERIODEINVOICE,KETERANGANBA")
            If DataAll <> "" Then
                SBDataList.Append("|")
                SBDataList.Append(DataAll)
            End If

            Dim dtDataList As DataTable = ObjFungsi.ConvertStringToDatatable(SBDataList.ToString)

            If Not IsNothing(dtDataList) Then
                If dtDataList.Rows.Count > 0 Then

                    MCon = ObjCon.SetConn_Master
                    If MCon.State <> ConnectionState.Open Then
                        MCon.Open()
                    End If

                    'table bantuan
                    Dim TempTable As String = "TempLaporanRekapBASelisih"

                    SqlQuery = "Drop Temporary Table If Exists `" & TempTable & "`"
                    SqlParam = New Dictionary(Of String, String)
                    ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                    SqlQuery = "Create Temporary Table `" & TempTable & "` Like RekapBASelisih"
                    SqlParam = New Dictionary(Of String, String)
                    ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                    SqlQuery = "Alter Table `" & TempTable & "` Drop Primary Key;"
                    SqlParam = New Dictionary(Of String, String)
                    ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                    Dim SB As New StringBuilder
                    Dim ErrorMessage As String = ""
                    Dim CurrRow As Integer = 0
                    Dim MaxRow As Integer = 10

                    Dim MaxDataRow As Integer = dtDataList.Rows.Count - 1
                    Dim Tanggal As String = ""

                    'Insert TempTable
                    For i As Integer = 0 To dtDataList.Rows.Count - 1

                        If CurrRow = 0 Then

                            SB = New StringBuilder
                            SB.Append("Insert Into `" & TempTable & "` (")
                            SB.Append("`Tipe`, `Ekspedisi`, `Nama`, `Ori`, `Dst`, `Resi`, `Tanggal`, `PeriodeInvoice`, `KeteranganBA`")
                            SB.Append(") Values ")

                            SqlParam = New Dictionary(Of String, String)

                        Else
                            SB.Append(",")
                        End If

                        SB.Append("(")
                        SB.Append("@Tipe" & i)
                        SB.Append(",@Ekspedisi" & i)
                        SB.Append(",@Nama" & i)
                        SB.Append(",@Ori" & i)
                        SB.Append(",@Dst" & i)
                        SB.Append(",@Resi" & i)
                        SB.Append(",@Tanggal" & i)
                        SB.Append(",@PeriodeInvoice" & i)
                        SB.Append(",@KeteranganBA" & i)
                        SB.Append(")")

                        If dtDataList.Rows(i).Item("TANGGAL").ToString = "" Then
                            Tanggal = DateTime.Now.ToString("yyyy-MM-dd")
                        Else
                            Tanggal = Format(CDate(dtDataList.Rows(i).Item("TANGGAL")), "yyyy-MM-dd")
                        End If

                        SqlParam.Add("@Tipe" & i, "" & dtDataList.Rows(i).Item("TIPE"))
                        SqlParam.Add("@Ekspedisi" & i, "" & dtDataList.Rows(i).Item("EKSPEDISI"))
                        SqlParam.Add("@Nama" & i, "" & dtDataList.Rows(i).Item("NAMA"))
                        SqlParam.Add("@Ori" & i, "" & dtDataList.Rows(i).Item("ORI"))
                        SqlParam.Add("@Dst" & i, "" & dtDataList.Rows(i).Item("DST"))
                        SqlParam.Add("@Resi" & i, "" & dtDataList.Rows(i).Item("RESI"))
                        SqlParam.Add("@Tanggal" & i, Tanggal)
                        SqlParam.Add("@PeriodeInvoice" & i, "" & dtDataList.Rows(i).Item("PERIODEINVOICE"))
                        SqlParam.Add("@KeteranganBA" & i, "" & dtDataList.Rows(i).Item("KETERANGANBA"))
                        CurrRow += 1

                        If CurrRow >= MaxRow Or i = MaxDataRow Then

                            If ObjSQL.SQLExecuteNonQuery(MCon, SB.ToString, SqlParam, ErrorMessage) Then
                            Else
                                lblError.Text = "</br>" & ErrorMessage
                            End If
                            CurrRow = 0

                        End If

                    Next
                    'Insert TempTable Selesai

                    'panggil stored procedure
                    SqlQuery = "call sp_rekapBAselisih_update(@TempTable,@MyPage,@UploadedPath, @UserId)"

                    SqlParam = New Dictionary(Of String, String)
                    SqlParam.Add("@TempTable", TempTable)
                    SqlParam.Add("@MyPage", MyPage)
                    SqlParam.Add("@UploadedPath", UploadedPath)
                    SqlParam.Add("@UserId", "" & Session("NIK"))

                    Dim dtQuery As DataTable = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

                    If Not IsNothing(dtQuery) Then
                        If dtQuery.Rows.Count > 0 Then

                            gvData.DataSource = dtQuery
                            gvData.DataBind()

                            TrData.Visible = True

                            ViewState("gvData") = dtQuery
                        Else
                            lblError.Text = "Berhasil Insert Data!"
                        End If
                    Else
                        lblError.Text = "Gagal Query"
                    End If

                End If
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            lblError.Text = ex.Message

        Finally
            Try
                MCon.Close()
            Catch ex As Exception
            End Try
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

    End Sub

    Protected Sub FileUploadStore_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileUpload.DataBinding
        lblError.Text = ""
    End Sub

End Class
